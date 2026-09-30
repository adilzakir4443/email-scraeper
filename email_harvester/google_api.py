"""
Google API fallback — used ONLY when proxy-based scraping fails AND the
user has opted in with --api-fallback and provided API keys in .env. Never
called opportunistically: every call site checks proxy results came back
empty first, so real API quota is never spent when the free (proxy-based)
path already works.

Three entry points:
  google_places_search   — DISCOVER fallback (Places Text Search + Details)
  google_custom_search   — generic CSE query, returns raw {title,link,snippet}
  find_website_via_cse   — RESOLVE fallback, wraps google_custom_search
"""

import logging
import time
import urllib.parse
from typing import Iterator, Optional

import httpx

logger = logging.getLogger(__name__)

_PLACES_TEXTSEARCH_URL = "https://maps.googleapis.com/maps/api/place/textsearch/json"
_PLACES_DETAILS_URL = "https://maps.googleapis.com/maps/api/place/details/json"
_CSE_URL = "https://www.googleapis.com/customsearch/v1"

# Google requires a short delay before a next_page_token becomes valid.
_NEXT_PAGE_DELAY = 2.0

# Directory/platform domains that are never the business's own site, even if
# they show up as the top CSE result for its name.
_DIRECTORY_DOMAINS = frozenset([
    "yelp.com", "yellowpages.com", "facebook.com", "google.com", "bing.com",
    "tripadvisor.com", "yell.com", "thomsonlocal.com", "instagram.com",
    "linkedin.com", "twitter.com", "mapquest.com", "bbb.org",
])


def _domain_of(url: str) -> str:
    try:
        netloc = urllib.parse.urlparse(url).netloc.lower()
        return netloc[4:] if netloc.startswith("www.") else netloc
    except Exception:
        return ""


# ---------------------------------------------------------------------------
# A. Google Places (DISCOVER fallback)
# ---------------------------------------------------------------------------

def _get_place_details(place_id: str, api_key: str) -> dict:
    """
    One Places Details call per result. Text Search alone doesn't return a
    phone number or website — both come from this call instead.
    """
    params = {
        "place_id": place_id,
        "fields": "name,formatted_phone_number,website",
        "key": api_key,
    }
    try:
        resp = httpx.get(_PLACES_DETAILS_URL, params=params, timeout=15)
        resp.raise_for_status()
        result = resp.json().get("result", {})
        return {
            "phone": result.get("formatted_phone_number"),
            "website": result.get("website"),
        }
    except Exception as exc:
        logger.debug("[GOOGLE_PLACES] Details call failed for place_id=%s: %s", place_id, exc)
        return {"phone": None, "website": None}


def google_places_search(
    niche: str, location: str, api_key: str, max_results: int = 100
) -> Iterator[dict]:
    """
    Yield business dicts from Google Places Text Search, in the same shape
    every DISCOVER scraper yields (business_name, website_url, phone,
    address, category, source) so the caller can upsert them identically
    to any other source. Paginates via next_page_token; one Details call
    per result fills in phone + website (Text Search alone omits both).
    """
    collected = 0
    places_calls = 0
    details_calls = 0
    next_page_token: Optional[str] = None

    while collected < max_results:
        params = {"query": f"{niche} in {location}", "key": api_key}
        if next_page_token:
            params["pagetoken"] = next_page_token

        try:
            resp = httpx.get(_PLACES_TEXTSEARCH_URL, params=params, timeout=15)
            resp.raise_for_status()
            data = resp.json()
        except Exception as exc:
            logger.warning("[GOOGLE_PLACES] Text Search request failed: %s", exc)
            break
        places_calls += 1

        status = data.get("status")
        if status not in ("OK", "ZERO_RESULTS"):
            logger.warning(
                "[GOOGLE_PLACES] API returned status=%s: %s",
                status, data.get("error_message", ""),
            )
            break

        for result in data.get("results", []):
            if collected >= max_results:
                break

            name = result.get("name")
            if not name:
                continue

            place_id = result.get("place_id")
            details = {"phone": None, "website": None}
            if place_id:
                details = _get_place_details(place_id, api_key)
                details_calls += 1

            types = result.get("types") or []
            category = types[0].replace("_", " ").title() if types else None

            yield {
                "business_name": name,
                "website_url": details.get("website"),
                "phone": details.get("phone"),
                "address": result.get("formatted_address"),
                "category": category,
                "source": "google_places",
            }
            collected += 1

        next_page_token = data.get("next_page_token")
        if not next_page_token or collected >= max_results:
            break
        time.sleep(_NEXT_PAGE_DELAY)

    logger.info(
        '[GOOGLE_PLACES] Found %d businesses for "%s" in "%s"', collected, niche, location,
    )
    logger.info(
        "[GOOGLE_PLACES] API quota used: %d Places calls + %d Details calls",
        places_calls, details_calls,
    )


# ---------------------------------------------------------------------------
# B. Google Custom Search Engine (generic)
# ---------------------------------------------------------------------------

def google_custom_search(query: str, api_key: str, cx: str, num: int = 10) -> list[dict]:
    """
    Run a Google Custom Search Engine query. Returns a list of
    {title, link, snippet} dicts — an empty list on any error (rate limit,
    quota exceeded, bad key/cx, network failure). Never raises: every
    caller treats "nothing found" and "the API failed" identically, since
    neither should crash the pipeline.
    """
    params = {"q": query, "key": api_key, "cx": cx, "num": min(max(num, 1), 10)}  # CSE caps num at 10
    try:
        resp = httpx.get(_CSE_URL, params=params, timeout=15)
        if resp.status_code == 429:
            logger.warning("[GOOGLE_CSE] Rate limited (429) for query %r", query)
            return []
        resp.raise_for_status()
        data = resp.json()
    except Exception as exc:
        logger.warning("[GOOGLE_CSE] Search failed for %r: %s", query, exc)
        return []

    if "error" in data:
        logger.warning(
            "[GOOGLE_CSE] API error for %r: %s", query, data["error"].get("message", data["error"]),
        )
        return []

    results = [
        {"title": item.get("title", ""), "link": item.get("link", ""), "snippet": item.get("snippet", "")}
        for item in data.get("items", [])
    ]
    logger.info('[GOOGLE_CSE] Search: "%s" → %d results', query, len(results))
    return results


# ---------------------------------------------------------------------------
# C. Website search (RESOLVE fallback)
# ---------------------------------------------------------------------------

def find_website_via_cse(
    business_name: str, location: str, api_key: str, cx: str
) -> Optional[str]:
    """
    Search via CSE for '"<name>" "<city>" official site' and return the
    first result whose domain isn't a known directory/platform site.
    """
    city = location.split(",")[0].strip() if location else location
    query = f'"{business_name}" "{city}" official site'
    results = google_custom_search(query, api_key, cx)

    for item in results:
        link = item.get("link")
        if not link:
            continue
        domain = _domain_of(link)
        if not domain or any(domain == d or domain.endswith("." + d) for d in _DIRECTORY_DOMAINS):
            continue
        logger.info("[GOOGLE_CSE] Website found for %s: %s", business_name, link)
        return link

    logger.info("[GOOGLE_CSE] No website found for %s", business_name)
    return None
