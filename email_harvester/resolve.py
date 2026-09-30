"""
Stage 2 — RESOLVE.

For each business with a raw website_url:
  - Strip common tracking params (utm_*, fbclid, gclid, etc.)
  - Force https://
  - Follow exactly one redirect to get the canonical URL
  - Mark resolve_status as done/failed/no_site

Businesses without a website_url get one more chance before being marked
no_site: a quick Bing search for "<name>" "<city>" official site, so a
business a directory simply didn't list a site for (common — many small
listings omit it even when the business has one) isn't wrongly written off.
Already-resolved rows (resolve_status != 'pending') are skipped — idempotent.
"""

import logging
import random
import re
import time
import urllib.parse
from typing import Optional

import httpx
from bs4 import BeautifulSoup

from .db import get_conn
from .proxy import get_pool
from .social import _make_client, _fetch, _unwrap_bing_redirect

logger = logging.getLogger(__name__)

# Politeness delay between the "no website" search-fallback lookups
_SEARCH_DELAY_MIN = 2.0
_SEARCH_DELAY_MAX = 3.0

# Directory/platform domains that are never the business's own site, even
# if they show up as the first Bing result for its name.
_DIRECTORY_DOMAINS = frozenset([
    "yelp.com", "yellowpages.com", "facebook.com", "bing.com", "google.com", "tripadvisor.com",
])

# Tracking params to strip
_STRIP_PARAMS: frozenset[str] = frozenset(
    [
        "utm_source", "utm_medium", "utm_campaign", "utm_term", "utm_content",
        "fbclid", "gclid", "msclkid", "yclid", "ref", "referrer",
        "_ga", "_gac", "mc_cid", "mc_eid",
    ]
)

_SCHEME_RE = re.compile(r"^https?://", re.IGNORECASE)


def _strip_tracking(url: str) -> str:
    """Remove known tracking query parameters from url."""
    try:
        parsed = urllib.parse.urlparse(url)
        qs = urllib.parse.parse_qs(parsed.query, keep_blank_values=True)
        clean = {k: v for k, v in qs.items() if k.lower() not in _STRIP_PARAMS}
        new_query = urllib.parse.urlencode(clean, doseq=True)
        cleaned = parsed._replace(query=new_query)
        return urllib.parse.urlunparse(cleaned)
    except Exception:
        return url


def _force_https(url: str) -> str:
    """Upgrade http:// to https://; add https:// if scheme missing."""
    url = url.strip()
    if not _SCHEME_RE.match(url):
        url = "https://" + url
    return re.sub(r"^http://", "https://", url, flags=re.IGNORECASE)


def _resolve_one(url: str) -> Optional[str]:
    """
    Follow at most one redirect and return the final URL.
    Returns None on network/DNS failure.
    """
    pool = get_pool()
    proxy = pool.get()
    proxies = {"http://": proxy, "https://": proxy} if proxy else None

    for attempt in range(3):
        try:
            with httpx.Client(
                proxies=proxies,
                follow_redirects=True,
                max_redirects=5,
                timeout=httpx.Timeout(15.0),
                verify=True,
            ) as client:
                resp = client.head(url, headers={
                    "User-Agent": "Mozilla/5.0 (compatible; EmailHarvester/1.0)",
                })
                # HEAD sometimes returns 405; fall back to GET
                if resp.status_code == 405:
                    resp = client.get(url, headers={
                        "User-Agent": "Mozilla/5.0 (compatible; EmailHarvester/1.0)",
                    })
                final_url = str(resp.url)
                if proxy:
                    pool.report_success(proxy)
                return final_url
        except (httpx.ConnectError, httpx.TimeoutException, httpx.TooManyRedirects) as exc:
            logger.debug("Resolve attempt %d failed for %s: %s", attempt, url, exc)
            if proxy:
                pool.report_failure(proxy)
            pool.backoff_sleep(attempt + 1)
        except Exception as exc:
            logger.warning("Unexpected error resolving %s: %s", url, exc)
            break

    return None


def _domain_of(url: str) -> str:
    try:
        netloc = urllib.parse.urlparse(url).netloc.lower()
        return netloc[4:] if netloc.startswith("www.") else netloc
    except Exception:
        return ""


def _domain_from_cite_text(cite_text: str) -> Optional[str]:
    """Bing shows each organic result's URL as breadcrumb text (e.g. a
    <cite> tag), separate from its href. Extract just the domain portion —
    used as a fallback when the href doesn't decode to a usable URL."""
    text = cite_text.strip()
    for prefix in ("https://", "http://"):
        if text.lower().startswith(prefix):
            text = text[len(prefix):]
            break
    domain = re.split(r"[/›]", text, maxsplit=1)[0].strip()  # › = "›"
    return domain or None


def _search_website_via_bing(business_name: str, location: str) -> Optional[str]:
    """
    Before giving up on a business with no website_url, search Bing for
    '"<name>" "<city>" official site' and check whether a real (non-
    directory) business site shows up in the first 3 organic results, with
    its domain also appearing in that result's title or snippet as a sanity
    check against a coincidental/unrelated match.

    Reuses social.py's Bing-search plumbing: Bing wraps every organic
    result's href in a bing.com/ck/a redirect whose real destination is
    base64-encoded (see social._unwrap_bing_redirect's docstring for the
    verified-live details — this isn't optional, a bare href is never a
    usable destination URL on a Bing results page).
    """
    pool = get_pool()
    proxy = pool.get()

    city = location.split(",")[0].strip() if location else location
    query = urllib.parse.quote_plus(f'"{business_name}" "{city}" official site')
    url = f"https://www.bing.com/search?q={query}"

    with _make_client(proxy) as client:
        html = _fetch(client, url)
    if not html:
        return None

    soup = BeautifulSoup(html, "lxml")
    for result in soup.select("li.b_algo")[:3]:
        link = result.select_one("h2 a[href]")
        cite = result.select_one("cite")

        real_url = _unwrap_bing_redirect(link["href"]) if link else None
        domain = _domain_of(real_url) if real_url and real_url.startswith("http") else None

        if not domain and cite:
            domain = _domain_from_cite_text(cite.get_text())
            if domain:
                real_url = f"https://{domain}"

        if not domain or not real_url:
            continue
        if any(domain == d or domain.endswith("." + d) for d in _DIRECTORY_DOMAINS):
            continue

        title_text = link.get_text(" ", strip=True) if link else ""
        snippet_tag = result.select_one(".b_caption") or result.select_one("p")
        snippet_text = snippet_tag.get_text(" ", strip=True) if snippet_tag else ""
        if domain in f"{title_text} {snippet_text}".lower():
            return real_url

    return None


def run_resolve(
    db_path: str,
    api_fallback: bool = False,
    google_api_key: str | None = None,
    google_cx: str | None = None,
) -> None:
    """
    Stage 2: normalise all pending website URLs.
    Already-resolved rows are skipped automatically.

    If the Bing search-fallback above finds nothing AND api_fallback is set
    AND Google API keys are configured, a Google Custom Search Engine query
    is tried as one more attempt before marking the business no_site. This
    never runs otherwise — it costs real API quota, so it only kicks in
    once the free path has genuinely come up empty.
    """
    logger.info("=== STAGE 2: RESOLVE ===")

    # Businesses with no website_url get one more chance before being
    # written off: a quick Bing search for their official site. Anything
    # found here gets website_url set and falls through to the normal
    # resolve loop below (same as a business that had a URL from the
    # start); anything not found is marked no_site.
    with get_conn(db_path) as conn:
        no_site_candidates = conn.execute(
            "SELECT id, business_name, location FROM businesses "
            "WHERE website_url IS NULL AND resolve_status='pending'"
        ).fetchall()

    if no_site_candidates:
        logger.info(
            "[RESOLVE] %d business(es) with no website_url — searching for official site",
            len(no_site_candidates),
        )

    found_via_search = 0
    for row in no_site_candidates:
        biz_id = row["id"]
        name = row["business_name"]
        location = row["location"]

        found_url = _search_website_via_bing(name, location)

        if found_url is None and api_fallback and google_api_key and google_cx:
            from .google_api import find_website_via_cse
            found_url = find_website_via_cse(name, location, google_api_key, google_cx)

        with get_conn(db_path) as conn:
            if found_url:
                conn.execute("UPDATE businesses SET website_url=? WHERE id=?", (found_url, biz_id))
                logger.info("[RESOLVE] Found website via search for %s: %s", name, found_url)
                found_via_search += 1
            else:
                conn.execute("UPDATE businesses SET resolve_status='no_site' WHERE id=?", (biz_id,))
                logger.info("[RESOLVE] No website found for %s after search", name)

        time.sleep(random.uniform(_SEARCH_DELAY_MIN, _SEARCH_DELAY_MAX))

    if no_site_candidates:
        logger.info(
            "[RESOLVE] Website search — found=%d  confirmed no_site=%d",
            found_via_search, len(no_site_candidates) - found_via_search,
        )

    with get_conn(db_path) as conn:
        rows = conn.execute(
            "SELECT id, website_url FROM businesses WHERE resolve_status='pending'"
        ).fetchall()

    total = len(rows)
    logger.info("[RESOLVE] %d URLs to process", total)

    done = failed = 0

    for row in rows:
        biz_id: int = row["id"]
        raw_url: str = row["website_url"]

        cleaned = _strip_tracking(_force_https(raw_url))
        final = _resolve_one(cleaned)

        if final:
            with get_conn(db_path) as conn:
                conn.execute(
                    "UPDATE businesses SET normalized_url=?, resolve_status='done' WHERE id=?",
                    (final, biz_id),
                )
            done += 1
        else:
            with get_conn(db_path) as conn:
                conn.execute(
                    "UPDATE businesses SET normalized_url=?, resolve_status='failed' WHERE id=?",
                    (cleaned, biz_id),  # keep cleaned URL even if resolve failed
                )
            failed += 1
            logger.debug("[RESOLVE] Failed to resolve biz_id=%d url=%s", biz_id, raw_url)

    logger.info("[RESOLVE] DONE — resolved=%d  failed=%d  total=%d", done, failed, total)
