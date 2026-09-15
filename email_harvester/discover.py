"""
Stage 1 — DISCOVER.

Scrapes Yellow Pages (primary), Bing Local (secondary), and Yelp (tertiary)
for a given niche + location and inserts all found businesses into SQLite.
Resumable: already-inserted businesses (same name + source) are skipped.
"""

import logging
import re
import time
import random
import urllib.parse
from typing import Iterator

import httpx
from bs4 import BeautifulSoup
from tenacity import (
    retry,
    stop_after_attempt,
    wait_exponential,
    retry_if_exception_type,
    before_sleep_log,
)

from .db import get_conn, upsert_business, count_businesses
from .proxy import get_pool

logger = logging.getLogger(__name__)

# Per-request delay range (seconds) — intentionally slow
DELAY_MIN = 3.0
DELAY_MAX = 8.0

USER_AGENTS = [
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/17.4 Safari/605.1.15",
    "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/123.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:125.0) Gecko/20100101 Firefox/125.0",
]


def _random_headers() -> dict[str, str]:
    return {
        "User-Agent": random.choice(USER_AGENTS),
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
        "Accept-Language": "en-US,en;q=0.9",
        "Accept-Encoding": "gzip, deflate, br",
        "DNT": "1",
        "Connection": "keep-alive",
        "Upgrade-Insecure-Requests": "1",
    }


def _politeness_sleep() -> None:
    time.sleep(random.uniform(DELAY_MIN, DELAY_MAX))


def _make_client(proxy: str | None) -> httpx.Client:
    proxies = {"http://": proxy, "https://": proxy} if proxy else None
    return httpx.Client(
        proxies=proxies,
        follow_redirects=True,
        timeout=httpx.Timeout(30.0),
        verify=True,
    )


# ---------------------------------------------------------------------------
# Retry decorator for transient network errors
# ---------------------------------------------------------------------------

def _fetch_with_retry(client: httpx.Client, url: str, attempt: int = 0) -> httpx.Response | None:
    """GET url; return Response or None on unrecoverable failure."""
    pool = get_pool()
    for i in range(4):
        try:
            resp = client.get(url, headers=_random_headers())
            if resp.status_code in (403, 429, 503):
                logger.warning("HTTP %d for %s (attempt %d)", resp.status_code, url, i)
                pool.backoff_sleep(i + 1)
                continue
            return resp
        except (httpx.ConnectError, httpx.TimeoutException, httpx.RemoteProtocolError) as exc:
            logger.warning("Network error fetching %s: %s (attempt %d)", url, exc, i)
            pool.backoff_sleep(i + 1)
    logger.error("Giving up on %s after retries", url)
    return None


# ---------------------------------------------------------------------------
# Yellow Pages
# ---------------------------------------------------------------------------

def _parse_yp_listing(card) -> dict | None:
    """Parse a single YP result card into a business dict."""
    try:
        name_tag = card.select_one("a.business-name span")
        name = name_tag.get_text(strip=True) if name_tag else None
        if not name:
            return None

        website_tag = card.select_one("a.track-visit-website")
        website = website_tag.get("href") if website_tag else None

        phone_tag = card.select_one("div.phones.phone.primary")
        phone = phone_tag.get_text(strip=True) if phone_tag else None

        addr_parts = []
        street = card.select_one("div.street-address")
        locality = card.select_one("div.locality")
        if street:
            addr_parts.append(street.get_text(strip=True))
        if locality:
            addr_parts.append(locality.get_text(strip=True))
        address = ", ".join(addr_parts) or None

        cat_tag = card.select_one("div.categories a")
        category = cat_tag.get_text(strip=True) if cat_tag else None

        return {
            "business_name": name,
            "website_url": website,
            "phone": phone,
            "address": address,
            "category": category,
        }
    except Exception as exc:
        logger.debug("YP parse error: %s", exc)
        return None


def _scrape_yellowpages(niche: str, location: str, max_results: int) -> Iterator[dict]:
    """Yield business dicts from Yellow Pages."""
    pool = get_pool()
    collected = 0
    page = 1

    while collected < max_results:
        query = urllib.parse.quote_plus(niche)
        geo = urllib.parse.quote_plus(location)
        url = f"https://www.yellowpages.com/search?search_terms={query}&geo_location_terms={geo}&page={page}"
        logger.info("[YP] Fetching page %d: %s", page, url)

        proxy = pool.get()
        with _make_client(proxy) as client:
            resp = _fetch_with_retry(client, url)

        if resp is None or resp.status_code != 200:
            logger.warning("[YP] No valid response for page %d, stopping.", page)
            break

        soup = BeautifulSoup(resp.text, "lxml")
        cards = soup.select("div.result div.info")

        if not cards:
            logger.info("[YP] No more results at page %d", page)
            break

        for card in cards:
            biz = _parse_yp_listing(card)
            if biz:
                biz["source"] = "yellowpages"
                yield biz
                collected += 1
                if collected >= max_results:
                    return

        page += 1
        _politeness_sleep()


# ---------------------------------------------------------------------------
# Bing Local
# ---------------------------------------------------------------------------

def _parse_bing_listing(card) -> dict | None:
    try:
        name_tag = card.select_one("div.b_title h2") or card.select_one(".lc_name")
        name = name_tag.get_text(strip=True) if name_tag else None
        if not name:
            return None

        website_tag = card.select_one("a.b_offsite") or card.select_one("a[data-tag='LocalResults.Website']")
        website = website_tag.get("href") if website_tag else None

        phone_tag = card.select_one(".b_phone") or card.select_one(".lc_phone")
        phone = phone_tag.get_text(strip=True) if phone_tag else None

        addr_tag = card.select_one(".b_address") or card.select_one(".lc_address")
        address = addr_tag.get_text(strip=True) if addr_tag else None

        cat_tag = card.select_one(".b_category") or card.select_one(".lc_type")
        category = cat_tag.get_text(strip=True) if cat_tag else None

        return {
            "business_name": name,
            "website_url": website,
            "phone": phone,
            "address": address,
            "category": category,
        }
    except Exception as exc:
        logger.debug("Bing parse error: %s", exc)
        return None


_BING_MAX_PAGES = 5


def _scrape_bing(niche: str, location: str, max_results: int) -> Iterator[dict]:
    """Yield business dicts from Bing Local."""
    pool = get_pool()
    collected = 0
    offset = 0
    pages_fetched = 0

    while collected < max_results and pages_fetched < _BING_MAX_PAGES:
        query = urllib.parse.quote_plus(f"{niche} near {location}")
        url = f"https://www.bing.com/search?q={query}&filters=local_listing%3Atrue&first={offset}"
        logger.info("[Bing] Fetching offset %d: %s", offset, url)

        proxy = pool.get()
        with _make_client(proxy) as client:
            resp = _fetch_with_retry(client, url)

        if resp is None or resp.status_code != 200:
            logger.warning("[Bing] No valid response at offset %d, stopping.", offset)
            break

        soup = BeautifulSoup(resp.text, "lxml")

        # Bing local pack
        cards = (
            soup.select("div.b_localList li")
            or soup.select("div.b_rs_li")
            or soup.select("li.b_algo")
        )

        if not cards:
            logger.info("[Bing] No more local results at offset %d", offset)
            break

        for card in cards:
            biz = _parse_bing_listing(card)
            if biz:
                biz["source"] = "bing"
                yield biz
                collected += 1
                if collected >= max_results:
                    return

        offset += 10
        _politeness_sleep()


# ---------------------------------------------------------------------------
# Yelp
# ---------------------------------------------------------------------------

def _parse_yelp_listing(card) -> dict | None:
    try:
        name_tag = (
            card.select_one("a.css-19v1rkv")
            or card.select_one("span.css-1egxyab")
            or card.select_one('[class*="businessName"]')
        )
        if not name_tag:
            # Fallback: look for h3 or h4
            name_tag = card.find(["h3", "h4"])
        name = name_tag.get_text(strip=True) if name_tag else None
        if not name:
            return None

        # Website link usually not in listing card on Yelp — use biz page URL
        link_tag = card.find("a", href=re.compile(r"^/biz/"))
        biz_url = None
        if link_tag:
            biz_url = "https://www.yelp.com" + link_tag["href"]

        phone_tag = card.select_one("p[class*='phone']") or card.find("p", string=re.compile(r"\(\d{3}\)"))
        phone = phone_tag.get_text(strip=True) if phone_tag else None

        addr_tag = card.select_one("address") or card.select_one("p[class*='address']")
        address = addr_tag.get_text(strip=True) if addr_tag else None

        cat_tag = card.select_one("span[class*='category']")
        category = cat_tag.get_text(strip=True) if cat_tag else None

        return {
            "business_name": name,
            "website_url": biz_url,   # Yelp listing page, not business site
            "phone": phone,
            "address": address,
            "category": category,
        }
    except Exception as exc:
        logger.debug("Yelp parse error: %s", exc)
        return None


def _scrape_yelp(niche: str, location: str, max_results: int) -> Iterator[dict]:
    """Yield business dicts from Yelp."""
    pool = get_pool()
    collected = 0
    start = 0

    while collected < max_results:
        desc = urllib.parse.quote_plus(niche)
        loc = urllib.parse.quote_plus(location)
        url = f"https://www.yelp.com/search?find_desc={desc}&find_loc={loc}&start={start}"
        logger.info("[Yelp] Fetching start=%d: %s", start, url)

        proxy = pool.get()
        with _make_client(proxy) as client:
            resp = _fetch_with_retry(client, url)

        if resp is None or resp.status_code != 200:
            logger.warning("[Yelp] No valid response at start=%d, stopping.", start)
            break

        soup = BeautifulSoup(resp.text, "lxml")

        # Yelp search results container
        cards = soup.select("ul.undefined > li") or soup.select('li[class*="border-color"]')

        if not cards:
            logger.info("[Yelp] No more results at start=%d", start)
            break

        found_any = False
        for card in cards:
            biz = _parse_yelp_listing(card)
            if biz:
                biz["source"] = "yelp"
                yield biz
                collected += 1
                found_any = True
                if collected >= max_results:
                    return

        if not found_any:
            logger.info("[Yelp] No parseable listings at start=%d, stopping.", start)
            break

        start += 10
        _politeness_sleep()


# ---------------------------------------------------------------------------
# Stage entry-point
# ---------------------------------------------------------------------------

def run_discover(db_path: str, niche: str, location: str, max_results: int) -> None:
    """
    Stage 1: scrape directories and populate the businesses table.
    Already-scraped businesses (same name+source) are silently skipped.
    """
    logger.info("=== STAGE 1: DISCOVER  niche=%r  location=%r  max=%d ===", niche, location, max_results)

    per_source = max(10, max_results // 3)  # split budget across sources

    sources: list[tuple[str, Iterator[dict]]] = [
        ("yellowpages", _scrape_yellowpages(niche, location, max_results)),
        ("bing", _scrape_bing(niche, location, per_source)),
        ("yelp", _scrape_yelp(niche, location, per_source)),
    ]

    total_inserted = 0

    for source_name, iterator in sources:
        logger.info("[DISCOVER] Starting source: %s", source_name)
        count = 0
        for biz in iterator:
            with get_conn(db_path) as conn:
                upsert_business(
                    conn,
                    niche=niche,
                    location=location,
                    business_name=biz["business_name"],
                    website_url=biz.get("website_url"),
                    phone=biz.get("phone"),
                    address=biz.get("address"),
                    category=biz.get("category"),
                    source=biz["source"],
                )
            count += 1
            total_inserted += 1
        logger.info("[DISCOVER] %s: inserted/seen %d businesses", source_name, count)

    with get_conn(db_path) as conn:
        counts = count_businesses(conn, niche, location)

    logger.info(
        "[DISCOVER] DONE — total=%d  with_site=%d  site_less=%d",
        counts["total"],
        counts["with_site"],
        counts["total"] - counts["with_site"],
    )
