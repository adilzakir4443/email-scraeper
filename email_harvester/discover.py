"""
Stage 1 — DISCOVER.

Scrapes business directories for a given niche + location and inserts all
found businesses into SQLite. Which directories run depends on the
detected region for --location (see _detect_region): US sources are Yellow
Pages, Bing Maps, Yelp, and Google Maps; UK sources are Yell.com, Thomson
Local, and Google Maps; Canada sources are YellowPages.ca and Google Maps;
an unrecognized location runs every source.
Resumable: already-inserted businesses (same name + source) are skipped.
"""

import json
import logging
import os
import re
import time
import random
import urllib.parse
from typing import Callable, Iterator

import httpx
from bs4 import BeautifulSoup

from .db import get_conn, upsert_business, count_businesses
from .proxy import get_pool
from .browser import (
    BrowserSession, human_delay, human_scroll, run_in_thread, safe_goto, parse_proxy_for_playwright,
)

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
    # Manually attach Proxy-Authorization too — httpx/httpcore normally
    # derive it from the proxy URL's own credentials, but this is cheap
    # insurance against the specific httpx versions/edge cases that drop it
    # on an HTTPS CONNECT tunnel.
    auth_header = get_pool().get_auth_header(proxy)
    return httpx.Client(
        proxies=proxies,
        follow_redirects=True,
        timeout=httpx.Timeout(30.0),
        verify=True,
        headers=auth_header or None,
    )


# ---------------------------------------------------------------------------
# Retry decorator for transient network errors
# ---------------------------------------------------------------------------

def _fetch_with_retry(
    client: httpx.Client, url: str, attempt: int = 0, proxy: str | None = None
) -> httpx.Response | None:
    """GET url; return Response or None on unrecoverable failure."""
    pool = get_pool()
    for i in range(4):
        try:
            resp = client.get(url, headers=_random_headers())
            if resp.status_code == 407:
                pool.mark_407(proxy)
                logger.warning("407 Proxy Auth failed for %s — check proxy credentials", url)
                break  # don't retry 407 — it won't fix itself with the same proxy
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
            resp = _fetch_with_retry(client, url, proxy=proxy)

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
    """
    Parse a single Bing Maps result card.

    Bing Maps (bing.com/maps) renders its local-business list client-side and
    embeds each listing's full structured data as JSON in a `data-entity`
    attribute on `div.b_maglistcard` — no fragile sub-selectors needed.
    (Bing's classic web-search local pack — div.b_localList / b_title / etc.
    — no longer returns any structured local-business data at all; it now
    just serves organic web results, which is why those selectors always
    matched zero real listings.)
    """
    try:
        raw = card.get("data-entity")
        if not raw:
            return None
        entity = json.loads(raw).get("entity", {})

        name = entity.get("title")
        if not name:
            return None

        return {
            "business_name": name,
            "website_url": entity.get("website"),
            "phone": entity.get("phone"),
            "address": entity.get("address"),
            "category": entity.get("primaryCategoryName"),
        }
    except Exception as exc:
        logger.debug("Bing parse error: %s", exc)
        return None


_BING_MAX_PAGES = 5
_BING_PAGE_SIZE = 20


def _scrape_bing(niche: str, location: str, max_results: int) -> Iterator[dict]:
    """Yield business dicts from Bing Maps local listings."""
    pool = get_pool()
    collected = 0
    first = 0
    pages_fetched = 0

    while collected < max_results and pages_fetched < _BING_MAX_PAGES:
        query = urllib.parse.quote_plus(f"{niche} near {location}")
        url = (
            f"https://www.bing.com/maps/overlaybfpr?q={query}"
            f"&mapsV10=1&count={_BING_PAGE_SIZE}&first={first}"
        )
        logger.info("[Bing] Fetching first=%d: %s", first, url)

        proxy = pool.get()
        with _make_client(proxy) as client:
            resp = _fetch_with_retry(client, url, proxy=proxy)

        if resp is None or resp.status_code != 200:
            logger.warning("[Bing] No valid response at first=%d, stopping.", first)
            break

        if os.environ.get("BING_DEBUG_HTML") and pages_fetched == 0:
            debug_path = os.environ["BING_DEBUG_HTML"]
            with open(debug_path, "w", encoding="utf-8") as f:
                f.write(resp.text)
            logger.info("[Bing] Saved raw response HTML to %s", debug_path)

        soup = BeautifulSoup(resp.text, "lxml")

        cards = soup.select("div.b_maglistcard[data-entity]")

        if not cards:
            logger.info("[Bing] No more local results at first=%d", first)
            break

        for card in cards:
            biz = _parse_bing_listing(card)
            if biz:
                biz["source"] = "bing"
                yield biz
                collected += 1
                if collected >= max_results:
                    return

        first += _BING_PAGE_SIZE
        pages_fetched += 1
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

        # The business's own website is usually not shown in the Yelp search
        # results card — only a link to the Yelp listing page itself. That
        # is NOT the business's website, and must never be stored as one:
        # resolve.py/crawl.py would then "crawl" Yelp's own page (surfacing
        # Yelp's contact info/tracking links, not the business's), and the
        # WRITE stage's "No Website" sheet would wrongly treat every Yelp
        # business as having a site, hiding it from that sheet even when it
        # truly has none.
        phone_tag = card.select_one("p[class*='phone']") or card.find("p", string=re.compile(r"\(\d{3}\)"))
        phone = phone_tag.get_text(strip=True) if phone_tag else None

        addr_tag = card.select_one("address") or card.select_one("p[class*='address']")
        address = addr_tag.get_text(strip=True) if addr_tag else None

        cat_tag = card.select_one("span[class*='category']")
        category = cat_tag.get_text(strip=True) if cat_tag else None

        return {
            "business_name": name,
            "website_url": None,   # not discoverable from the Yelp search card
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
            resp = _fetch_with_retry(client, url, proxy=proxy)

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
# Google Maps
# ---------------------------------------------------------------------------

def _parse_google_maps_card(card) -> dict | None:
    """
    Parse a single Google Maps search-result card (div[role="article"]).

    Google Maps renders these as loosely-structured "·"-joined text rows
    rather than semantic per-field markup, so this leans on the few
    reliably-scoped selectors that do exist — the name link's aria-label,
    the phone span's dedicated class ("UsdlK"), and the "Visit <name>'s
    website" link's aria-label (an actual business website, as opposed to a
    "Book online" widget link, which is NOT the same thing) — and falls
    back to splitting text rows for category/address.
    """
    try:
        name_link = card.select_one("a.hfpxzc")
        name = name_link.get("aria-label") if name_link else None
        if not name:
            return None

        # Paid ad cards (marked with an aria-label="Sponsored" heading) have
        # a "Visit <name>'s website" link too, but it points at a
        # /aclk?...  Google Ads redirect, not the business's real site —
        # confirmed live. Skip these entirely rather than store a fake URL.
        if card.select_one('h1[aria-label="Sponsored"]'):
            return None

        website = None
        for a in card.select("a[href]"):
            aria = (a.get("aria-label") or "").lower()
            if aria.startswith("visit") and "website" in aria:
                website = a["href"]
                break

        phone_tag = card.select_one("span.UsdlK")
        phone = phone_tag.get_text(strip=True) if phone_tag else None

        category = None
        address = None
        for div in card.select("div.W4Efsd"):
            if div.select_one("div.W4Efsd"):
                continue  # a wrapper div, not a leaf info row
            if div.select_one("span.UsdlK"):
                continue  # the hours/phone row
            if any("star" in (s.get("aria-label") or "").lower() for s in div.select("span[aria-label]")):
                continue  # the star-rating row (role="img" is also used by
                          # unrelated icon glyphs elsewhere, so this checks
                          # the aria-label text specifically, not the role)
            text = div.get_text(" ", strip=True)
            if not text or re.search(r"\b(AM|PM|Open|Closes|Closed)\b", text, re.IGNORECASE):
                continue  # empty, or an hours/status row with no phone number
            if "·" in text:
                # Some cards interleave a private-use-area icon glyph
                # character (e.g. a pin icon) between "·"-joined segments —
                # confirmed live — so filter those out before assembling
                # category/address rather than assuming exactly 2 parts.
                raw_parts = [p.strip() for p in text.split("·")]
                parts = [p for p in raw_parts if p and not all(0xE000 <= ord(ch) <= 0xF8FF for ch in p)]
                if len(parts) >= 2:
                    category, address = parts[0], " ".join(parts[1:])
                elif len(parts) == 1:
                    category = parts[0]
            else:
                # Some cards show only a bare category with no address at
                # all (e.g. a service-area business with no public address).
                category = text
            break

        return {
            "business_name": name,
            "website_url": website,
            "phone": phone,
            "address": address,
            "category": category,
        }
    except Exception as exc:
        logger.debug("Google Maps parse error: %s", exc)
        return None


def _scrape_google_maps(niche: str, location: str, max_results: int) -> Iterator[dict]:
    """
    Yield business dicts from Google Maps search results.

    Google Maps renders its results feed entirely client-side (confirmed
    live: the raw HTTP response contains none of the listing markup), so
    this uses headless Chromium via Playwright rather than a plain GET, and
    scrolls the results feed to load more than the initial page of cards.

    The whole Playwright session runs inside run_in_thread(): Playwright's
    sync API raises if the calling thread already has a running asyncio
    event loop (confirmed live — see browser.run_in_thread's docstring),
    which some Click/CLI invocations do. Since that means this can no
    longer yield progressively while scrolling, results are collected into
    a list inside the thread and handed back as a batch once the thread
    returns — every caller already just iterates this to completion
    regardless, so that's not a behavior change that matters.
    """
    try:
        from playwright.sync_api import sync_playwright
    except ImportError:
        logger.error("Playwright not installed — skipping Google Maps")
        return

    query = urllib.parse.quote_plus(niche) + "+near+" + urllib.parse.quote_plus(location)
    url = f"https://www.google.com/maps/search/{query}"
    logger.info("[GoogleMaps] Fetching: %s", url)

    def _do_scrape() -> list[dict]:
        results: list[dict] = []
        seen_names: set[str] = set()

        with sync_playwright() as pw:
            browser = pw.chromium.launch(
                headless=True,
                args=["--no-sandbox", "--disable-dev-shm-usage"],
            )
            try:
                page = browser.new_page(user_agent=random.choice(USER_AGENTS))
                page.goto(url, timeout=30_000, wait_until="domcontentloaded")
                page.wait_for_timeout(3000)

                feed = page.locator('div[role="feed"]')
                stagnant_scrolls = 0

                for _ in range(30):  # hard cap so a stuck feed can't loop forever
                    count = page.locator('div[role="article"]').count()
                    if count == 0:
                        break

                    soup = BeautifulSoup(page.content(), "lxml")
                    for card in soup.select('div[role="article"]'):
                        biz = _parse_google_maps_card(card)
                        if not biz or biz["business_name"] in seen_names:
                            continue
                        seen_names.add(biz["business_name"])
                        biz["source"] = "google_maps"
                        results.append(biz)
                        if len(results) >= max_results:
                            return results

                    try:
                        feed.evaluate("el => el.scrollTop = el.scrollHeight")
                    except Exception:
                        break
                    page.wait_for_timeout(int(random.uniform(1200, 2000)))

                    new_count = page.locator('div[role="article"]').count()
                    if new_count <= count:
                        stagnant_scrolls += 1
                        if stagnant_scrolls >= 3:
                            break
                    else:
                        stagnant_scrolls = 0
            finally:
                browser.close()

        return results

    try:
        results = run_in_thread(_do_scrape)
    except Exception as exc:
        logger.warning("[GoogleMaps] Session error: %s", exc)
        return

    yield from results


# ---------------------------------------------------------------------------
# Browser-mode (Playwright) scrapers — used two ways:
#   - as a fallback retry for a primary httpx-based source that came back
#     with zero results (--browser), since that's the externally-observable
#     signal a blocked/403'd request actually produces in this codebase —
#     every _fetch_with_retry caller already swallows the failure internally
#     and just yields nothing, rather than raising, so "0 results" is what
#     "blocked" looks like from run_discover()'s side;
#   - as standalone browser-only sources (BBB, Manta) that have no httpx
#     equivalent at all in this file.
# ---------------------------------------------------------------------------

def _scrape_yellowpages_browser(
    niche: str, location: str, max_results: int, headless: bool = True, proxy: str | None = None,
) -> Iterator[dict]:
    """Browser-rendered retry for Yellow Pages, reusing _parse_yp_listing
    since the rendered DOM uses the same markup the httpx version parses.

    The whole session runs inside run_in_thread() — see its docstring: the
    sync Playwright API raises if the calling thread already has a running
    asyncio event loop, which this sidesteps by always running in a fresh
    worker thread instead. That also means results are collected into a
    list inside the thread rather than yielded progressively while
    scrolling — every caller already just iterates this to completion
    regardless, so that's not a behavior change that matters."""
    if proxy is None:
        proxy = get_pool().get()

    query = urllib.parse.quote_plus(niche)
    geo = urllib.parse.quote_plus(location)
    url = f"https://www.yellowpages.com/search?search_terms={query}&geo_location_terms={geo}"
    logger.info("[YP-Browser] Fetching: %s", url)

    def _do_scrape() -> list[dict]:
        results: list[dict] = []
        seen_names: set[str] = set()

        with BrowserSession(headless=headless, proxy=parse_proxy_for_playwright(proxy)) as session:
            page = session.new_page()
            if not safe_goto(page, url):
                return results
            human_delay()

            for _ in range(10):  # page-click cap so a stuck "Next" can't loop forever
                try:
                    page.wait_for_selector(".result, .organic", timeout=10_000)
                except Exception:
                    break

                human_scroll(page)
                soup = BeautifulSoup(page.content(), "lxml")
                for card in soup.select(".result, .organic"):
                    biz = _parse_yp_listing(card)
                    if not biz or biz["business_name"] in seen_names:
                        continue
                    seen_names.add(biz["business_name"])
                    biz["source"] = "yellowpages"
                    results.append(biz)
                    if len(results) >= max_results:
                        return results

                next_btn = page.locator('a.next, [aria-label="Next"]').first
                if next_btn.count() == 0:
                    break
                try:
                    next_btn.click()
                    human_delay()
                except Exception:
                    break

        return results

    try:
        results = run_in_thread(_do_scrape)
    except Exception as exc:
        logger.warning("[YP-Browser] Session error: %s", exc)
        return

    yield from results


def _scrape_yelp_browser(
    niche: str, location: str, max_results: int, headless: bool = True, proxy: str | None = None,
) -> Iterator[dict]:
    """Browser-rendered retry for Yelp, reusing _parse_yelp_listing — which
    deliberately never captures a website_url from the search card (see its
    docstring: Yelp's card links to Yelp's own listing page, never the real
    business site, so this doesn't try to extract an "external link" either,
    despite that otherwise being available on a rendered card).

    Runs inside run_in_thread() — see _scrape_yellowpages_browser's
    docstring for why, and why results are collected rather than yielded
    progressively."""
    if proxy is None:
        proxy = get_pool().get()

    desc = urllib.parse.quote_plus(niche)
    loc = urllib.parse.quote_plus(location)
    url = f"https://www.yelp.com/search?find_desc={desc}&find_loc={loc}"
    logger.info("[Yelp-Browser] Fetching: %s", url)

    def _do_scrape() -> list[dict]:
        results: list[dict] = []
        seen_names: set[str] = set()

        with BrowserSession(headless=headless, proxy=parse_proxy_for_playwright(proxy)) as session:
            page = session.new_page()
            if not safe_goto(page, url):
                return results
            human_delay()

            for _ in range(10):
                try:
                    page.wait_for_selector('li[class*="border-color"]', timeout=10_000)
                except Exception:
                    break

                human_scroll(page)
                soup = BeautifulSoup(page.content(), "lxml")
                cards = soup.select("ul.undefined > li") or soup.select('li[class*="border-color"]')
                if not cards:
                    break

                for card in cards:
                    biz = _parse_yelp_listing(card)
                    if not biz or biz["business_name"] in seen_names:
                        continue
                    seen_names.add(biz["business_name"])
                    biz["source"] = "yelp"
                    results.append(biz)
                    if len(results) >= max_results:
                        return results

                next_btn = page.locator('a[aria-label="Next"]').first
                if next_btn.count() == 0:
                    break
                try:
                    next_btn.click()
                    human_delay()
                except Exception:
                    break

        return results

    try:
        results = run_in_thread(_do_scrape)
    except Exception as exc:
        logger.warning("[Yelp-Browser] Session error: %s", exc)
        return

    yield from results


# NOTE: bbb.org sits behind a Cloudflare interstitial ("Just a moment...")
# that returned the challenge page instead of real results to every attempt
# made while building this scraper — proxied and unproxied, headless and
# (in a quick manual check) headed. The selectors below are a best-effort
# guess at BBB's real markup, following this module's established
# fallback-chain pattern; they're UNVERIFIED and may need adjustment once
# run somewhere BBB doesn't challenge (e.g. a residential proxy / real VPS).
def _parse_bbb_listing(card) -> dict | None:
    """Parse a single BBB search result card. UNVERIFIED — see note above."""
    try:
        name_tag = (
            card.select_one("h3.result-business-name")
            or card.select_one("a.text-blue-600")
            or card.find(["h2", "h3"])
        )
        name = name_tag.get_text(strip=True) if name_tag else None
        if not name:
            return None

        website_tag = card.select_one("a[href*='bbb.org/'] + a") or card.select_one("a.website-link")
        website = website_tag.get("href") if website_tag else None

        phone_tag = card.select_one("[class*='phone']") or card.find(string=re.compile(r"\(\d{3}\)"))
        phone = phone_tag.get_text(strip=True) if hasattr(phone_tag, "get_text") else (
            str(phone_tag).strip() if phone_tag else None
        )

        addr_tag = card.select_one("[class*='address']") or card.select_one("address")
        address = addr_tag.get_text(" ", strip=True) if addr_tag else None

        return {
            "business_name": name,
            "website_url": website,
            "phone": phone,
            "address": address,
            "category": None,
        }
    except Exception as exc:
        logger.debug("BBB parse error: %s", exc)
        return None


def _scrape_bbb_browser(
    niche: str, location: str, max_results: int, headless: bool = True, proxy: str | None = None,
) -> Iterator[dict]:
    """Yield business dicts from the Better Business Bureau. UNVERIFIED —
    see _parse_bbb_listing's note; degrades to zero results rather than
    raising if BBB's Cloudflare challenge blocks the session.

    Runs inside run_in_thread() — see _scrape_yellowpages_browser's
    docstring for why, and why results are collected rather than yielded
    progressively."""
    if proxy is None:
        proxy = get_pool().get()

    niche_q = urllib.parse.quote_plus(niche)
    loc_q = urllib.parse.quote_plus(location)
    url = f"https://www.bbb.org/search?find_text={niche_q}&find_loc={loc_q}"
    logger.info("[BBB-Browser] Fetching: %s", url)

    def _do_scrape() -> list[dict]:
        results: list[dict] = []
        seen_names: set[str] = set()

        with BrowserSession(headless=headless, proxy=parse_proxy_for_playwright(proxy)) as session:
            page = session.new_page()
            if not safe_goto(page, url):
                return results
            human_delay()

            try:
                page.wait_for_selector('[class*="result"]', timeout=10_000)
            except Exception:
                logger.warning("[BBB-Browser] No results rendered — likely blocked")
                return results

            human_scroll(page)
            soup = BeautifulSoup(page.content(), "lxml")
            for card in soup.select('[class*="result-card"], [class*="search-result"]'):
                biz = _parse_bbb_listing(card)
                if not biz or biz["business_name"] in seen_names:
                    continue
                seen_names.add(biz["business_name"])
                biz["source"] = "bbb"
                results.append(biz)
                if len(results) >= max_results:
                    return results

        return results

    try:
        results = run_in_thread(_do_scrape)
    except Exception as exc:
        logger.warning("[BBB-Browser] Session error: %s", exc)
        return

    yield from results


# NOTE: manta.com returned a direct 403 to a plain unproxied request while
# building this scraper. The selectors below are a best-effort guess at
# Manta's real markup, following this module's established fallback-chain
# pattern; UNVERIFIED and may need adjustment against real markup.
def _parse_manta_listing(card) -> dict | None:
    """Parse a single Manta search result card. UNVERIFIED — see note above."""
    try:
        name_tag = (
            card.select_one("h2.company-name")
            or card.select_one("a.company-name")
            or card.find(["h2", "h3"])
        )
        name = name_tag.get_text(strip=True) if name_tag else None
        if not name:
            return None

        website_tag = card.select_one("a.website") or card.select_one("a[href^='http']:not([href*='manta.com'])")
        website = website_tag.get("href") if website_tag else None

        phone_tag = card.select_one("[class*='phone']")
        phone = phone_tag.get_text(strip=True) if phone_tag else None

        addr_tag = card.select_one("[class*='address']") or card.select_one("address")
        address = addr_tag.get_text(" ", strip=True) if addr_tag else None

        cat_tag = card.select_one("[class*='category']")
        category = cat_tag.get_text(strip=True) if cat_tag else None

        return {
            "business_name": name,
            "website_url": website,
            "phone": phone,
            "address": address,
            "category": category,
        }
    except Exception as exc:
        logger.debug("Manta parse error: %s", exc)
        return None


def _scrape_manta_browser(
    niche: str, location: str, max_results: int, headless: bool = True, proxy: str | None = None,
) -> Iterator[dict]:
    """Yield business dicts from Manta. UNVERIFIED — see _parse_manta_listing's
    note; degrades to zero results rather than raising if blocked.

    Runs inside run_in_thread() — see _scrape_yellowpages_browser's
    docstring for why, and why results are collected rather than yielded
    progressively."""
    if proxy is None:
        proxy = get_pool().get()

    niche_q = urllib.parse.quote_plus(niche)
    loc_q = urllib.parse.quote_plus(location)
    url = f"https://www.manta.com/search?search={niche_q}&location={loc_q}"
    logger.info("[Manta-Browser] Fetching: %s", url)

    def _do_scrape() -> list[dict]:
        results: list[dict] = []
        seen_names: set[str] = set()

        with BrowserSession(headless=headless, proxy=parse_proxy_for_playwright(proxy)) as session:
            page = session.new_page()
            if not safe_goto(page, url):
                return results
            human_delay()

            try:
                page.wait_for_selector('[class*="company"]', timeout=10_000)
            except Exception:
                logger.warning("[Manta-Browser] No results rendered — likely blocked")
                return results

            human_scroll(page)
            soup = BeautifulSoup(page.content(), "lxml")
            for card in soup.select('[class*="company-card"], [class*="result"]'):
                biz = _parse_manta_listing(card)
                if not biz or biz["business_name"] in seen_names:
                    continue
                seen_names.add(biz["business_name"])
                biz["source"] = "manta"
                results.append(biz)
                if len(results) >= max_results:
                    return results

        return results

    try:
        results = run_in_thread(_do_scrape)
    except Exception as exc:
        logger.warning("[Manta-Browser] Session error: %s", exc)
        return

    yield from results


# ---------------------------------------------------------------------------
# YellowPages Canada
#
# NOTE: yellowpages.ca is behind a CloudFront WAF that returned a blanket
# 403 "Request blocked" to every request made while building this scraper —
# plain httpx, through 10 different proxies, and even a full headless-
# Chromium session all got the same block on the bare homepage, so this
# could not be verified against real markup. The selectors below follow the
# task's specification and this codebase's established fallback-chain
# pattern; they may need adjustment against real yellowpages.ca HTML.
# ---------------------------------------------------------------------------

def _parse_yp_ca_listing(card) -> dict | None:
    """Parse a single YellowPages.ca result card. UNVERIFIED — see module note."""
    try:
        name_tag = (
            card.select_one("a.listing__name")
            or card.select_one("h3.listing__name")
            or card.select_one(".business-name")
            or card.find(["h2", "h3"])
        )
        name = name_tag.get_text(strip=True) if name_tag else None
        if not name:
            return None

        website_tag = (
            card.select_one("a.mlr__item--website")
            or card.select_one("a[href*='websiteClick']")
            or card.select_one("a.listing__website")
        )
        website = website_tag.get("href") if website_tag else None

        phone_tag = card.select_one(".mlr__item--phone") or card.select_one(".listing__phone")
        phone = phone_tag.get_text(strip=True) if phone_tag else None

        addr_tag = card.select_one(".listing__address") or card.select_one("address")
        address = addr_tag.get_text(" ", strip=True) if addr_tag else None

        cat_tag = card.select_one(".listing__category") or card.select_one(".categories")
        category = cat_tag.get_text(strip=True) if cat_tag else None

        return {
            "business_name": name,
            "website_url": website,
            "phone": phone,
            "address": address,
            "category": category,
        }
    except Exception as exc:
        logger.debug("YP Canada parse error: %s", exc)
        return None


def _scrape_yellowpages_ca(niche: str, location: str, max_results: int) -> Iterator[dict]:
    """Yield business dicts from YellowPages.ca. UNVERIFIED — see module note."""
    pool = get_pool()
    collected = 0
    page = 1

    while collected < max_results:
        niche_slug = urllib.parse.quote(niche)
        loc_slug = urllib.parse.quote(location)
        url = f"https://www.yellowpages.ca/search/si/{page}/{niche_slug}/{loc_slug}"
        logger.info("[YP-CA] Fetching page %d: %s", page, url)

        proxy = pool.get()
        with _make_client(proxy) as client:
            resp = _fetch_with_retry(client, url, proxy=proxy)

        if resp is None or resp.status_code != 200:
            logger.warning("[YP-CA] No valid response for page %d, stopping.", page)
            break

        soup = BeautifulSoup(resp.text, "lxml")
        cards = soup.select("div.listing--enhanced") or soup.select("div.listing")

        if not cards:
            logger.info("[YP-CA] No more results at page %d", page)
            break

        for card in cards:
            biz = _parse_yp_ca_listing(card)
            if biz:
                biz["source"] = "yellowpages_ca"
                yield biz
                collected += 1
                if collected >= max_results:
                    return

        page += 1
        _politeness_sleep()


# ---------------------------------------------------------------------------
# Yell.com (UK)
#
# NOTE: yell.com is behind Cloudflare and returned a "Attention Required!"
# bot-challenge page to every request made while building this scraper —
# plain httpx and a full headless-Chromium session both got the same block
# on the bare search page, so this could not be verified against real
# markup. The selectors below follow the task's specification and this
# codebase's established fallback-chain pattern; they may need adjustment
# against real yell.com HTML.
# ---------------------------------------------------------------------------

def _parse_yell_listing(card) -> dict | None:
    """Parse a single Yell.com result card. UNVERIFIED — see module note."""
    try:
        name_tag = (
            card.select_one("span[itemprop='name']")
            or card.select_one("h2.businessCapsule--title")
            or card.select_one("a.businessCapsule--title")
            or card.find(["h2", "h3"])
        )
        name = name_tag.get_text(strip=True) if name_tag else None
        if not name:
            return None

        website_tag = (
            card.select_one("a.btn--website")
            or card.select_one("a[href*='websiteClick']")
            or card.select_one("a[data-website]")
        )
        website = website_tag.get("href") if website_tag else None

        phone_tag = card.select_one("span.business--telephoneNumber") or card.select_one("[itemprop='telephone']")
        phone = phone_tag.get_text(strip=True) if phone_tag else None

        addr_tag = card.select_one("span[itemprop='address']") or card.select_one(".businessCapsule--address")
        address = addr_tag.get_text(" ", strip=True) if addr_tag else None

        cat_tag = card.select_one(".businessCapsule--classification") or card.select_one("[itemprop='category']")
        category = cat_tag.get_text(strip=True) if cat_tag else None

        return {
            "business_name": name,
            "website_url": website,
            "phone": phone,
            "address": address,
            "category": category,
        }
    except Exception as exc:
        logger.debug("Yell parse error: %s", exc)
        return None


def _scrape_yell(niche: str, location: str, max_results: int) -> Iterator[dict]:
    """Yield business dicts from Yell.com (UK). UNVERIFIED — see module note."""
    pool = get_pool()
    collected = 0
    page = 1

    while collected < max_results:
        niche_slug = urllib.parse.quote(niche)
        loc_slug = urllib.parse.quote(location)
        url = f"https://www.yell.com/s/{niche_slug}/{loc_slug}/"
        if page > 1:
            url += f"page-{page}/"
        logger.info("[Yell] Fetching page %d: %s", page, url)

        proxy = pool.get()
        with _make_client(proxy) as client:
            resp = _fetch_with_retry(client, url, proxy=proxy)

        if resp is None or resp.status_code != 200:
            logger.warning("[Yell] No valid response for page %d, stopping.", page)
            break

        soup = BeautifulSoup(resp.text, "lxml")
        cards = soup.select("div.businessCapsule--directory") or soup.select("div[class*='businessCapsule']")

        if not cards:
            logger.info("[Yell] No more results at page %d", page)
            break

        for card in cards:
            biz = _parse_yell_listing(card)
            if biz:
                biz["source"] = "yell_uk"
                yield biz
                collected += 1
                if collected >= max_results:
                    return

        page += 1
        _politeness_sleep()


# ---------------------------------------------------------------------------
# Thomson Local (UK)
#
# NOTE: thomsonlocal.com returned a 403 to every request made while building
# this scraper (same class of block as yellowpages.ca/yell.com above), so
# this could not be verified against real markup. The task gave no specific
# selectors for this site, so the fallback chain below is a best-effort
# generic guess based on common directory-site markup conventions; it is
# the most likely of the four new scrapers to need real-world adjustment.
# ---------------------------------------------------------------------------

def _parse_thomson_local_listing(card) -> dict | None:
    """Parse a single Thomson Local result card. UNVERIFIED — see module note."""
    try:
        name_tag = (
            card.select_one("a[class*='name']")
            or card.select_one("[class*='businessName']")
            or card.find(["h2", "h3"])
        )
        name = name_tag.get_text(strip=True) if name_tag else None
        if not name:
            return None

        website_tag = card.select_one("a[class*='website']") or card.select_one(
            "a[href^='http']:not([href*='thomsonlocal.com'])"
        )
        website = website_tag.get("href") if website_tag else None

        phone_tag = card.select_one("[class*='phone']") or card.select_one("[class*='telephone']")
        phone = phone_tag.get_text(strip=True) if phone_tag else None

        addr_tag = card.select_one("address") or card.select_one("[class*='address']")
        address = addr_tag.get_text(" ", strip=True) if addr_tag else None

        cat_tag = card.select_one("[class*='category']")
        category = cat_tag.get_text(strip=True) if cat_tag else None

        return {
            "business_name": name,
            "website_url": website,
            "phone": phone,
            "address": address,
            "category": category,
        }
    except Exception as exc:
        logger.debug("Thomson Local parse error: %s", exc)
        return None


def _scrape_thomson_local(niche: str, location: str, max_results: int) -> Iterator[dict]:
    """Yield business dicts from Thomson Local (UK). UNVERIFIED — see module note."""
    pool = get_pool()
    collected = 0
    page = 1

    while collected < max_results:
        niche_slug = urllib.parse.quote(niche)
        loc_slug = urllib.parse.quote(location)
        url = f"https://www.thomsonlocal.com/search/{niche_slug}/{loc_slug}"
        if page > 1:
            url += f"?page={page}"
        logger.info("[ThomsonLocal] Fetching page %d: %s", page, url)

        proxy = pool.get()
        with _make_client(proxy) as client:
            resp = _fetch_with_retry(client, url, proxy=proxy)

        if resp is None or resp.status_code != 200:
            logger.warning("[ThomsonLocal] No valid response for page %d, stopping.", page)
            break

        soup = BeautifulSoup(resp.text, "lxml")
        cards = (
            soup.select("div[class*='listing-card']")
            or soup.select("div[class*='ResultCard']")
            or soup.select("li[class*='result']")
        )

        if not cards:
            logger.info("[ThomsonLocal] No more results at page %d", page)
            break

        for card in cards:
            biz = _parse_thomson_local_listing(card)
            if biz:
                biz["source"] = "thomson_local"
                yield biz
                collected += 1
                if collected >= max_results:
                    return

        page += 1
        _politeness_sleep()


# ---------------------------------------------------------------------------
# Niche relevance filtering — a scraper can return listings that don't
# actually match the requested niche (a directory site miscategorizing a
# business, an aggressive local-pack match, etc.), so every source's
# results get checked against the niche before being persisted.
# ---------------------------------------------------------------------------

_NICHE_SYNONYMS: dict[str, list[str]] = {
    "plumber": ["plumber", "plumbing", "pipe", "drain", "sewer"],
    "dentist": ["dentist", "dental", "orthodont", "teeth", "tooth"],
    "lawyer": ["lawyer", "attorney", "law firm", "legal", "solicitor"],
    "electrician": ["electrician", "electric", "electrical", "wiring"],
    "restaurant": ["restaurant", "cafe", "diner", "eatery", "bistro", "grill"],
    "accountant": ["accountant", "accounting", "bookkeep", "cpa", "tax"],
}


def _niche_keywords(niche: str) -> list[str]:
    """
    Return the keyword list to match a niche's businesses against. Checks
    the synonym map first — matching either the exact word or its stripped
    plural ("plumbers" -> "plumber" entry) — and falls back to the
    niche string's own first-5-character stem for anything not mapped.
    """
    niche_lower = niche.lower().strip()
    for word in re.findall(r"[a-z]+", niche_lower):
        for candidate in (word, word.rstrip("s")):
            if candidate in _NICHE_SYNONYMS:
                return _NICHE_SYNONYMS[candidate]
    stem = niche_lower[:5].strip()
    return [stem] if stem else [niche_lower]


def _is_relevant(
    business_name: str | None, category: str | None, niche: str
) -> tuple[bool, bool]:
    """
    Check whether a scraped listing plausibly matches the requested niche.

    Returns (is_relevant, checked) rather than a bare bool, since the two
    outcomes that both "keep the record" need to be distinguishable from
    each other downstream (relevance_checked is persisted on the row):
      - (True,  True)  — a niche keyword was found in the name or category:
                          confidently relevant.
      - (True,  False) — no category at all, and no keyword match in the
                          name either: too little information to judge, so
                          it's kept but flagged as unchecked/uncertain
                          rather than silently dropped.
      - (False, True)  — a category is present but doesn't match any niche
                          keyword (e.g. niche=plumbers, category=restaurant)
                          — treated as a contradiction and dropped. Simple
                          keyword matching can't distinguish "genuinely
                          contradicts" from "just an unmapped category
                          wording", so any non-matching-but-present category
                          is treated the same conservative way.
    """
    keywords = _niche_keywords(niche)
    name_lower = (business_name or "").lower()
    category_lower = (category or "").lower()

    if any(kw in name_lower for kw in keywords) or any(kw in category_lower for kw in keywords):
        return True, True

    if not category_lower:
        return True, False

    return False, True


# ---------------------------------------------------------------------------
# Region detection — picks which directory sources are worth trying for a
# given --location, since e.g. scraping Yell.com/Thomson Local for a US
# address (or Yellow Pages US for a UK one) would just waste time on a
# directory that doesn't cover that country.
# ---------------------------------------------------------------------------

_US_STATE_ABBR = frozenset([
    "al", "ak", "az", "ar", "ca", "co", "ct", "de", "fl", "ga", "hi", "id",
    "il", "in", "ia", "ks", "ky", "la", "me", "md", "ma", "mi", "mn", "ms",
    "mo", "mt", "ne", "nv", "nh", "nj", "nm", "ny", "nc", "nd", "oh", "ok",
    "or", "pa", "ri", "sc", "sd", "tn", "tx", "ut", "vt", "va", "wa", "wv",
    "wi", "wy", "dc",
])
_CA_PROVINCE_ABBR = frozenset([
    "on", "qc", "bc", "ab", "mb", "sk", "ns", "nb", "pe", "nl", "yt", "nt", "nu",
])
_UK_MARKERS = (
    "united kingdom", "england", "scotland", "wales", "northern ireland",
    "london", "manchester", "birmingham", "glasgow", "liverpool", "bristol",
    "edinburgh", "leeds", "sheffield", "newcastle", "nottingham", "cardiff",
    "belfast",
)
_CA_MARKERS = (
    "canada", "ontario", "toronto", "quebec", "montreal", "vancouver",
    "british columbia", "alberta", "calgary", "ottawa", "winnipeg",
    "edmonton", "hamilton", "mississauga",
)


def _detect_region(location: str) -> str:
    """Classify --location as 'us', 'uk', 'ca', or 'unknown'."""
    loc = location.lower().strip()

    # Most reliable signal: an explicit ", XX" state/province/country suffix.
    tail_match = re.search(r",\s*([a-z]{2,3})\s*$", loc)
    if tail_match:
        tail = tail_match.group(1)
        if tail in _US_STATE_ABBR:
            return "us"
        if tail in _CA_PROVINCE_ABBR:
            return "ca"
        if tail == "uk":
            return "uk"

    if "united states" in loc or "usa" in loc:
        return "us"
    if any(marker in loc for marker in _UK_MARKERS):
        return "uk"
    if any(marker in loc for marker in _CA_MARKERS):
        return "ca"

    return "unknown"


# ---------------------------------------------------------------------------
# Stage entry-point
# ---------------------------------------------------------------------------

# Every available source, keyed by name. Each factory takes (niche,
# location, budget) and returns an Iterator[dict] — the same shape every
# _scrape_* function above already has.
_SOURCE_FACTORIES: dict[str, Callable[[str, str, int], Iterator[dict]]] = {
    "yellowpages": _scrape_yellowpages,
    "bing": _scrape_bing,
    "yelp": _scrape_yelp,
    "google_maps": _scrape_google_maps,
    "yellowpages_ca": _scrape_yellowpages_ca,
    "yell_uk": _scrape_yell,
    "thomson_local": _scrape_thomson_local,
}

# Which sources to run per detected region. The first source in each list is
# treated as "primary" and gets the full max_results budget (matching this
# module's original YP-US-first design); the rest split per_source.
_REGION_SOURCES: dict[str, list[str]] = {
    "us": ["yellowpages", "bing", "yelp", "google_maps"],
    "uk": ["yell_uk", "thomson_local", "google_maps"],
    "ca": ["yellowpages_ca", "google_maps"],
}
_ALL_SOURCES = ["yellowpages", "bing", "yelp", "yellowpages_ca", "yell_uk", "thomson_local", "google_maps"]

# A primary httpx-based source that comes back with 0 results gets retried
# through a real browser instead — only these two have a browser-rendered
# counterpart; the others (Bing, Google Maps, the UK/CA directories) are
# either already Playwright-based or haven't needed one so far.
_BROWSER_RETRY_FACTORIES: dict[str, Callable[..., Iterator[dict]]] = {
    "yellowpages": _scrape_yellowpages_browser,
    "yelp": _scrape_yelp_browser,
}

# Browser-only sources with no httpx equivalent at all — only added when
# --browser is set, and only for US/unrecognized locations (both are
# US-centric directories).
_BROWSER_ONLY_SOURCES: dict[str, Callable[..., Iterator[dict]]] = {
    "bbb": _scrape_bbb_browser,
    "manta": _scrape_manta_browser,
}


def run_discover(
    db_path: str,
    niche: str,
    location: str,
    max_results: int,
    api_fallback: bool = False,
    google_places_key: str | None = None,
    browser_mode: bool = False,
    headless: bool = True,
) -> None:
    """
    Stage 1: scrape directories and populate the businesses table.
    Already-scraped businesses (same name+source) are silently skipped.

    Which directories get scraped depends on --location: a recognisable US/
    UK/Canadian location runs only the sources that actually cover that
    country (no point hitting Yell.com for a Texas address, or Yellow Pages
    US for a London one); anything else runs every source. Each source is
    independent — if one fails outright or returns nothing, that's logged
    and the rest still run; a single bad scraper can't take down the stage.

    If browser_mode is set, a primary source that comes back with 0 results
    (this codebase's fetch helpers already swallow 403/407/etc. internally
    and just yield nothing, rather than raising — so "0 results" is the
    externally-visible shape a block takes) gets retried through a real
    Playwright session instead of plain httpx, and two browser-only
    directories with no httpx path at all (BBB, Manta) are added as extra
    sources for a US/unrecognized location.

    If every proxy-based source still comes back with nothing AND
    api_fallback is set AND a Google Places API key is configured, Google
    Places is tried as a last resort. This never runs otherwise — it costs
    real API quota, so it's only for when the free path has genuinely
    failed, not a routine supplement to it.
    """
    region = _detect_region(location)
    source_names = list(_REGION_SOURCES.get(region, _ALL_SOURCES))
    if browser_mode and region in ("us", "unknown"):
        source_names += list(_BROWSER_ONLY_SOURCES.keys())

    logger.info(
        "=== STAGE 1: DISCOVER  niche=%r  location=%r  max=%d  region=%s  sources=%s  browser_mode=%s ===",
        niche, location, max_results, region, source_names, browser_mode,
    )

    per_source = max(10, max_results // max(1, len(source_names)))

    total_inserted = 0

    def _consume(make_iterator: Callable[[], Iterator[dict]], source_label: str) -> tuple[int, int]:
        """Call make_iterator() and consume its results, upserting relevant
        ones. Returns (inserted_count, filtered_count); exceptions are
        caught and logged, never allowed to take down the rest of the
        stage — including one raised synchronously by make_iterator()
        itself, not just one raised partway through iteration, which is
        why the call happens inside this try rather than at the call site."""
        nonlocal total_inserted
        inserted = 0
        filtered = 0
        try:
            for biz in make_iterator():
                is_relevant, checked = _is_relevant(
                    biz.get("business_name"), biz.get("category"), niche
                )
                if not is_relevant:
                    filtered += 1
                    continue

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
                        relevance_checked=checked,
                    )
                inserted += 1
                total_inserted += 1
        except Exception as exc:
            logger.warning(
                "[DISCOVER] Source %s failed (%s) — continuing with remaining sources",
                source_label, exc,
            )
        return inserted, filtered

    for i, source_name in enumerate(source_names):
        budget = max_results if i == 0 else per_source

        if source_name in _BROWSER_ONLY_SOURCES:
            logger.info("[DISCOVER] Starting browser-only source: %s (budget=%d)", source_name, budget)
            browser_only_factory = _BROWSER_ONLY_SOURCES[source_name]
            count, filtered = _consume(
                lambda f=browser_only_factory: f(niche, location, budget, headless=headless),
                source_name,
            )
        else:
            factory = _SOURCE_FACTORIES[source_name]
            logger.info("[DISCOVER] Starting source: %s (budget=%d)", source_name, budget)
            count, filtered = _consume(lambda f=factory: f(niche, location, budget), source_name)

            if count == 0 and browser_mode and source_name in _BROWSER_RETRY_FACTORIES:
                logger.info("[DISCOVER] Retrying %s with browser mode", source_name)
                browser_factory = _BROWSER_RETRY_FACTORIES[source_name]
                browser_count, browser_filtered = _consume(
                    lambda f=browser_factory: f(niche, location, budget, headless=headless),
                    f"{source_name} (browser)",
                )
                count += browser_count
                filtered += browser_filtered

        if filtered:
            logger.info("[DISCOVER] Filtered %d irrelevant results from %s", filtered, source_name)

        if count == 0:
            logger.warning("[DISCOVER] Source %s returned 0 results", source_name)

        logger.info("[DISCOVER] %s: inserted/seen %d businesses", source_name, count)

    if total_inserted == 0 and api_fallback and google_places_key:
        logger.info("[DISCOVER] All proxy sources failed — trying Google Places API fallback")
        from .google_api import google_places_search

        fallback_count = 0
        try:
            for biz in google_places_search(niche, location, google_places_key, max_results):
                is_relevant, checked = _is_relevant(
                    biz.get("business_name"), biz.get("category"), niche
                )
                if not is_relevant:
                    continue
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
                        relevance_checked=checked,
                    )
                fallback_count += 1
                total_inserted += 1
        except Exception as exc:
            logger.warning("[DISCOVER] Google Places fallback failed: %s", exc)

        logger.info("[DISCOVER] Google Places fallback added %d businesses", fallback_count)

    with get_conn(db_path) as conn:
        counts = count_businesses(conn, niche, location)

    logger.info(
        "[DISCOVER] DONE — total=%d  with_site=%d  site_less=%d",
        counts["total"],
        counts["with_site"],
        counts["total"] - counts["with_site"],
    )
