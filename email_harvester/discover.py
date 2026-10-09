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
#
# Field extraction here works directly against Playwright ElementHandles
# (query_selector/inner_text) rather than a BeautifulSoup snapshot, via
# _try_selectors/_try_selector_attr below — each field gets a list of
# fallback CSS selectors tried in order, since a site's markup commonly
# has more than one variant in the wild (A/B tests, listing types, etc.)
# and a single hardcoded selector silently matching nothing was exactly
# the bug this was written to fix.
# ---------------------------------------------------------------------------

_PHONE_LIKE_RE = re.compile(r"\(?\d{3}\)?[\s.\-]?\d{3}[\s.\-]?\d{4}")


def _try_selectors(element, selectors: list[str]) -> str:
    """Try multiple CSS selectors against a Playwright element, return the
    first non-empty inner_text() found."""
    for selector in selectors:
        try:
            el = element.query_selector(selector)
            if el:
                text = el.inner_text().strip()
                if text:
                    return text
        except Exception:
            continue
    return ""


def _try_selector_attr(element, selectors: list[str], attr: str) -> str:
    """Like _try_selectors, but returns an attribute value (e.g. href)
    instead of text — needed for website links, which _try_selectors
    alone can't get at."""
    for selector in selectors:
        try:
            el = element.query_selector(selector)
            if el:
                value = el.get_attribute(attr)
                if value:
                    return value.strip()
        except Exception:
            continue
    return ""


def _find_phone_like_text(element, selector: str) -> str:
    """Scan every element matching selector for one whose text looks like
    a phone number — used where a site has no dedicated phone class/
    attribute to select directly, only a generic container that also
    holds other text."""
    try:
        for el in element.query_selector_all(selector):
            text = el.inner_text().strip()
            if _PHONE_LIKE_RE.search(text):
                return text
    except Exception:
        pass
    return ""


def _take_debug_screenshot(page, path: str, source_label: str) -> None:
    """Best-effort screenshot for diagnosing a zero-result scrape — e.g. to
    tell a changed selector apart from a bot-block page from the next run.
    Never raises: a failed screenshot (missing directory, closed page,
    etc.) is a secondary problem, not worth losing the real result over."""
    try:
        page.screenshot(path=path)
        logger.warning("[%s] Zero results — saved debug screenshot to %s", source_label, path)
    except Exception as exc:
        logger.debug("[%s] Screenshot failed: %s", source_label, exc)


def _save_debug_html(page, path: str, source_label: str) -> None:
    """Best-effort page-HTML dump alongside a debug screenshot — lets a
    block page and a markup change be told apart without re-running live.
    Never raises, same rationale as _take_debug_screenshot."""
    try:
        with open(path, "w", encoding="utf-8") as f:
            f.write(page.content())
        logger.warning("[%s] Page HTML saved to %s", source_label, path)
    except Exception as exc:
        logger.debug("[%s] HTML dump failed: %s", source_label, exc)


_YP_CARD_SELECTOR = "div.result, div.v-card"
_YP_NAME_SELECTORS = ["a.business-name > span", "h2.n > a"]
_YP_PHONE_SELECTORS = ["div.phones.phone.primary", "p.phone"]
_YP_STREET_SELECTORS = ["span.street-address"]
_YP_LOCALITY_SELECTORS = ["span.locality"]
_YP_CATEGORY_SELECTORS = ["div.categories > a", "p.body > a"]
_YP_WEBSITE_SELECTORS = ["a.track-visit-website", "a[data-analytics='website']"]


def _scrape_yellowpages_browser(
    niche: str, location: str, max_results: int, headless: bool = True, proxy: str | None = None,
) -> Iterator[dict]:
    """Browser-rendered retry for Yellow Pages, extracting fields directly
    from the rendered DOM (not a bs4 snapshot) via _try_selectors, with
    several fallback selectors per field.

    UNVERIFIED against real result markup: every attempt made while
    building this fix — proxied and unproxied — hit a Cloudflare
    "Attention Required!" interstitial instead of real search results
    (confirmed live), the same sandbox-wide block already documented for
    bbb.org/manta.com/yell.com elsewhere in this file. The selectors below
    follow the exact structure given in the task spec; on 0 results this
    now saves a debug screenshot to /tmp/yp_debug.png (POSIX path — only
    meaningful on a real Linux VPS, not this Windows dev sandbox) so that
    can be checked directly instead of guessing blind.

    Runs inside run_in_thread() — see its docstring: the sync Playwright
    API raises if the calling thread already has a running asyncio event
    loop, which this sidesteps by always running in a fresh worker thread
    instead. That also means results are collected into a list inside the
    thread rather than yielded progressively while scrolling — every
    caller already just iterates this to completion regardless, so that's
    not a behavior change that matters."""
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
                    page.wait_for_selector(_YP_CARD_SELECTOR, timeout=15_000)
                except Exception:
                    logger.warning("[YP-Browser] Selector timed out — no results page detected")
                    break

                human_scroll(page)
                cards = page.query_selector_all(_YP_CARD_SELECTOR)
                logger.info("[YP-Browser] Found %d result cards", len(cards))

                for card in cards:
                    name = _try_selectors(card, _YP_NAME_SELECTORS)
                    if not name or name in seen_names:
                        continue
                    seen_names.add(name)

                    street = _try_selectors(card, _YP_STREET_SELECTORS)
                    locality = _try_selectors(card, _YP_LOCALITY_SELECTORS)
                    address = ", ".join(p for p in (street, locality) if p) or None

                    biz = {
                        "business_name": name,
                        "website_url": _try_selector_attr(card, _YP_WEBSITE_SELECTORS, "href") or None,
                        "phone": _try_selectors(card, _YP_PHONE_SELECTORS) or None,
                        "address": address,
                        "category": _try_selectors(card, _YP_CATEGORY_SELECTORS) or None,
                        "source": "yellowpages",
                    }
                    logger.debug("[YP-Browser] Extracted: %s", biz)
                    results.append(biz)
                    if len(results) >= max_results:
                        break

                if len(results) >= max_results:
                    break

                next_btn = page.query_selector("a.next.ajax-page")
                if next_btn is None:
                    break
                try:
                    next_btn.click()
                    human_delay()
                except Exception:
                    break

            if len(results) == 0:
                _take_debug_screenshot(page, "/tmp/yp_debug.png", "YP-Browser")

        return results

    try:
        results = run_in_thread(_do_scrape)
    except Exception as exc:
        logger.warning("[YP-Browser] Session error: %s", exc)
        return

    yield from results


# NOTE: yelp.com sits behind DataDome, a dedicated anti-bot service — not
# a generic Cloudflare interstitial. Confirmed live, consistently across
# 7 of 8 attempts made while fixing this (proxied and unproxied; the 8th
# was a plain navigation timeout): the response is a near-empty page
# (title literally "yelp.com", ~1.6KB) embedding an
# <iframe src="https://geo.captcha-delivery.com/captcha/?..."> — DataDome's
# CAPTCHA challenge. This is the exact scenario the task's own spec
# anticipated ("if it's a CAPTCHA page, Yelp requires residential proxies
# — no workaround without them"). One concrete finding from capturing that
# real page: its <title> is plain "yelp.com", containing none of
# "access denied"/"robot"/"captcha"/"blocked" — a title-only check
# (as originally specified) would silently miss this real block entirely.
# _is_yelp_blocked() below also checks for the captcha-delivery.com iframe
# directly, which is what actually caught it. The card/field selectors
# below are UNVERIFIED against real result markup, since no attempt here
# ever got past DataDome to a real results page.
_YELP_WAIT_SELECTOR = 'div[data-testid="serp-ia-card"], li.y-css-1iy9ks6'
_YELP_CARD_SELECTOR_CHAIN = [
    "div[data-testid='serp-ia-card']",
    "li.y-css-1iy9ks6",
    "div.businessName__09f24__EYSZE",
    "div[class*='businessName']",
]
_YELP_NAME_SELECTORS = ["a[class*='businessName']", "h3 > a[name]", "span[class*='display-name']"]
_YELP_ADDRESS_SELECTORS = ["address > p", "span[class*='raw-text']", "p[class*='address']"]
_YELP_CATEGORY_SELECTORS = ["span[class*='category']", "a[class*='category-str-list']"]
_YELP_RATING_SELECTOR = "div[aria-label*='star rating']"
_YELP_BLOCK_TITLE_MARKERS = ("access denied", "robot", "captcha", "blocked")
_YELP_BLOCK_CONTENT_MARKERS = ("captcha-delivery.com", "datadome")

_YELP_EXTRA_HEADERS = {
    "Referer": "https://www.google.com/",
    "Sec-Fetch-Dest": "document",
    "Sec-Fetch-Mode": "navigate",
    "Sec-Fetch-Site": "cross-site",
    "Upgrade-Insecure-Requests": "1",
}


def _is_yelp_blocked(page) -> bool:
    """True if this page is Yelp's bot/CAPTCHA challenge rather than real
    results. Checks the title markers the task spec suggested, plus the
    DataDome iframe/content signal confirmed live — see the module note
    above for why the title check alone isn't enough."""
    try:
        title = (page.title() or "").lower()
        if any(marker in title for marker in _YELP_BLOCK_TITLE_MARKERS):
            return True
    except Exception:
        pass
    try:
        if page.query_selector("div#captcha-container, div.yelp-error"):
            return True
    except Exception:
        pass
    try:
        if page.query_selector('iframe[src*="captcha-delivery"]'):
            return True
    except Exception:
        pass
    try:
        html_lower = page.content().lower()
        if any(marker in html_lower for marker in _YELP_BLOCK_CONTENT_MARKERS):
            return True
    except Exception:
        pass
    return False


def _extract_yelp_cards(page):
    """Try each card selector in priority order, return the first that
    matches anything (mirrors _try_selectors' fallback-chain approach, but
    for finding the repeating card list itself rather than a field)."""
    for selector in _YELP_CARD_SELECTOR_CHAIN:
        cards = page.query_selector_all(selector)
        if cards:
            return cards
    return []


def _extract_yelp_rating(card) -> str:
    try:
        el = card.query_selector(_YELP_RATING_SELECTOR)
        if el:
            label = el.get_attribute("aria-label")
            if label:
                return label.strip()
    except Exception:
        pass
    return ""


def _scrape_yelp_browser(
    niche: str, location: str, max_results: int, headless: bool = True, proxy: str | None = None,
) -> Iterator[dict]:
    """Browser-rendered retry for Yelp. UNVERIFIED against real result
    markup — see the module note above; this sandbox never got past
    Yelp's DataDome challenge. Neither phone nor website_url is captured
    here: Yelp doesn't show either on the search card itself (phone is
    found by the CRAWL stage once a site is known; website_url is left
    for RESOLVE's business-name search, same reasoning already documented
    on _parse_yelp_listing for the httpx-based scraper).

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
            try:
                page.set_extra_http_headers(_YELP_EXTRA_HEADERS)
            except Exception:
                pass

            if not safe_goto(page, url):
                _take_debug_screenshot(page, "/tmp/yelp_debug.png", "Yelp-Browser")
                _save_debug_html(page, "/tmp/yelp_debug.html", "Yelp-Browser")
                return results

            try:
                page.wait_for_load_state("networkidle", timeout=25_000)
            except Exception:
                pass  # best-effort — still proceed to the checks below

            # Human-like pacing before touching the page content.
            try:
                page.mouse.move(random.randint(100, 800), random.randint(100, 500))
            except Exception:
                pass
            time.sleep(random.uniform(0.5, 1.5))
            human_scroll(page, times=2)
            time.sleep(random.uniform(1.0, 2.5))

            if _is_yelp_blocked(page):
                logger.warning("[Yelp-Browser] Blocked by CAPTCHA/bot detection")
                _take_debug_screenshot(page, "/tmp/yelp_blocked.png", "Yelp-Browser")
                return results

            try:
                page.wait_for_selector(_YELP_WAIT_SELECTOR, timeout=15_000)
            except Exception:
                logger.warning("[Yelp-Browser] Selector timed out — no results page detected")
                _take_debug_screenshot(page, "/tmp/yelp_debug.png", "Yelp-Browser")
                _save_debug_html(page, "/tmp/yelp_debug.html", "Yelp-Browser")
                return results

            for _ in range(10):  # page-click cap so a stuck "Next" can't loop forever
                # Scroll to trigger any lazy-loaded cards before (re-)querying.
                for _ in range(3):
                    try:
                        page.evaluate("window.scrollBy(0, 800)")
                    except Exception:
                        break
                    time.sleep(random.uniform(1.0, 2.0))

                cards = _extract_yelp_cards(page)
                logger.info("[Yelp-Browser] Found %d result cards", len(cards))

                for card in cards:
                    name = _try_selectors(card, _YELP_NAME_SELECTORS)
                    if not name or name in seen_names:
                        continue
                    seen_names.add(name)

                    category = _try_selectors(card, _YELP_CATEGORY_SELECTORS)
                    rating = _extract_yelp_rating(card)
                    if rating:
                        category = f"{category} {rating}" if category else rating

                    biz = {
                        "business_name": name,
                        "website_url": None,
                        "phone": None,
                        "address": _try_selectors(card, _YELP_ADDRESS_SELECTORS) or None,
                        "category": category or None,
                        "source": "yelp",
                    }
                    logger.debug("[Yelp-Browser] Extracted: %s", biz)
                    results.append(biz)
                    if len(results) >= max_results:
                        break

                if len(results) >= max_results:
                    break

                next_btn = (
                    page.query_selector("a[aria-label='Next']")
                    or page.query_selector("a.next-link")
                    or page.query_selector("a[class*='pagination-link_anchor'][rel='next']")
                )
                if next_btn is None:
                    break
                try:
                    next_btn.click()
                    human_delay()
                except Exception:
                    break

            if len(results) == 0:
                _take_debug_screenshot(page, "/tmp/yelp_debug.png", "Yelp-Browser")
                _save_debug_html(page, "/tmp/yelp_debug.html", "Yelp-Browser")

        return results

    try:
        results = run_in_thread(_do_scrape)
    except Exception as exc:
        logger.warning("[Yelp-Browser] Session error: %s", exc)
        return

    yield from results


# NOTE: bbb.org sits behind an intermittent Cloudflare interstitial ("Just
# a moment..." / "You have been blocked") — roughly 2 of every 3 attempts
# made while fixing this hit the challenge instead of real results, but
# the third got through. That real page was captured and used to verify
# the selectors below directly (not guessed): the card container is
# `div.result-card`, not the task spec's guessed `[data-card-type='biz']`/
# `.MuiGrid-item` (BBB's real markup has neither — no MUI, no schema.org
# itemprop attributes at all). Those guessed selectors are kept as
# trailing fallbacks in case BBB's markup varies by result type, but the
# first selector in each list below is confirmed against genuine BBB
# output. One confirmed real-world gap: none of 15 real cards checked had
# an external website link on the search-results page at all (only a
# profile-page link, a "Get a Quote" link, and a tel: link) — BBB simply
# doesn't expose it at this stage, so website_url will realistically stay
# empty for BBB results regardless of selector; this is the same
# documented limitation already noted for Yelp's _parse_yelp_listing. On
# 0 results this now saves a debug screenshot to /tmp/bbb_debug.png
# (POSIX path — meaningful on a real Linux VPS, not this Windows dev
# sandbox) so an actual block vs. a markup change can be told apart.
_BBB_WAIT_SELECTOR = "div.result-card, div[data-card-type='biz'], .MuiGrid-item"
_BBB_CARD_SELECTOR = "div.result-card, div[data-card-type='biz'], .MuiGrid-item"
_BBB_NAME_SELECTORS = ["h3.result-business-name", 'span[itemprop="name"]', "h3.MuiTypography-h3"]
_BBB_PHONE_SELECTORS = ['a[href^="tel:"]', 'span[itemprop="telephone"]']
_BBB_PHONE_FALLBACK_SELECTOR = "div.MuiBox-root > p"
_BBB_ADDRESS_SELECTORS = ["p.text-size-5.text-gray-70"]
_BBB_STREET_SELECTORS = ['span[itemprop="streetAddress"]']
_BBB_LOCALITY_SELECTORS = ['span[itemprop="addressLocality"]']
_BBB_CATEGORY_SELECTORS = ["p.text-size-4.text-gray-70"]
_BBB_WEBSITE_SELECTORS = [
    'a[data-testid="biz-website"]',
    'a[href*="http"]:not([href*="bbb.org"])',
]
# Sponsored/ad listings (confirmed live) render an "advertisement:" label
# as part of the same heading inner_text() reads the name from, e.g.
# "advertisement:\nAbacus Plumbing, Air Conditioning & Electrical" — strip
# it so it doesn't pollute the extracted business name.
_BBB_AD_PREFIX_RE = re.compile(r"^advertisement:\s*", re.IGNORECASE)


def _scrape_bbb_browser(
    niche: str, location: str, max_results: int, headless: bool = True, proxy: str | None = None,
) -> Iterator[dict]:
    """Yield business dicts from the Better Business Bureau. Selectors
    confirmed against genuine real-world result markup (see module note
    above); still degrades to zero results rather than raising when
    BBB's intermittent Cloudflare challenge blocks a given session.

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
                page.wait_for_load_state("networkidle", timeout=20_000)
            except Exception:
                pass  # best-effort — still try the selector wait below

            try:
                page.wait_for_selector(_BBB_WAIT_SELECTOR, timeout=15_000)
            except Exception:
                logger.warning("[BBB-Browser] Selector timed out — likely blocked")
                _take_debug_screenshot(page, "/tmp/bbb_debug.png", "BBB-Browser")
                return results

            human_scroll(page)
            cards = page.query_selector_all(_BBB_CARD_SELECTOR)
            logger.info("[BBB-Browser] Found %d business cards", len(cards))

            for card in cards:
                name = _BBB_AD_PREFIX_RE.sub("", _try_selectors(card, _BBB_NAME_SELECTORS)).strip()
                if not name or name in seen_names:
                    continue
                seen_names.add(name)

                phone = _try_selectors(card, _BBB_PHONE_SELECTORS)
                if not phone:
                    phone = _find_phone_like_text(card, _BBB_PHONE_FALLBACK_SELECTOR)

                # Confirmed real markup combines the full address into one
                # element; the itemprop street/locality split is kept as a
                # fallback for a schema.org-marked-up variant, unverified.
                address = _try_selectors(card, _BBB_ADDRESS_SELECTORS)
                if not address:
                    street = _try_selectors(card, _BBB_STREET_SELECTORS)
                    locality = _try_selectors(card, _BBB_LOCALITY_SELECTORS)
                    address = ", ".join(p for p in (street, locality) if p)

                biz = {
                    "business_name": name,
                    "website_url": _try_selector_attr(card, _BBB_WEBSITE_SELECTORS, "href") or None,
                    "phone": phone or None,
                    "address": address or None,
                    "category": _try_selectors(card, _BBB_CATEGORY_SELECTORS) or None,
                    "source": "bbb",
                }
                logger.debug("[BBB-Browser] Extracted: %s", biz)
                results.append(biz)
                if len(results) >= max_results:
                    break

            if len(results) == 0:
                _take_debug_screenshot(page, "/tmp/bbb_debug.png", "BBB-Browser")

        return results

    try:
        results = run_in_thread(_do_scrape)
    except Exception as exc:
        logger.warning("[BBB-Browser] Session error: %s", exc)
        return

    yield from results


# NOTE: manta.com is a harder case than YP/BBB — most attempts hit a
# Cloudflare "Just a moment..." challenge, but one attempt actually got
# through to a real (non-challenged) page and still found 0 cards. That
# page's body text read "We encountered an error while performing your
# search." rather than showing any listings — i.e. this specific query
# string (?search=X&location=Y, which does match the real homepage search
# form's own field names/ids) reached Manta's real backend but didn't
# return results the way a direct URL hit apparently expects. This looks
# more like Manta's search needing the homepage form actually filled in
# and submitted (a client-rendered flow) rather than a pure selector
# problem, but that's a bigger rework than selectors alone and wasn't
# confirmed either way in the time available. The selectors below still
# follow the exact structure given in the task spec; UNVERIFIED against
# real result markup, and on 0 results now saves a debug screenshot to
# /tmp/manta_debug.png (POSIX path — meaningful on a real Linux VPS, not
# this Windows dev sandbox) so an actual block vs. this "no results"
# response can be told apart from a markup change.
_MANTA_WAIT_SELECTOR = 'div.search-results, article.company-result, div[data-cy="company-card"]'
_MANTA_CARD_SELECTOR = 'article.company-result, div[data-cy="company-card"]'
_MANTA_NAME_SELECTORS = ["h2.company-name > a", 'a[data-cy="company-name"]']
_MANTA_PHONE_SELECTORS = ["span.phone", 'div[data-cy="phone"]']
_MANTA_ADDRESS_SELECTORS = ["span.address", 'div[data-cy="address"]']
_MANTA_WEBSITE_SELECTORS = ['a[data-cy="website"]', "a.website-link"]


def _scrape_manta_browser(
    niche: str, location: str, max_results: int, headless: bool = True, proxy: str | None = None,
) -> Iterator[dict]:
    """Yield business dicts from Manta. UNVERIFIED — see the module note
    above; degrades to zero results rather than raising if blocked.

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
                page.wait_for_load_state("networkidle", timeout=20_000)
            except Exception:
                pass  # best-effort — still try the selector wait below

            try:
                page.wait_for_selector(_MANTA_WAIT_SELECTOR, timeout=15_000)
            except Exception:
                logger.warning("[Manta-Browser] Selector timed out — likely blocked")
                _take_debug_screenshot(page, "/tmp/manta_debug.png", "Manta-Browser")
                return results

            human_scroll(page)
            cards = page.query_selector_all(_MANTA_CARD_SELECTOR)
            logger.info("[Manta-Browser] Found %d company cards", len(cards))

            for card in cards:
                name = _try_selectors(card, _MANTA_NAME_SELECTORS)
                if not name or name in seen_names:
                    continue
                seen_names.add(name)

                biz = {
                    "business_name": name,
                    "website_url": _try_selector_attr(card, _MANTA_WEBSITE_SELECTORS, "href") or None,
                    "phone": _try_selectors(card, _MANTA_PHONE_SELECTORS) or None,
                    "address": _try_selectors(card, _MANTA_ADDRESS_SELECTORS) or None,
                    "category": None,
                    "source": "manta",
                }
                logger.debug("[Manta-Browser] Extracted: %s", biz)
                results.append(biz)
                if len(results) >= max_results:
                    break

            if len(results) == 0:
                _take_debug_screenshot(page, "/tmp/manta_debug.png", "Manta-Browser")

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
