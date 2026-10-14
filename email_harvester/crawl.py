"""
Stage 3 — CRAWL + EXTRACT.

For each resolved site:
  1. Fetch homepage + /contact, /about, /team + same-domain contact/about links
  2. Extract emails from each page via extract.py
  3. If zero emails found statically → retry ONCE with Playwright Chromium
  4. Store extracted emails in the emails table

Resumable: crawl_status is one of:
  pending     — not yet attempted
  done        — crawled successfully, emails found
  no_emails   — crawled successfully, zero emails found anywhere on the site
  blocked     — HTTP 403/429 persisted through retries — not retried again
  temp_failed — connection error/timeout — retried on the next run
  invalid_url — HTTP 404/410/other non-200 — not retried again
crawl_error holds the specific reason (an HTTP code or exception message)
behind a blocked/temp_failed/invalid_url status.
Politeness: 3–8s random delay between sites; 2–3 concurrent workers (sync here).
"""

import logging
import random
import re
import time
import urllib.parse
from typing import Optional

import httpx
from bs4 import BeautifulSoup

from .browser import BrowserSession, dismiss_cookie_banner, human_delay, human_scroll, run_in_thread
from .db import get_conn, upsert_email, upsert_social
from .extract import extract_emails
from .proxy import get_pool
from .social import _extract_social_from_html

logger = logging.getLogger(__name__)

DELAY_MIN = 3.0
DELAY_MAX = 8.0
MAX_PAGES_PER_SITE = 8   # cap to keep runtime reasonable

# Anchor text patterns that suggest contact/about pages
_CONTACT_ANCHOR_RE = re.compile(
    r"\b(contact|about|team|staff|people|reach\s+us|get\s+in\s+touch)\b",
    re.IGNORECASE,
)

# Standard sub-paths to always try
_DEFAULT_PATHS = ["/contact", "/contact-us", "/about", "/about-us", "/team", "/staff"]

USER_AGENTS = [
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/17.4 Safari/605.1.15",
    "Mozilla/5.0 (X11; Linux x86_64; rv:125.0) Gecko/20100101 Firefox/125.0",
]


def _random_headers() -> dict[str, str]:
    return {
        "User-Agent": random.choice(USER_AGENTS),
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8",
        "Accept-Language": "en-US,en;q=0.9",
        "Accept-Encoding": "gzip, deflate, br",
        "DNT": "1",
    }


def _make_client(proxy: str | None) -> httpx.Client:
    proxies = {"http://": proxy, "https://": proxy} if proxy else None
    headers = _random_headers()
    headers.update(get_pool().get_auth_header(proxy))
    return httpx.Client(
        proxies=proxies,
        follow_redirects=True,
        timeout=httpx.Timeout(20.0),
        verify=True,
        headers=headers,
    )


def _base_domain(url: str) -> str:
    """Return scheme + netloc for same-domain link detection."""
    parsed = urllib.parse.urlparse(url)
    return f"{parsed.scheme}://{parsed.netloc}"


def _same_domain(link: str, base: str) -> bool:
    """True if link belongs to the same netloc as base."""
    try:
        link_host = urllib.parse.urlparse(link).netloc
        base_host = urllib.parse.urlparse(base).netloc
        return link_host == base_host or link_host.endswith("." + base_host)
    except Exception:
        return False


def _absolute(href: str, base_url: str) -> str | None:
    """Resolve relative href against base_url; return None for non-http(s)."""
    try:
        abs_url = urllib.parse.urljoin(base_url, href)
        if urllib.parse.urlparse(abs_url).scheme not in ("http", "https"):
            return None
        return abs_url
    except Exception:
        return None


def _fetch_page(
    client: httpx.Client, url: str, proxy: str | None = None,
) -> tuple[Optional[str], Optional[str], Optional[str]]:
    """
    Fetch a single page. Returns (html, status, error):
      - success                          : (html_text, None, None)
      - 403/429/407 persisted thru retries: (None, "blocked", "HTTP <code>")
      - connection error/timeout persisted : (None, "temp_failed", str(exc))
      - any other non-200 (404/410/etc.)  : (None, "invalid_url", "HTTP <code>")
    status/error are only meaningful to the caller for the homepage fetch —
    see _crawl_site_static's docstring; a sub-page failure is routine
    (every site has some dead links) and callers looping over sub-pages
    just check html is None, same as before this returned extra detail.
    """
    pool = get_pool()
    last_status: Optional[str] = None
    last_error: Optional[str] = None
    for attempt in range(3):
        try:
            resp = client.get(url)
            if resp.status_code == 407:
                pool.mark_407(proxy)
                logger.warning("407 Proxy Auth failed for %s — check proxy credentials", url)
                return None, "blocked", "HTTP 407"  # don't retry — won't fix itself with the same proxy
            if resp.status_code in (403, 429):
                logger.warning("HTTP %d for %s (attempt %d)", resp.status_code, url, attempt)
                last_status, last_error = "blocked", f"HTTP {resp.status_code}"
                pool.backoff_sleep(attempt + 1)
                continue
            if resp.status_code == 200:
                return resp.text, None, None
            # 404/410 etc — page doesn't exist, stop trying
            logger.debug("HTTP %d for %s — skipping", resp.status_code, url)
            return None, "invalid_url", f"HTTP {resp.status_code}"
        except (httpx.ConnectError, httpx.TimeoutException, httpx.RemoteProtocolError) as exc:
            logger.debug("Fetch error %s: %s (attempt %d)", url, exc, attempt)
            last_status, last_error = "temp_failed", str(exc)
            pool.backoff_sleep(attempt + 1)
    return None, last_status, last_error


def _discover_contact_links(html: str, base_url: str) -> list[str]:
    """Return same-domain URLs whose anchor text matches contact/about patterns."""
    soup = BeautifulSoup(html, "lxml")
    seen: set[str] = set()
    links: list[str] = []

    for a in soup.find_all("a", href=True):
        text = a.get_text(strip=True)
        href = a["href"]
        if _CONTACT_ANCHOR_RE.search(text) or _CONTACT_ANCHOR_RE.search(href):
            abs_url = _absolute(href, base_url)
            if abs_url and _same_domain(abs_url, base_url) and abs_url not in seen:
                seen.add(abs_url)
                links.append(abs_url)

    return links[:MAX_PAGES_PER_SITE]


def _crawl_site_static(
    base_url: str,
) -> tuple[list[dict], list[str], dict, Optional[str], Optional[str]]:
    """
    Fetch homepage + standard paths + discovered contact links.
    Returns (emails_list, pages_visited, social_links, failure_status, failure_error).
    social_links is {"facebook": url|None, "instagram": url|None, "linkedin": url|None},
    merged across every page visited (first match per platform wins).

    failure_status/failure_error are both None on any success — including
    a homepage that loaded fine but ultimately yielded zero emails, which
    is run_crawl()'s job to classify as 'no_emails', not this function's.
    They're only set when the homepage itself could never be fetched at
    all, classifying why: 'blocked' (403/429/407 persisted through
    retries), 'temp_failed' (connection error/timeout persisted through
    retries), or 'invalid_url' (404/410/other non-200). A sub-page that
    fails after a successful homepage doesn't affect this — every real
    site has some dead internal links, and the homepage already proved
    the site itself is reachable.

    Tries a direct (no-proxy) connection first — a proxy from the pool is
    only spent if the direct attempt comes back with nothing at all
    (connection error, timeout, or any other failure _fetch_page couldn't
    resolve on its own retries). Once the homepage decides which one
    actually works, the rest of that site's pages reuse the same choice
    rather than re-deciding per page.
    """
    pool = get_pool()
    proxy = pool.get()

    all_emails: dict[str, dict] = {}  # email -> first-seen dict
    pages_visited: list[str] = []
    social: dict[str, str | None] = {"facebook": None, "instagram": None, "linkedin": None}

    def _merge_social(page_html: str) -> None:
        for key, value in _extract_social_from_html(page_html).items():
            if social[key] is None and value:
                social[key] = value

    # Homepage: try direct first, fall back to proxy only if direct found
    # nothing. dict.fromkeys dedupes the no-proxies-configured case
    # (proxy is already None) so direct isn't attempted twice. Keeps the
    # LAST candidate's status/error — if both direct and proxy fail, the
    # proxy attempt's classification is the more informative one to report
    # (direct failing is often just "no proxy configured" background noise).
    html: Optional[str] = None
    used_proxy: str | None = None
    failure_status: Optional[str] = None
    failure_error: Optional[str] = None
    for attempt_proxy in dict.fromkeys([None, proxy]):
        try:
            with _make_client(attempt_proxy) as client:
                attempt_html, attempt_status, attempt_error = _fetch_page(
                    client, base_url, proxy=attempt_proxy
                )
        except Exception as exc:
            logger.debug(
                "[CRAWL] %s connection failed for %s: %s",
                "Direct" if attempt_proxy is None else "Proxy", base_url, exc,
            )
            failure_status, failure_error = "temp_failed", str(exc)
            continue
        if attempt_html:
            html = attempt_html
            used_proxy = attempt_proxy
            break
        failure_status, failure_error = attempt_status, attempt_error

    if html is None:
        return [], [], social, failure_status, failure_error

    logger.debug("[CRAWL] %s — using %s", base_url, "proxy" if used_proxy else "direct")

    pages_visited.append(base_url)
    for item in extract_emails(html, base_url):
        all_emails.setdefault(item["email"], item | {"source_url": base_url})
    _merge_social(html)

    # Discover contact/about links from homepage
    discovered = _discover_contact_links(html, base_url)

    # Build full list: default paths + discovered, deduped
    domain = _base_domain(base_url)
    to_visit: list[str] = []
    seen_paths: set[str] = {base_url}

    for path in _DEFAULT_PATHS:
        url = domain + path
        if url not in seen_paths:
            to_visit.append(url)
            seen_paths.add(url)

    for url in discovered:
        if url not in seen_paths:
            to_visit.append(url)
            seen_paths.add(url)

    # Visit each sub-page, reusing whichever connection (direct or proxy)
    # the homepage fetch already proved works for this site. A sub-page's
    # own status/error is intentionally discarded — a dead /contact link
    # doesn't make the whole site 'blocked' or 'invalid_url' when the
    # homepage itself loaded fine.
    with _make_client(used_proxy) as client:
        for page_url in to_visit[:MAX_PAGES_PER_SITE]:
            time.sleep(random.uniform(0.5, 1.5))  # intra-site politeness
            sub_html, _sub_status, _sub_error = _fetch_page(client, page_url, proxy=used_proxy)
            if sub_html is None:
                continue
            pages_visited.append(page_url)
            for item in extract_emails(sub_html, page_url):
                all_emails.setdefault(item["email"], item | {"source_url": page_url})
            _merge_social(sub_html)

    return list(all_emails.values()), pages_visited, social, None, None


def _crawl_site_playwright(base_url: str) -> tuple[list[dict], dict]:
    """
    Playwright fallback: render homepage + /contact in a real browser
    session (via BrowserSession — shared stealth/fingerprint setup with the
    directory scrapers), for sites whose static HTML comes back empty
    because their content loads client-side. Only called if the static
    crawl yields zero emails. Also extracts social links from the rendered
    pages and returns them alongside the emails: a JS-rendered homepage
    that hides its emails from the static pass hides its social links from
    it too, so run_crawl() merges these into whatever the static pass found
    rather than discarding them.

    The whole session runs inside run_in_thread() — see its docstring in
    browser.py: the sync Playwright API raises if the calling thread
    already has a running asyncio event loop, which this sidesteps by
    always running in a fresh worker thread instead.
    """
    try:
        from playwright.sync_api import TimeoutError as PwTimeout
    except ImportError:
        logger.error("Playwright not installed — skipping dynamic fallback for %s", base_url)
        return [], {"facebook": None, "instagram": None, "linkedin": None}

    def _do_scrape() -> tuple[list[dict], dict]:
        all_emails: dict[str, dict] = {}
        social: dict[str, str | None] = {"facebook": None, "instagram": None, "linkedin": None}

        def _merge_social(html: str) -> None:
            for key, value in _extract_social_from_html(html).items():
                if social[key] is None and value:
                    social[key] = value

        with BrowserSession(headless=True) as session:
            page = session.new_page()

            for path_url in [base_url, base_url.rstrip("/") + "/contact"]:
                try:
                    page.goto(path_url, timeout=20_000, wait_until="networkidle")
                    dismiss_cookie_banner(page)
                    human_scroll(page, times=2)
                    html = page.content()
                    for item in extract_emails(html, path_url):
                        all_emails.setdefault(item["email"], item | {"source_url": path_url})
                    _merge_social(html)
                except PwTimeout:
                    logger.debug("[PW] Timeout loading %s", path_url)
                except Exception as exc:
                    logger.debug("[PW] Error loading %s: %s", path_url, exc)
                human_delay(800, 1500)

        return list(all_emails.values()), social

    try:
        return run_in_thread(_do_scrape)
    except Exception as exc:
        logger.error("[PW] Playwright session error for %s: %s", base_url, exc)
        return [], {"facebook": None, "instagram": None, "linkedin": None}


def run_crawl(db_path: str, niche: str | None = None, location: str | None = None) -> None:
    """
    Stage 3: crawl all resolved sites and extract emails.

    Processes businesses whose crawl_status is 'pending' or 'temp_failed' —
    the latter is a transient connection error/timeout from a previous run,
    worth retrying. 'blocked' (403/429 persisted through retries) and
    'invalid_url' (404/410) are deliberately excluded: retrying either
    wastes time, since neither fixes itself on a plain re-request the way
    a transient network error might.

    niche/location restrict this to one campaign's businesses — a plain
    `? IS NULL OR column=?` pair for each, so the default (both None)
    crawls everything pending/temp_failed, same as before this parameter
    existed.
    """
    logger.info("=== STAGE 3: CRAWL + EXTRACT ===")

    with get_conn(db_path) as conn:
        rows = conn.execute(
            """
            SELECT id, normalized_url, website_url
            FROM businesses
            WHERE resolve_status IN ('done','failed')
              AND crawl_status IN ('pending', 'temp_failed')
              AND (normalized_url IS NOT NULL OR website_url IS NOT NULL)
              AND (? IS NULL OR niche=?)
              AND (? IS NULL OR location=?)
            """,
            (niche, niche, location, location),
        ).fetchall()

    total = len(rows)
    logger.info("[CRAWL] %d sites to crawl", total)

    crawled = emails_found = playwright_fallbacks = 0
    status_counts: dict[str, int] = {}

    for i, row in enumerate(rows, 1):
        biz_id: int = row["id"]
        site_url: str = row["normalized_url"] or row["website_url"]

        logger.info("[CRAWL] (%d/%d) %s", i, total, site_url)

        try:
            emails, pages, social, failure_status, failure_error = _crawl_site_static(site_url)
        except Exception as exc:
            # Unexpected, unclassified failure — treat as temp_failed
            # (worth retrying) rather than a permanent block, since nothing
            # here actually diagnosed it as one.
            logger.error("[CRAWL] Unhandled error for %s: %s", site_url, exc)
            with get_conn(db_path) as conn:
                conn.execute(
                    "UPDATE businesses SET crawl_status='temp_failed', crawl_error=? WHERE id=?",
                    (str(exc), biz_id),
                )
            status_counts["temp_failed"] = status_counts.get("temp_failed", 0) + 1
            time.sleep(random.uniform(DELAY_MIN, DELAY_MAX))
            continue

        # If static crawl returned nothing, try Playwright once — a real
        # browser session can sometimes get past a 403 plain httpx can't.
        if not emails:
            logger.info("[CRAWL] Zero emails static for %s — trying Playwright", site_url)
            playwright_fallbacks += 1
            emails, pw_social = _crawl_site_playwright(site_url)
            for key, value in pw_social.items():
                if not social.get(key) and value:
                    social[key] = value

        if emails:
            crawl_status, crawl_error = "done", None
        elif failure_status:
            # The homepage itself was never reachable, even after the
            # Playwright retry above — keep the specific reason
            # _crawl_site_static already diagnosed, rather than flattening
            # it to a generic "no emails".
            crawl_status, crawl_error = failure_status, failure_error
        else:
            # Homepage loaded fine (by either path) but genuinely has no
            # email anywhere — a real "nothing to find", not a
            # connectivity problem.
            crawl_status, crawl_error = "no_emails", None

        if crawl_status == "blocked":
            logger.warning("[CRAWL] %s → blocked (%s) — skipping future retries", site_url, crawl_error)
        elif crawl_status == "temp_failed":
            logger.warning("[CRAWL] %s → temp_failed — will retry next run", site_url)
        elif crawl_status == "invalid_url":
            logger.warning("[CRAWL] %s → invalid_url (%s)", site_url, crawl_error)

        # Persist emails
        for item in emails:
            with get_conn(db_path) as conn:
                upsert_email(
                    conn,
                    business_id=biz_id,
                    email=item["email"],
                    source_url=item.get("source_url"),
                    extract_method=item.get("method", "regex"),
                )

        with get_conn(db_path) as conn:
            conn.execute(
                "UPDATE businesses SET crawl_status=?, crawl_error=? WHERE id=?",
                (crawl_status, crawl_error, biz_id),
            )

        # Opportunistically save any social links found while we already had
        # the pages fetched — SOCIAL stage will skip businesses this covers.
        if any(social.values()):
            with get_conn(db_path) as conn:
                upsert_social(
                    conn,
                    business_id=biz_id,
                    facebook_url=social.get("facebook"),
                    instagram_url=social.get("instagram"),
                    linkedin_url=social.get("linkedin"),
                )

        status_counts[crawl_status] = status_counts.get(crawl_status, 0) + 1
        emails_found += len(emails)
        crawled += 1
        logger.info(
            "[CRAWL] %s — pages=%d emails=%d status=%s",
            site_url, len(pages), len(emails), crawl_status,
        )

        time.sleep(random.uniform(DELAY_MIN, DELAY_MAX))

    logger.info(
        "[CRAWL] DONE — crawled=%d  done=%d  no_emails=%d  blocked=%d  temp_failed=%d  "
        "invalid_url=%d  emails_raw=%d  playwright_fallbacks=%d",
        crawled,
        status_counts.get("done", 0),
        status_counts.get("no_emails", 0),
        status_counts.get("blocked", 0),
        status_counts.get("temp_failed", 0),
        status_counts.get("invalid_url", 0),
        emails_found, playwright_fallbacks,
    )
