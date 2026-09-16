"""
Stage 3 — CRAWL + EXTRACT.

For each resolved site:
  1. Fetch homepage + /contact, /about, /team + same-domain contact/about links
  2. Extract emails from each page via extract.py
  3. If zero emails found statically → retry ONCE with Playwright Chromium
  4. Store extracted emails in the emails table

Resumable: crawl_status tracks done/failed/no_emails_static/pending.
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
    return httpx.Client(
        proxies=proxies,
        follow_redirects=True,
        timeout=httpx.Timeout(20.0),
        verify=True,
        headers=_random_headers(),
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


def _fetch_page(client: httpx.Client, url: str) -> Optional[str]:
    """Fetch a single page and return its HTML, or None on error."""
    pool = get_pool()
    for attempt in range(3):
        try:
            resp = client.get(url)
            if resp.status_code in (403, 429):
                logger.warning("HTTP %d for %s (attempt %d)", resp.status_code, url, attempt)
                pool.backoff_sleep(attempt + 1)
                continue
            if resp.status_code == 200:
                return resp.text
            # 404/410 etc — page doesn't exist, stop trying
            logger.debug("HTTP %d for %s — skipping", resp.status_code, url)
            return None
        except (httpx.ConnectError, httpx.TimeoutException, httpx.RemoteProtocolError) as exc:
            logger.debug("Fetch error %s: %s (attempt %d)", url, exc, attempt)
            pool.backoff_sleep(attempt + 1)
    return None


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


def _crawl_site_static(base_url: str) -> tuple[list[dict], list[str], dict]:
    """
    Fetch homepage + standard paths + discovered contact links.
    Returns (emails_list, pages_visited, social_links).
    social_links is {"facebook": url|None, "instagram": url|None, "linkedin": url|None},
    merged across every page visited (first match per platform wins).
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

    with _make_client(proxy) as client:
        # Homepage first
        html = _fetch_page(client, base_url)
        if html is None:
            return [], [], social

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

        # Visit each sub-page
        for page_url in to_visit[:MAX_PAGES_PER_SITE]:
            time.sleep(random.uniform(0.5, 1.5))  # intra-site politeness
            sub_html = _fetch_page(client, page_url)
            if sub_html is None:
                continue
            pages_visited.append(page_url)
            for item in extract_emails(sub_html, page_url):
                all_emails.setdefault(item["email"], item | {"source_url": page_url})
            _merge_social(sub_html)

    return list(all_emails.values()), pages_visited, social


def _crawl_site_playwright(base_url: str) -> list[dict]:
    """
    Playwright fallback: render homepage + /contact in headless Chromium.
    Only called if static crawl yields zero emails.
    """
    try:
        from playwright.sync_api import sync_playwright, TimeoutError as PwTimeout
    except ImportError:
        logger.error("Playwright not installed — skipping dynamic fallback for %s", base_url)
        return []

    all_emails: dict[str, dict] = {}

    try:
        with sync_playwright() as pw:
            browser = pw.chromium.launch(
                headless=True,
                args=["--no-sandbox", "--disable-dev-shm-usage"],
            )
            context = browser.new_context(
                user_agent=random.choice(USER_AGENTS),
                java_script_enabled=True,
                ignore_https_errors=True,
            )
            page = context.new_page()

            for path_url in [base_url, base_url.rstrip("/") + "/contact"]:
                try:
                    page.goto(path_url, timeout=20_000, wait_until="networkidle")
                    html = page.content()
                    for item in extract_emails(html, path_url):
                        all_emails.setdefault(item["email"], item | {"source_url": path_url})
                except PwTimeout:
                    logger.debug("[PW] Timeout loading %s", path_url)
                except Exception as exc:
                    logger.debug("[PW] Error loading %s: %s", path_url, exc)
                time.sleep(1.0)

            context.close()
            browser.close()
    except Exception as exc:
        logger.error("[PW] Playwright session error for %s: %s", base_url, exc)

    return list(all_emails.values())


def run_crawl(db_path: str) -> None:
    """
    Stage 3: crawl all resolved sites and extract emails.
    Skips businesses whose crawl_status is not 'pending'.
    """
    logger.info("=== STAGE 3: CRAWL + EXTRACT ===")

    with get_conn(db_path) as conn:
        rows = conn.execute(
            """
            SELECT id, normalized_url, website_url
            FROM businesses
            WHERE resolve_status IN ('done','failed')
              AND crawl_status = 'pending'
              AND (normalized_url IS NOT NULL OR website_url IS NOT NULL)
            """
        ).fetchall()

    total = len(rows)
    logger.info("[CRAWL] %d sites to crawl", total)

    crawled = failed = emails_found = playwright_fallbacks = 0

    for i, row in enumerate(rows, 1):
        biz_id: int = row["id"]
        site_url: str = row["normalized_url"] or row["website_url"]

        logger.info("[CRAWL] (%d/%d) %s", i, total, site_url)

        try:
            emails, pages, social = _crawl_site_static(site_url)
        except Exception as exc:
            logger.error("[CRAWL] Unhandled error for %s: %s", site_url, exc)
            with get_conn(db_path) as conn:
                conn.execute(
                    "UPDATE businesses SET crawl_status='failed' WHERE id=?", (biz_id,)
                )
            failed += 1
            time.sleep(random.uniform(DELAY_MIN, DELAY_MAX))
            continue

        # If static crawl returned nothing, try Playwright once
        if not emails:
            logger.info("[CRAWL] Zero emails static for %s — trying Playwright", site_url)
            playwright_fallbacks += 1
            emails = _crawl_site_playwright(site_url)
            crawl_status = "done" if emails else "no_emails_static"
        else:
            crawl_status = "done"

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
                "UPDATE businesses SET crawl_status=? WHERE id=?",
                (crawl_status, biz_id),
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

        emails_found += len(emails)
        crawled += 1
        logger.info(
            "[CRAWL] %s — pages=%d emails=%d status=%s",
            site_url, len(pages), len(emails), crawl_status,
        )

        time.sleep(random.uniform(DELAY_MIN, DELAY_MAX))

    logger.info(
        "[CRAWL] DONE — crawled=%d  failed=%d  emails_raw=%d  playwright_fallbacks=%d",
        crawled, failed, emails_found, playwright_fallbacks,
    )
