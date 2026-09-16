"""
Stage 3b — SOCIAL.

For each business:
  1. If it has a resolved (or raw) website, fetch its homepage + /contact
     page and look for Facebook/Instagram/LinkedIn profile links.
  2. If it has no website, or its site yielded nothing, fall back to a Bing
     web search restricted to each platform's domain and take the first
     matching organic result.

Resumable: social_status tracks pending/done. Businesses already marked
'done' (by this stage, or opportunistically by CRAWL while it was already
fetching pages for emails) are skipped.
"""

import base64
import logging
import random
import re
import time
import urllib.parse

import httpx
from bs4 import BeautifulSoup

from .db import get_conn, upsert_social
from .proxy import get_pool

logger = logging.getLogger(__name__)

DELAY_MIN = 3.0
DELAY_MAX = 8.0

_FB_RE = re.compile(r'https?://(?:www\.)?facebook\.com/[^\s"\'<>]+', re.IGNORECASE)
_IG_RE = re.compile(r'https?://(?:www\.)?instagram\.com/[^\s"\'<>]+', re.IGNORECASE)
_LI_RE = re.compile(r'https?://(?:www\.)?linkedin\.com/[^\s"\'<>]+', re.IGNORECASE)

# Facebook path segments that are never a real business page — share dialogs,
# tracking pixels, embeddable widgets, hashtag pages, etc. Matched against the
# FIRST PATH SEGMENT only, not as a substring, so a page genuinely named e.g.
# "facebook.com/SharersDelight" is never wrongly filtered out.
_FB_IGNORE_PATHS = frozenset(["sharer", "share", "dialog", "plugins", "tr", "hashtag"])

USER_AGENTS = [
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/17.4 Safari/605.1.15",
    "Mozilla/5.0 (X11; Linux x86_64; rv:125.0) Gecko/20100101 Firefox/125.0",
]

_PLATFORM_DOMAINS = {
    "facebook": "facebook.com",
    "instagram": "instagram.com",
    "linkedin": "linkedin.com",
}

_PLATFORM_PATTERNS = {
    "facebook": _FB_RE,
    "instagram": _IG_RE,
    "linkedin": _LI_RE,
}


def _random_headers() -> dict[str, str]:
    return {
        "User-Agent": random.choice(USER_AGENTS),
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
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


def _fetch(client: httpx.Client, url: str) -> str | None:
    """Fetch a single page and return its HTML, or None on error."""
    pool = get_pool()
    for attempt in range(3):
        try:
            resp = client.get(url)
            if resp.status_code in (403, 429):
                logger.debug("HTTP %d for %s (attempt %d)", resp.status_code, url, attempt)
                pool.backoff_sleep(attempt + 1)
                continue
            if resp.status_code == 200:
                return resp.text
            logger.debug("HTTP %d for %s — skipping", resp.status_code, url)
            return None
        except (httpx.ConnectError, httpx.TimeoutException, httpx.RemoteProtocolError) as exc:
            logger.debug("Fetch error %s: %s (attempt %d)", url, exc, attempt)
            pool.backoff_sleep(attempt + 1)
    return None


def _is_ignored_fb_path(url: str) -> bool:
    """True if the first path segment of *url* is a known non-business path
    (share dialogs, tracking pixels, widgets, hashtags) — checked as a whole
    path segment, never a substring."""
    try:
        path = urllib.parse.urlparse(url).path
    except Exception:
        return False
    segments = [s for s in path.split("/") if s]
    if not segments:
        return False
    first = segments[0].split(".")[0].lower()
    return first in _FB_IGNORE_PATHS


def _unwrap_bing_redirect(url: str) -> str:
    """
    Bing wraps every organic-result link in a bing.com/ck/a redirect whose
    real destination is base64-encoded in the "u" query parameter (prefixed
    with a 2-character marker, commonly "a1"), e.g.:
        https://www.bing.com/ck/a?...&u=a1aHR0cHM6Ly9wbGF5ZXJvay5jb20v&ntb=1
        -> https://playerok.com/
    A raw <a href> on a Bing results page is essentially always this
    redirect form rather than the bare destination URL, so without decoding
    it, pattern-matching against the destination domain would never succeed
    even for a fully relevant result. Returns *url* unchanged if it isn't a
    Bing redirect, or if it can't be decoded.
    """
    try:
        parsed = urllib.parse.urlparse(url)
        if "bing.com" not in parsed.netloc or not parsed.path.endswith("/ck/a"):
            return url
        encoded = urllib.parse.parse_qs(parsed.query).get("u", [None])[0]
        if not encoded or len(encoded) < 3:
            return url
        payload = encoded[2:]  # strip the leading marker (e.g. "a1")
        padded = payload + "=" * (-len(payload) % 4)
        decoded = base64.urlsafe_b64decode(padded).decode("utf-8", errors="ignore")
        return decoded if decoded.startswith("http") else url
    except Exception:
        return url


def _reconstruct_from_cite(cite_text: str, domain: str) -> str | None:
    """
    Bing shows each result's URL as breadcrumb text (e.g. a <cite> tag)
    separate from its href, formatted like
    "https://www.facebook.com › pagename" instead of a real path. Used as a
    last-resort fallback when a result's href can't be decoded — best-effort
    only, since Bing may abbreviate or alter the displayed path.
    """
    if domain not in cite_text.lower():
        return None
    text = cite_text.strip()
    for prefix in ("https://", "http://"):
        if text.lower().startswith(prefix):
            text = text[len(prefix):]
            break
    parts = [p.strip() for p in text.replace("›", "/").split("/") if p.strip()]
    if not parts:
        return None
    host = parts[0]
    path = "/".join(parts[1:])
    return f"https://{host}/{path}".rstrip("/") if path else f"https://{host}"


def _first_valid_match(
    pattern: "re.Pattern[str]",
    hrefs: list[str],
    html: str,
    ignore=None,
) -> str | None:
    """Prefer a match found in an actual <a href>; fall back to scanning the
    raw HTML (links embedded in scripts/attributes rather than anchors)."""
    for href in hrefs:
        if pattern.search(href) and (ignore is None or not ignore(href)):
            return href
    for m in pattern.finditer(html):
        url = m.group(0)
        if ignore is None or not ignore(url):
            return url
    return None


def _extract_social_from_html(html: str) -> dict[str, str | None]:
    """Scan a page's HTML for Facebook/Instagram/LinkedIn profile links.
    Returns {"facebook": url|None, "instagram": url|None, "linkedin": url|None}.
    """
    soup = BeautifulSoup(html, "lxml")
    hrefs = [a["href"] for a in soup.find_all("a", href=True)]

    return {
        "facebook": _first_valid_match(_FB_RE, hrefs, html, _is_ignored_fb_path),
        "instagram": _first_valid_match(_IG_RE, hrefs, html),
        "linkedin": _first_valid_match(_LI_RE, hrefs, html),
    }


def _extract_from_website(site_url: str) -> dict[str, str | None]:
    """Fetch a business's homepage + /contact page and look for social links."""
    pool = get_pool()
    proxy = pool.get()

    result: dict[str, str | None] = {"facebook": None, "instagram": None, "linkedin": None}

    parsed = urllib.parse.urlparse(site_url)
    domain = f"{parsed.scheme}://{parsed.netloc}"
    pages = [site_url]
    contact_url = domain.rstrip("/") + "/contact"
    if contact_url not in pages:
        pages.append(contact_url)

    with _make_client(proxy) as client:
        for page_url in pages:
            html = _fetch(client, page_url)
            if not html:
                continue
            found = _extract_social_from_html(html)
            for key, value in found.items():
                if result[key] is None and value:
                    result[key] = value
            if all(result.values()):
                break
            time.sleep(random.uniform(0.5, 1.5))

    return result


def _bing_search_social(business_name: str, location: str, platform: str) -> str | None:
    """
    Search Bing for business_name + location restricted to *platform*'s
    domain (via site:) and return the first matching organic-result link.
    """
    pool = get_pool()
    proxy = pool.get()

    domain = _PLATFORM_DOMAINS[platform]
    pattern = _PLATFORM_PATTERNS[platform]

    query = urllib.parse.quote_plus(f"{business_name} {location} site:{domain}")
    url = f"https://www.bing.com/search?q={query}"

    with _make_client(proxy) as client:
        html = _fetch(client, url)

    if not html:
        return None

    def _valid(candidate: str) -> bool:
        return not (platform == "facebook" and _is_ignored_fb_path(candidate))

    soup = BeautifulSoup(html, "lxml")
    for result in soup.select("li.b_algo"):
        link = result.select_one("h2 a[href]")
        if not link:
            continue
        real_url = _unwrap_bing_redirect(link["href"])
        if pattern.search(real_url) and _valid(real_url):
            return real_url

        # The href may not have decoded cleanly (format drift, ad slot,
        # etc.) — fall back to reconstructing from the visible URL
        # breadcrumb, which shows the real domain even when the href doesn't.
        cite = result.select_one("cite")
        if cite:
            reconstructed = _reconstruct_from_cite(cite.get_text(), domain)
            if reconstructed and pattern.search(reconstructed) and _valid(reconstructed):
                return reconstructed

    return None


def _find_social_no_website(business_name: str, location: str) -> dict[str, str | None]:
    """Search Bing for all three platforms — used when a business has no
    website, or its website yielded no social links."""
    result: dict[str, str | None] = {}
    for platform in ("facebook", "instagram", "linkedin"):
        result[platform] = _bing_search_social(business_name, location, platform)
        time.sleep(random.uniform(1.0, 2.0))
    return result


def run_social(db_path: str) -> None:
    """
    Stage 3b: find Facebook/Instagram/LinkedIn links for each business.
    Skips businesses whose social_status is already 'done'.
    """
    logger.info("=== STAGE 3b: SOCIAL ===")

    with get_conn(db_path) as conn:
        rows = conn.execute(
            """
            SELECT id, business_name, location, normalized_url, website_url
            FROM businesses
            WHERE social_status != 'done'
            """
        ).fetchall()

    total = len(rows)
    logger.info("[SOCIAL] %d businesses to process", total)

    found_counts = {"facebook": 0, "instagram": 0, "linkedin": 0}
    processed = 0

    for i, row in enumerate(rows, 1):
        biz_id: int = row["id"]
        site_url = row["normalized_url"] or row["website_url"]
        business_name: str = row["business_name"]
        location: str = row["location"]

        logger.info("[SOCIAL] (%d/%d) %s", i, total, business_name)

        try:
            social = _extract_from_website(site_url) if site_url else None
            if not social or not any(social.values()):
                social = _find_social_no_website(business_name, location)
        except Exception as exc:
            logger.error("[SOCIAL] Unhandled error for %s: %s", business_name, exc)
            social = {"facebook": None, "instagram": None, "linkedin": None}

        with get_conn(db_path) as conn:
            upsert_social(
                conn,
                business_id=biz_id,
                facebook_url=social.get("facebook"),
                instagram_url=social.get("instagram"),
                linkedin_url=social.get("linkedin"),
            )

        for platform in found_counts:
            if social.get(platform):
                found_counts[platform] += 1

        processed += 1
        time.sleep(random.uniform(DELAY_MIN, DELAY_MAX))

    logger.info(
        "[SOCIAL] DONE — processed=%d  facebook=%d  instagram=%d  linkedin=%d",
        processed, found_counts["facebook"], found_counts["instagram"], found_counts["linkedin"],
    )
