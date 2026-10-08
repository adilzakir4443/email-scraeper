"""
Stage — ENRICH (--enrich flag, bypasses the normal DISCOVER/.../WRITE pipeline).

Reads an existing Excel file the user already has, detects which of a fixed
set of columns are missing per row (by header name, not position), and
fills in ONLY what can be found and independently verified from real
sources — website crawl, Bing search, and (only once the free proxy-based
path has genuinely come up empty AND --api-fallback is set with keys in
.env) Google API. Never guesses: a field stays blank if nothing real was
found for it. Saves back to the same file (or "<name>_enriched_N.xlsx" if
that file is locked elsewhere), with every newly-filled cell highlighted
light green so the user can see exactly what changed.
"""

import logging
import random
import re
import time
import urllib.parse
from difflib import SequenceMatcher
from pathlib import Path
from typing import Optional

import httpx
import openpyxl
from openpyxl.cell.cell import MergedCell
from openpyxl.styles import PatternFill
from bs4 import BeautifulSoup

from .crawl import _crawl_site_static
from .proxy import get_pool
from .social import _bing_search_social, _FB_RE, _IG_RE, _LI_RE
from .social import _make_client, _fetch, _unwrap_bing_redirect

logger = logging.getLogger(__name__)

ENRICHED_FILL = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")

DIRECTORY_DOMAINS = frozenset([
    "yelp.com", "yellowpages.com", "facebook.com", "google.com", "bing.com",
    "tripadvisor.com", "yell.com", "thomsonlocal.com", "instagram.com",
    "linkedin.com", "twitter.com", "mapquest.com", "bbb.org", "manta.com",
    "foursquare.com", "superpages.com", "citysearch.com",
])

PHONE_RE = re.compile(r"(\+?1[\s.\-]?)?(\(?\d{3}\)?[\s.\-]?\d{3}[\s.\-]?\d{4})")
ADDRESS_RE = re.compile(
    r"\d+\s+[A-Za-z0-9\s]+(?:St|Ave|Rd|Blvd|Dr|Ln|Way|Court|Ct|Place|Pl|Suite|Ste)[\w\s,]*",
    re.IGNORECASE,
)

# Header text (case-insensitive, stripped) -> internal field name. Column
# order in the input file is never assumed — every column is located by its
# header text.
_COLUMN_ALIASES: dict[str, str] = {
    "company name": "business_name",
    "name": "business_name",
    "phone": "phone",
    "category": "category",
    "email": "email",
    "website": "website",
    "facebook": "facebook",
    "instagram": "instagram",
    "linkedin": "linkedin",
    "address": "address",
}

_PLATFORM_PATTERNS = {"facebook": _FB_RE, "instagram": _IG_RE, "linkedin": _LI_RE}

_EMAIL_RE = re.compile(r"[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}")
_EMAIL_FALSE_POSITIVE_MARKERS = ("sentry", "noreply", "no-reply", "example", "test")

_BROWSER_HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                  "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
}


# ---------------------------------------------------------------------------
# Small generic helpers
# ---------------------------------------------------------------------------

def _is_directory_domain(url: str) -> bool:
    """True if url's domain (or a subdomain of it) is a known directory/
    platform site, never the business's own site."""
    try:
        netloc = urllib.parse.urlparse(url).netloc.lower()
    except Exception:
        return False
    domain = netloc[4:] if netloc.startswith("www.") else netloc
    return any(domain == d or domain.endswith("." + d) for d in DIRECTORY_DOMAINS)


def _fuzzy_match(a: str, b: str) -> float:
    return SequenceMatcher(None, a.lower(), b.lower()).ratio()


def _social_url_plausible(url: str, business_name: str) -> bool:
    """
    Loose sanity check on a social-profile URL found via search: does its
    path (the username/slug) contain any significant word from the
    business name? A social platform's search result can't be verified by
    fetching+checking a title the way a normal website can (Facebook/
    Instagram/LinkedIn are typically behind bot walls for that) — confirmed
    live in earlier testing that without this, an entirely unrelated
    profile (a celebrity's Instagram) can slip through just because it
    matched a site: domain restriction. This cheap local check is the
    safety net instead.
    """
    try:
        path = urllib.parse.urlparse(url).path
    except Exception:
        return True
    slug = re.sub(r"[^a-z0-9]", "", path.lower())
    words = [w for w in re.findall(r"[a-z0-9]+", business_name.lower()) if len(w) >= 4]
    if not words or not slug:
        return True
    return any(word in slug for word in words)


def _extract_emails_from_results(results: list[dict]) -> Optional[str]:
    for result in results:
        text = f"{result.get('title', '')} {result.get('snippet', '')}"
        for match in _EMAIL_RE.finditer(text):
            email = match.group(0).lower()
            if any(marker in email for marker in _EMAIL_FALSE_POSITIVE_MARKERS):
                continue
            return email
    return None


def _format_phone(raw: str) -> str:
    digits = re.sub(r"\D", "", raw)
    if len(digits) == 11 and digits.startswith("1"):
        digits = digits[1:]
    if len(digits) == 10:
        return f"({digits[0:3]}) {digits[3:6]}-{digits[6:10]}"
    return raw.strip()


def _extract_phone_from_results(results: list[dict]) -> Optional[str]:
    for result in results:
        text = f"{result.get('title', '')} {result.get('snippet', '')}"
        match = PHONE_RE.search(text)
        if match:
            return _format_phone(match.group(0))
    return None


def _extract_address_from_results(results: list[dict]) -> Optional[str]:
    for result in results:
        text = f"{result.get('title', '')} {result.get('snippet', '')}"
        match = ADDRESS_RE.search(text)
        if match:
            return match.group(0).strip()
    return None


def _get_crawl_result(website: str, cache: dict) -> tuple[list[dict], dict]:
    """
    Crawl website via _crawl_site_static, caching the result per business
    row. Multiple fields (email, facebook, instagram, linkedin) can each
    need data from the same site within one row — without this cache each
    would trigger its own separate crawl of the same pages.
    """
    if website not in cache:
        try:
            emails, _pages, social_links = _crawl_site_static(website)
        except Exception as exc:
            logger.debug("[ENRICH] Crawl failed for %s: %s", website, exc)
            emails, social_links = [], {}
        cache[website] = (emails, social_links)
    return cache[website]


# ---------------------------------------------------------------------------
# Bing search
# ---------------------------------------------------------------------------

def _bing_search(query: str, proxy: Optional[str]) -> list[dict]:
    """
    Run a Bing search and return organic results as {url, title, snippet}
    dicts. Decodes Bing's ck/a redirect wrapper on every result link (see
    social._unwrap_bing_redirect's docstring — confirmed live that a bare
    href on a Bing results page is a useless redirect, never the actual
    destination), so 'url' is always the real target. Returns [] on error.
    """
    url = f"https://www.bing.com/search?q={urllib.parse.quote_plus(query)}"
    try:
        with _make_client(proxy) as client:
            html = _fetch(client, url, proxy=proxy)
    except Exception as exc:
        logger.debug("[ENRICH] Bing search failed for %r: %s", query, exc)
        return []

    if not html:
        return []

    try:
        soup = BeautifulSoup(html, "lxml")
        results = []
        for result in soup.select("li.b_algo"):
            link = result.select_one("h2 a[href]")
            if not link:
                continue
            real_url = _unwrap_bing_redirect(link["href"])
            if not real_url or not real_url.startswith("http"):
                continue
            title = link.get_text(" ", strip=True)
            snippet_tag = result.select_one(".b_caption") or result.select_one("p")
            snippet = snippet_tag.get_text(" ", strip=True) if snippet_tag else ""
            results.append({"url": real_url, "title": title, "snippet": snippet})
        return results
    except Exception as exc:
        logger.debug("[ENRICH] Bing result parsing failed for %r: %s", query, exc)
        return []


# ---------------------------------------------------------------------------
# Website
# ---------------------------------------------------------------------------

def _verify_website(url: str, business_name: str) -> bool:
    """
    Never trust a search result on its own: fetch the candidate page
    directly, confirm it responds 200, and confirm its <title> contains at
    least one word from business_name (fuzzy match, threshold 0.4).
    """
    try:
        resp = httpx.get(url, timeout=15, headers=_BROWSER_HEADERS, follow_redirects=True)
        if resp.status_code != 200:
            return False
        soup = BeautifulSoup(resp.text, "lxml")
        title = soup.title.get_text() if soup.title else ""
    except Exception as exc:
        logger.debug("[ENRICH] Website verify failed for %s: %s", url, exc)
        return False

    title_words = re.findall(r"[a-z0-9]+", title.lower())
    name_words = [w for w in re.findall(r"[a-z0-9]+", business_name.lower()) if len(w) > 2]
    if not title_words or not name_words:
        return False
    return any(_fuzzy_match(nw, tw) >= 0.4 for nw in name_words for tw in title_words)


def _find_website(
    business_name: str,
    city: str,
    proxy: Optional[str],
    api_fallback: bool,
    google_api_key: Optional[str],
    google_cx: Optional[str],
) -> Optional[str]:
    # Step 1: proxy-based Bing search
    results = _bing_search(f'"{business_name}" "{city}" official site', proxy)
    for result in results:
        url = result.get("url", "")
        if not url or _is_directory_domain(url):
            continue
        if _verify_website(url, business_name):
            return url

    # Step 2: Google API fallback — only once Bing has genuinely found
    # nothing real; never used opportunistically (costs API quota).
    if api_fallback and google_api_key and google_cx:
        from .google_api import find_website_via_cse
        url = find_website_via_cse(business_name, city, google_api_key, google_cx)
        if url and _verify_website(url, business_name):
            return url

    return None


# ---------------------------------------------------------------------------
# Email
# ---------------------------------------------------------------------------

def _find_email(
    business_name: str,
    city: str,
    website: Optional[str],
    proxy: Optional[str],
    api_fallback: bool,
    google_api_key: Optional[str],
    google_cx: Optional[str],
    crawl_cache: Optional[dict] = None,
) -> Optional[str]:
    if crawl_cache is None:
        crawl_cache = {}

    # Step 1: crawl the website, if we have one — prefer a non-role,
    # non-freemail address among whatever was found.
    if website:
        crawled_emails, _social = _get_crawl_result(website, crawl_cache)
        if crawled_emails:
            from .verify import _is_role, _is_freemail

            def _rank(item: dict) -> tuple[bool, bool]:
                local, _, domain = item["email"].partition("@")
                return (_is_role(local), _is_freemail(domain))

            best = sorted(crawled_emails, key=_rank)[0]
            return best["email"]

    # Step 2: Bing search
    results = _bing_search(f'"{business_name}" "{city}" email contact', proxy)
    found = _extract_emails_from_results(results)
    if found:
        return found

    # Step 3: Google API fallback
    if api_fallback and google_api_key and google_cx:
        from .google_api import google_custom_search
        results = google_custom_search(
            f'"{business_name}" "{city}" email contact', google_api_key, google_cx
        )
        found = _extract_emails_from_results(results)
        if found:
            return found

    return None


# ---------------------------------------------------------------------------
# Phone
# ---------------------------------------------------------------------------

def _find_phone(
    business_name: str,
    city: str,
    proxy: Optional[str],
    api_fallback: bool,
    google_api_key: Optional[str],
    google_cx: Optional[str],
) -> Optional[str]:
    results = _bing_search(f'"{business_name}" "{city}" phone', proxy)
    found = _extract_phone_from_results(results)
    if found:
        return found

    if api_fallback and google_api_key and google_cx:
        from .google_api import google_custom_search
        results = google_custom_search(
            f'"{business_name}" "{city}" phone number', google_api_key, google_cx
        )
        found = _extract_phone_from_results(results)
        if found:
            return found

    return None


# ---------------------------------------------------------------------------
# Social (Facebook / Instagram / LinkedIn)
# ---------------------------------------------------------------------------

def _find_social(
    business_name: str,
    city: str,
    website: Optional[str],
    platform: str,
    proxy: Optional[str],
    api_fallback: bool,
    google_api_key: Optional[str],
    google_cx: Optional[str],
    crawl_cache: Optional[dict] = None,
) -> Optional[str]:
    if crawl_cache is None:
        crawl_cache = {}

    # Step 1: the business's own website — most trustworthy source, no
    # plausibility check needed since it's literally linked from the site.
    if website:
        _emails, social_links = _get_crawl_result(website, crawl_cache)
        found = social_links.get(platform)
        if found:
            return found

    # Step 2: existing Bing-based social search from social.py
    found = _bing_search_social(business_name, city, platform)
    if found and _social_url_plausible(found, business_name):
        return found

    # Step 3: Google API fallback
    if api_fallback and google_api_key and google_cx:
        from .google_api import google_custom_search
        pattern = _PLATFORM_PATTERNS[platform]
        results = google_custom_search(
            f'"{business_name}" "{city}" {platform}', google_api_key, google_cx
        )
        for item in results:
            link = item.get("link", "")
            if link and pattern.search(link) and _social_url_plausible(link, business_name):
                return link

    return None


# ---------------------------------------------------------------------------
# Address
# ---------------------------------------------------------------------------

def _find_address(
    business_name: str,
    city: str,
    proxy: Optional[str],
    api_fallback: bool,
    google_api_key: Optional[str],
    google_cx: Optional[str],
) -> Optional[str]:
    results = _bing_search(f'"{business_name}" "{city}" address', proxy)
    found = _extract_address_from_results(results)
    if found:
        return found

    # A Places Details call (mentioned as an alternative Step 2 source)
    # needs a place_id, which nothing earlier in this flow produces for an
    # arbitrary spreadsheet row — a CSE query is the only fallback actually
    # reachable with the inputs available here.
    if api_fallback and google_api_key and google_cx:
        from .google_api import google_custom_search
        results = google_custom_search(
            f'"{business_name}" "{city}" address', google_api_key, google_cx
        )
        found = _extract_address_from_results(results)
        if found:
            return found

    return None


# ---------------------------------------------------------------------------
# Read / write phases
# ---------------------------------------------------------------------------

def _detect_columns(header_row) -> dict[str, int]:
    """Map internal field name -> 0-based column index, located by header
    text (case-insensitive, stripped) rather than an assumed column order."""
    mapping: dict[str, int] = {}
    for idx, cell in enumerate(header_row):
        text = str(cell.value or "").strip().lower()
        field = _COLUMN_ALIASES.get(text)
        if field and field not in mapping:
            mapping[field] = idx
    return mapping


def _save_enriched(wb, input_path: str) -> str:
    """Save back to input_path; if it's locked elsewhere, fall back to
    "<name>_enriched_N.xlsx" (N = 1, 2, 3, ... — first path that doesn't
    already exist) rather than losing the run's results."""
    in_file = Path(input_path)
    try:
        wb.save(str(in_file))
        return str(in_file)
    except PermissionError:
        stem = in_file.stem
        suffix = in_file.suffix
        counter = 1
        while True:
            new_path = in_file.parent / f"{stem}_enriched_{counter}{suffix}"
            if not new_path.exists():
                break
            counter += 1
        wb.save(str(new_path))
        logger.warning("[ENRICH] File locked — saved to %s", new_path)
        return str(new_path)


def run_enrich(
    input_path: str,
    niche: str,
    location: str,
    api_fallback: bool = False,
    google_api_key: Optional[str] = None,
    google_cx: Optional[str] = None,
    google_places_key: Optional[str] = None,
) -> str:
    """
    Read an existing Excel file, detect which of the supported columns are
    missing per row, and fill in only what can be found and verified from
    real sources — proxy-based search always runs first; Google API is
    only tried once that's genuinely come up empty AND api_fallback is set
    with keys configured. Returns the path actually saved to.

    google_places_key is accepted for call-site consistency with the other
    stages (main() passes all three Google keys through uniformly) but
    nothing in this stage's defined behavior uses Places specifically.
    """
    logger.info("=== STAGE: ENRICH  input=%r ===", input_path)

    wb = openpyxl.load_workbook(input_path)
    ws = wb.active

    header_row = next(ws.iter_rows(min_row=1, max_row=1))
    columns = _detect_columns(header_row)

    if "business_name" not in columns:
        raise ValueError(
            f"No 'Company Name' column found in {input_path} — cannot enrich without it"
        )

    city = location.split(",")[0].strip() if location else ""

    pool = get_pool()
    proxy = pool.get()

    total_rows = 0
    enriched_rows = 0
    fields_filled = 0

    data_rows = list(ws.iter_rows(min_row=2))
    total_data_rows = len(data_rows)

    for i, row_cells in enumerate(data_rows, 1):
        values: dict[str, str] = {}
        for field, col_idx in columns.items():
            if col_idx >= len(row_cells):
                continue
            cell = row_cells[col_idx]
            values[field] = str(cell.value).strip() if cell.value is not None else ""

        name = values.get("business_name", "").strip()
        if not name:
            continue  # cannot enrich without a name

        total_rows += 1
        row_changed = False
        website = values.get("website", "").strip()
        crawl_cache: dict[str, tuple] = {}

        logger.info("[ENRICH] (%d/%d) Processing: %s", i, total_data_rows, name)

        def _write_field(field: str, value: str) -> None:
            nonlocal row_changed, fields_filled
            col_idx = columns.get(field)
            if col_idx is None or col_idx >= len(row_cells):
                return
            cell = row_cells[col_idx]
            if isinstance(cell, MergedCell):
                # MergedCell.value is read-only in openpyxl (AttributeError
                # on assignment) — only the merge's top-left anchor cell is
                # writable, which lives in a different column than the one
                # this field maps to, so there's no safe cell to write this
                # value into. Skip rather than crash the whole run over one
                # row's leftover formatting.
                logger.warning(
                    "[ENRICH] Skipping %s for %s — cell is part of a merged range",
                    field, name,
                )
                return
            cell.value = value
            cell.fill = ENRICHED_FILL
            row_changed = True
            fields_filled += 1

        # 1. Website
        if "website" in columns and not website:
            try:
                found = _find_website(name, city, proxy, api_fallback, google_api_key, google_cx)
            except Exception as exc:
                logger.warning("[ENRICH] Website search failed for %s: %s", name, exc)
                found = None
            if found:
                _write_field("website", found)
                website = found
                logger.info("[ENRICH] Website → %s", found)
            time.sleep(random.uniform(1.0, 2.0))

        # 2. Email
        if "email" in columns and not values.get("email", "").strip():
            try:
                found = _find_email(
                    name, city, website, proxy, api_fallback, google_api_key, google_cx, crawl_cache
                )
            except Exception as exc:
                logger.warning("[ENRICH] Email search failed for %s: %s", name, exc)
                found = None
            if found:
                _write_field("email", found)
                logger.info("[ENRICH] Email → %s", found)
            else:
                logger.info("[ENRICH] No email found for %s", name)
            time.sleep(random.uniform(1.0, 2.0))

        # 3. Phone
        if "phone" in columns and not values.get("phone", "").strip():
            try:
                found = _find_phone(name, city, proxy, api_fallback, google_api_key, google_cx)
            except Exception as exc:
                logger.warning("[ENRICH] Phone search failed for %s: %s", name, exc)
                found = None
            if found:
                _write_field("phone", found)
                logger.info("[ENRICH] Phone → %s", found)
            time.sleep(random.uniform(1.0, 2.0))

        # 4-6. Social
        for platform in ("facebook", "instagram", "linkedin"):
            if platform in columns and not values.get(platform, "").strip():
                try:
                    found = _find_social(
                        name, city, website, platform, proxy,
                        api_fallback, google_api_key, google_cx, crawl_cache,
                    )
                except Exception as exc:
                    logger.warning(
                        "[ENRICH] %s search failed for %s: %s", platform.title(), name, exc
                    )
                    found = None
                if found:
                    _write_field(platform, found)
                time.sleep(random.uniform(0.5, 1.5))

        # 7. Address
        if "address" in columns and not values.get("address", "").strip():
            try:
                found = _find_address(name, city, proxy, api_fallback, google_api_key, google_cx)
            except Exception as exc:
                logger.warning("[ENRICH] Address search failed for %s: %s", name, exc)
                found = None
            if found:
                _write_field("address", found)
            time.sleep(random.uniform(1.0, 2.0))

        if row_changed:
            enriched_rows += 1

        time.sleep(random.uniform(2.0, 4.0))

    out_path = _save_enriched(wb, input_path)

    logger.info(
        "[ENRICH] DONE — processed=%d enriched=%d fields_filled=%d",
        total_rows, enriched_rows, fields_filled,
    )
    return out_path
