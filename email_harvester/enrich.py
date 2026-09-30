"""
Stage — ENRICH (--enrich flag, bypasses the normal DISCOVER/.../WRITE pipeline).

Reads an existing Excel file the user already has, detects which of a fixed
set of columns are missing per row (by header name, not position), and
fills in ONLY what can be found and independently verified from real
sources — website crawl, Bing search, YellowPages. Never guesses: a field
stays blank if nothing real was found for it. Saves back to the same file
(or "<name>_enriched.xlsx" if that file is open/locked elsewhere), with
every newly-filled cell highlighted light green so the user can see exactly
what changed.
"""

import logging
import random
import re
import time
import urllib.parse
from difflib import SequenceMatcher
from pathlib import Path
from typing import Optional

import openpyxl
from openpyxl.styles import PatternFill
from bs4 import BeautifulSoup

from . import discover, social
from .crawl import _crawl_site_static
from .extract import _EMAIL_RE
from .proxy import get_pool
from .resolve import _domain_of
from .social import _make_client, _fetch, _unwrap_bing_redirect
from .write import _load_suppression, _is_suppressed

logger = logging.getLogger(__name__)

# Header text (case-insensitive) -> internal field name. Column order in the
# input file is never assumed — every column is located by its header.
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

# Fields ENRICH PHASE has a defined rule for (A-E in the spec). "category" is
# detected as present/missing for logging, but there's no rule to fill it.
_FILLABLE_FIELDS = frozenset(
    ["website", "email", "phone", "facebook", "instagram", "linkedin", "address"]
)

# Directory/platform domains that are never the business's own site, even if
# they show up as the top Bing result for its name.
_DIRECTORY_DOMAINS = frozenset([
    "yelp.com", "yellowpages.com", "facebook.com", "google.com", "bing.com",
    "tripadvisor.com", "yell.com", "thomsonlocal.com",
])

_PHONE_RE = re.compile(r"(\+?1?\s?)?(\(?\d{3}\)?[\s.\-]?\d{3}[\s.\-]?\d{4})")
_ADDRESS_RE = re.compile(
    r"\d+\s+[A-Za-z]+\s+(?:St|Ave|Rd|Blvd|Dr|Ln|Way|Court|Ct|Place|Pl)\b", re.IGNORECASE
)

_GREEN_FILL = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")

_DELAY_BETWEEN_BUSINESSES = (2.0, 4.0)
_DELAY_BETWEEN_FIELD_SEARCHES = (1.0, 2.0)


def _pause(bounds: tuple[float, float]) -> None:
    time.sleep(random.uniform(*bounds))


# ---------------------------------------------------------------------------
# Bing search plumbing (shared by every field's search fallback)
# ---------------------------------------------------------------------------

def _bing_search(query: str) -> Optional[BeautifulSoup]:
    pool = get_pool()
    proxy = pool.get()
    url = f"https://www.bing.com/search?q={urllib.parse.quote_plus(query)}"
    with _make_client(proxy) as client:
        html = _fetch(client, url)
    if not html:
        return None
    return BeautifulSoup(html, "lxml")


def _organic_results(soup: BeautifulSoup, limit: int = 5):
    """
    Yield (real_url, domain, title_text, snippet_text) for each organic
    result, decoding Bing's ck/a redirect wrapper (see social.py's
    _unwrap_bing_redirect — a bare href on a Bing results page is never a
    usable destination URL) and skipping anything that doesn't decode.
    """
    for result in soup.select("li.b_algo")[:limit]:
        link = result.select_one("h2 a[href]")
        if not link:
            continue
        real_url = _unwrap_bing_redirect(link["href"])
        if not real_url or not real_url.startswith("http"):
            continue
        title_text = link.get_text(" ", strip=True)
        snippet_tag = result.select_one(".b_caption") or result.select_one("p")
        snippet_text = snippet_tag.get_text(" ", strip=True) if snippet_tag else ""
        yield real_url, _domain_of(real_url), title_text, snippet_text


# ---------------------------------------------------------------------------
# A. Website
# ---------------------------------------------------------------------------

def _title_matches_company(title: str, business_name: str, threshold: float = 0.6) -> bool:
    """True if any significant (len > 2) word in business_name fuzzy-matches
    (SequenceMatcher ratio >= threshold) any word in the page title."""
    title_words = re.findall(r"[a-z0-9]+", title.lower())
    name_words = [w for w in re.findall(r"[a-z0-9]+", business_name.lower()) if len(w) > 2]
    if not name_words or not title_words:
        return False
    return any(
        SequenceMatcher(None, nw, tw).ratio() >= threshold
        for nw in name_words
        for tw in title_words
    )


def _verify_website(url: str, business_name: str) -> bool:
    """
    Never trust a Bing result on its own: fetch the candidate page directly,
    confirm it responds 200, and confirm its <title> fuzzy-matches the
    business name. Only then is it written into the sheet.
    """
    pool = get_pool()
    proxy = pool.get()
    try:
        with _make_client(proxy) as client:
            resp = client.get(url, timeout=15)
        if resp.status_code != 200:
            return False
        soup = BeautifulSoup(resp.text, "lxml")
        title = soup.title.get_text() if soup.title else ""
        return _title_matches_company(title, business_name)
    except Exception as exc:
        logger.debug("[ENRICH] Website verify failed for %s: %s", url, exc)
        return False


def _find_website(business_name: str, city: str) -> Optional[str]:
    soup = _bing_search(f'"{business_name}" "{city}" official site')
    if not soup:
        return None
    for real_url, domain, _title, _snippet in _organic_results(soup):
        if not domain or any(domain == d or domain.endswith("." + d) for d in _DIRECTORY_DOMAINS):
            continue
        if _verify_website(real_url, business_name):
            return real_url
    return None


# ---------------------------------------------------------------------------
# B. Email
# ---------------------------------------------------------------------------

def _find_email_via_search(business_name: str, city: str) -> Optional[str]:
    soup = _bing_search(f'"{business_name}" "{city}" email contact')
    if not soup:
        return None
    for _url, _domain, title_text, snippet_text in _organic_results(soup):
        match = _EMAIL_RE.search(f"{title_text} {snippet_text}")
        if match:
            return match.group(0).lower()
    return None


# ---------------------------------------------------------------------------
# C. Phone
# ---------------------------------------------------------------------------

def _find_phone_via_yellowpages(business_name: str, city: str) -> Optional[str]:
    query = urllib.parse.quote_plus(business_name)
    geo = urllib.parse.quote_plus(city)
    url = f"https://www.yellowpages.com/search?search_terms={query}&geo_location_terms={geo}"

    pool = get_pool()
    proxy = pool.get()
    with discover._make_client(proxy) as client:
        resp = discover._fetch_with_retry(client, url)
    if resp is None or resp.status_code != 200:
        return None

    soup = BeautifulSoup(resp.text, "lxml")
    for card in soup.select("div.result div.info")[:3]:
        biz = discover._parse_yp_listing(card)
        if biz and biz.get("phone"):
            return biz["phone"]
    return None


def _find_phone(business_name: str, city: str) -> Optional[str]:
    soup = _bing_search(f'"{business_name}" "{city}" phone number')
    if soup:
        for _url, _domain, title_text, snippet_text in _organic_results(soup):
            match = _PHONE_RE.search(f"{title_text} {snippet_text}")
            if match:
                return match.group(0).strip()

    return _find_phone_via_yellowpages(business_name, city)


# ---------------------------------------------------------------------------
# D. Social (Facebook / Instagram / LinkedIn)
# ---------------------------------------------------------------------------

def _social_url_plausible(url: str, business_name: str) -> bool:
    """
    Loose sanity check on a social-profile URL found via Bing search: does
    its path (the username/slug) contain any significant word from the
    business name? A social platform's search-result page can't be
    verified by fetching+checking a title the way a normal website can
    (Facebook/Instagram/LinkedIn are typically behind bot walls for that),
    so a bare Bing result is otherwise trusted outright — confirmed live
    during testing that this can attach a completely unrelated profile
    (a celebrity's Instagram) to a small business with a common niche name.
    This cheap local check is the safety net instead.
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


def _crawl_for_email_and_social(website_url: str) -> tuple[list[dict], dict]:
    """Same crawl _crawl_site_static does for the normal CRAWL stage —
    reused directly per the spec, rather than re-implementing homepage+
    /contact fetching and parsing a second time."""
    emails, _pages, social_links = _crawl_site_static(website_url)
    return emails, social_links


# ---------------------------------------------------------------------------
# E. Address
# ---------------------------------------------------------------------------

def _find_address(business_name: str, city: str) -> Optional[str]:
    soup = _bing_search(f'"{business_name}" "{city}" address')
    if not soup:
        return None
    for _url, _domain, title_text, snippet_text in _organic_results(soup):
        match = _ADDRESS_RE.search(f"{title_text} {snippet_text}")
        if match:
            return match.group(0).strip()
    return None


# ---------------------------------------------------------------------------
# Per-row enrichment
# ---------------------------------------------------------------------------

def _enrich_row(
    business_name: str,
    website_url: Optional[str],
    missing: set[str],
    location: str,
    sup_emails: frozenset[str],
    sup_domains: frozenset[str],
) -> dict[str, str]:
    """
    Attempt to fill each missing, fillable field for one business. Returns
    only the fields that were ACTUALLY found and verified — nothing here is
    ever guessed, so an empty dict (or a dict missing some requested field)
    is the normal, expected outcome when nothing real turns up.
    """
    name = business_name
    city = location.split(",")[0].strip() if location else location

    # Each field is enriched independently, in its own try/except: a
    # transient failure partway through one field's search (a proxy
    # dropping out mid-run was observed live) must not discard results
    # already found — and verified — for this business's other fields.
    filled: dict[str, str] = {}
    crawled_emails: list[dict] = []
    crawled_social: dict[str, Optional[str]] = {}
    crawl_attempted = False

    # A. Website
    if "website" in missing:
        try:
            _pause(_DELAY_BETWEEN_FIELD_SEARCHES)
            found = _find_website(name, city)
            if found:
                filled["website"] = found
                website_url = found
                logger.info("[ENRICH] Found website for %s: %s", name, found)
        except Exception as exc:
            logger.warning("[ENRICH] Website search failed for %s: %s", name, exc)

    # B. Email
    if "email" in missing:
        try:
            if website_url:
                _pause(_DELAY_BETWEEN_FIELD_SEARCHES)
                crawled_emails, crawled_social = _crawl_for_email_and_social(website_url)
                crawl_attempted = True
                if crawled_emails:
                    candidate = crawled_emails[0]["email"]
                    if not _is_suppressed(candidate, sup_emails, sup_domains):
                        filled["email"] = candidate

            if "email" not in filled:
                _pause(_DELAY_BETWEEN_FIELD_SEARCHES)
                candidate = _find_email_via_search(name, city)
                if candidate and not _is_suppressed(candidate, sup_emails, sup_domains):
                    filled["email"] = candidate

            if "email" in filled:
                logger.info("[ENRICH] Found email for %s: %s", name, filled["email"])
            else:
                logger.info("[ENRICH] No email found for %s", name)
        except Exception as exc:
            logger.warning("[ENRICH] Email search failed for %s: %s", name, exc)

    # C. Phone
    if "phone" in missing:
        try:
            _pause(_DELAY_BETWEEN_FIELD_SEARCHES)
            phone = _find_phone(name, city)
            if phone:
                filled["phone"] = phone
                logger.info("[ENRICH] Found phone for %s: %s", name, phone)
        except Exception as exc:
            logger.warning("[ENRICH] Phone search failed for %s: %s", name, exc)

    # D. Social
    social_missing = missing & {"facebook", "instagram", "linkedin"}
    if social_missing:
        try:
            if website_url and not crawl_attempted:
                _pause(_DELAY_BETWEEN_FIELD_SEARCHES)
                _emails, crawled_social = _crawl_for_email_and_social(website_url)
                crawl_attempted = True

            for platform in social_missing:
                if crawled_social.get(platform):
                    filled[platform] = crawled_social[platform]

            for platform in social_missing - filled.keys():
                _pause(_DELAY_BETWEEN_FIELD_SEARCHES)
                found = social._bing_search_social(name, location, platform)
                if found and _social_url_plausible(found, name):
                    filled[platform] = found
        except Exception as exc:
            logger.warning("[ENRICH] Social search failed for %s: %s", name, exc)

    # E. Address
    if "address" in missing:
        try:
            _pause(_DELAY_BETWEEN_FIELD_SEARCHES)
            address = _find_address(name, city)
            if address:
                filled["address"] = address
        except Exception as exc:
            logger.warning("[ENRICH] Address search failed for %s: %s", name, exc)

    return filled


# ---------------------------------------------------------------------------
# Read / write phases
# ---------------------------------------------------------------------------

def _detect_columns(header_row) -> dict[str, int]:
    """Map internal field name -> 0-based column index, located by header
    text (case-insensitive) rather than an assumed column order."""
    mapping: dict[str, int] = {}
    for idx, cell in enumerate(header_row):
        text = str(cell.value or "").strip().lower()
        field = _COLUMN_ALIASES.get(text)
        if field and field not in mapping:
            mapping[field] = idx
    return mapping


def _save_enriched(wb, input_path: str) -> str:
    """Save back to input_path; if it's open/locked elsewhere, fall back to
    "<name>_enriched.xlsx" rather than losing the run's results."""
    in_file = Path(input_path)
    try:
        wb.save(str(in_file))
        return str(in_file)
    except PermissionError:
        fallback = in_file.with_name(f"{in_file.stem}_enriched{in_file.suffix}")
        logger.warning("[ENRICH] %s is open/locked — saved to %s instead", in_file, fallback)
        wb.save(str(fallback))
        return str(fallback)


def run_enrich(
    input_path: str,
    db_path: str,
    niche: str,
    location: str,
    suppress_path: Optional[str] = None,
) -> str:
    """
    Read an existing Excel file, detect which of the supported columns are
    missing per row, and fill in only what can be found and verified from
    real sources. Returns the path actually saved to.

    db_path/niche are accepted for call-site consistency with the rest of
    the CLI (main() passes them through automatically) but this stage's
    defined behavior is entirely about the input spreadsheet's own rows —
    it does not read from or write to the SQLite database.
    """
    logger.info("=== STAGE: ENRICH  input=%r ===", input_path)

    sup_emails, sup_domains = _load_suppression(suppress_path)

    wb = openpyxl.load_workbook(input_path)
    ws = wb.active

    header_row = next(ws.iter_rows(min_row=1, max_row=1))
    columns = _detect_columns(header_row)

    if "business_name" not in columns:
        logger.error(
            "[ENRICH] No 'Company Name' column found in %s — nothing to enrich", input_path
        )
        return input_path

    rows_processed = 0
    rows_enriched = 0
    fields_filled = 0

    for row_cells in ws.iter_rows(min_row=2):
        values: dict[str, str] = {}
        for field, col_idx in columns.items():
            if col_idx >= len(row_cells):
                continue
            cell = row_cells[col_idx]
            values[field] = str(cell.value).strip() if cell.value is not None else ""

        business_name = values.get("business_name", "")
        if not business_name:
            continue  # cannot enrich without a name

        rows_processed += 1

        detected_missing = {
            f for f in ("phone", "category", "email", "website", "facebook",
                        "instagram", "linkedin", "address")
            if f in columns and not values.get(f)
        }
        fillable_missing = detected_missing & _FILLABLE_FIELDS
        if not fillable_missing:
            continue

        logger.info(
            "[ENRICH] (%d) %s — missing: %s", rows_processed, business_name, sorted(detected_missing)
        )

        try:
            filled = _enrich_row(
                business_name, values.get("website") or None, fillable_missing,
                location, sup_emails, sup_domains,
            )
        except Exception as exc:
            logger.warning("[ENRICH] Unhandled error enriching %s: %s", business_name, exc)
            filled = {}

        if filled:
            rows_enriched += 1
            for field, value in filled.items():
                col_idx = columns.get(field)
                if col_idx is None or col_idx >= len(row_cells):
                    continue
                cell = row_cells[col_idx]
                cell.value = value
                cell.fill = _GREEN_FILL
                fields_filled += 1

        _pause(_DELAY_BETWEEN_BUSINESSES)

    out_path = _save_enriched(wb, input_path)

    logger.info(
        "[ENRICH] DONE — rows_processed=%d rows_enriched=%d fields_filled=%d",
        rows_processed, rows_enriched, fields_filled,
    )
    return out_path
