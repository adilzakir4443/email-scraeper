"""
Stage 4b — WRITE.

Reads verified emails from SQLite, applies suppression list,
deduplicates (by email then by domain), and writes an Excel file with two
sheets:
  "Leads"      — one row per verified email.
  "No Website" — businesses with no site at all (nothing for CRAWL/VERIFY
                 to work with), so their only follow-up contact is phone /
                 social media, which SOCIAL's Bing-search fallback fills in.

Sheet 1 columns (exact order):
  Company Name | Owner Name | Phone | Category | Email | Website |
  Facebook | Instagram | LinkedIn | Address | Comment

Sheet 2 columns (exact order):
  Company Name | Phone | Category | Address | Facebook | Instagram | LinkedIn

Comment format:
  mx_status=acceptable; role=false; source=yellowpages
"""

import csv
import logging
import re
import sqlite3
from datetime import datetime, timezone
from pathlib import Path
from typing import Optional

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

from .db import get_conn

logger = logging.getLogger(__name__)

COLUMNS = [
    "Company Name",
    "Owner Name",
    "Phone",
    "Category",
    "Email",
    "Website",
    "Facebook",
    "Instagram",
    "LinkedIn",
    "Address",
    "Comment",
]

NO_WEBSITE_COLUMNS = [
    "Company Name",
    "Phone",
    "Category",
    "Address",
    "Facebook",
    "Instagram",
    "LinkedIn",
]

# Tiers to include in output (invalid is dropped entirely)
OUTPUT_TIERS = frozenset(["acceptable", "risky"])


def _load_suppression(suppress_path: Optional[str]) -> tuple[frozenset[str], frozenset[str]]:
    """
    Load suppressed emails and domains from a CSV file.
    File must have columns: type, value  (type = email | domain)
    Returns (suppressed_emails, suppressed_domains).
    """
    if not suppress_path:
        return frozenset(), frozenset()

    path = Path(suppress_path)
    if not path.exists():
        logger.warning("Suppression file not found: %s", suppress_path)
        return frozenset(), frozenset()

    emails: set[str] = set()
    domains: set[str] = set()

    try:
        with open(path, newline="", encoding="utf-8") as f:
            reader = csv.DictReader(f)
            for row in reader:
                kind = (row.get("type") or "").strip().lower()
                value = (row.get("value") or "").strip().lower()
                if not value:
                    continue
                if kind == "email":
                    emails.add(value)
                elif kind == "domain":
                    domains.add(value)
                else:
                    # Guess: if value contains @ it's an email
                    if "@" in value:
                        emails.add(value)
                    else:
                        domains.add(value)
    except Exception as exc:
        logger.error("Error reading suppression file %s: %s", suppress_path, exc)

    logger.info(
        "[WRITE] Suppression: %d emails, %d domains loaded", len(emails), len(domains)
    )
    return frozenset(emails), frozenset(domains)


def _is_suppressed(
    email: str,
    sup_emails: frozenset[str],
    sup_domains: frozenset[str],
) -> bool:
    email_lower = email.lower()
    if email_lower in sup_emails:
        return True
    try:
        domain = email_lower.split("@", 1)[1]
    except IndexError:
        return False
    return domain in sup_domains


def _build_comment(row: sqlite3.Row) -> str:
    """Build the Comment cell value from verify fields and source."""
    parts = []

    mx = row["mx_status"] or "mailbox_unverified"
    # Never say "verified" — use mx_status / mailbox_unverified
    parts.append(f"mx_status={mx}")

    parts.append(f"role={'true' if row['is_role'] else 'false'}")
    parts.append(f"freemail={'true' if row['is_freemail'] else 'false'}")
    parts.append(f"source={row['source']}")
    parts.append(f"extract={row['extract_method'] or 'unknown'}")
    parts.append(f"tier={row['tier']}")

    return "; ".join(parts)


def _domain_of(email: str) -> str:
    try:
        return email.lower().split("@", 1)[1]
    except IndexError:
        return ""


def _deduplicate(rows: list[sqlite3.Row]) -> list[sqlite3.Row]:
    """
    Deduplicate: first by exact email, then keep only one email per domain
    (prefer acceptable over risky; if same tier prefer shorter email).
    """
    # Step 1: unique by email (rows already have unique emails from DB UNIQUE constraint,
    # but we might have same email from multiple businesses — keep first)
    seen_emails: set[str] = set()
    unique_by_email: list[sqlite3.Row] = []
    for row in rows:
        key = row["email"].lower()
        if key not in seen_emails:
            seen_emails.add(key)
            unique_by_email.append(row)

    # Step 2: one email per domain — keep best-tier entry
    tier_rank = {"acceptable": 0, "risky": 1, "invalid": 2}
    best_by_domain: dict[str, sqlite3.Row] = {}
    for row in unique_by_email:
        domain = _domain_of(row["email"])
        existing = best_by_domain.get(domain)
        if existing is None:
            best_by_domain[domain] = row
        else:
            if tier_rank.get(row["tier"], 9) < tier_rank.get(existing["tier"], 9):
                best_by_domain[domain] = row

    return list(best_by_domain.values())


def _style_header(ws) -> None:
    """Apply bold + background to header row."""
    header_fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF", size=11)
    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=False)


def _autofit_columns(ws, columns: list[str]) -> None:
    """Approximate auto-fit: size each column to its longest cell value."""
    for col_idx, col_name in enumerate(columns, 1):
        col_letter = get_column_letter(col_idx)
        max_len = len(col_name)
        for cell in ws[col_letter]:
            try:
                if cell.value:
                    max_len = max(max_len, len(str(cell.value)))
            except Exception:
                pass
        ws.column_dimensions[col_letter].width = min(max_len + 4, 60)


def _load_or_create_workbook(out_file: Path) -> "openpyxl.Workbook":
    """
    Open *out_file* to update it in place if it already exists and is a
    valid workbook; otherwise start a fresh one. Every run used to build a
    brand new Workbook() and overwrite out_path unconditionally, so
    re-running the tool against the same output file (a common workflow:
    build one master leads list, refreshed across runs) silently discarded
    everything written by earlier runs.
    """
    if out_file.exists():
        try:
            wb = openpyxl.load_workbook(str(out_file))
            logger.info("[WRITE] Updating existing workbook %s", out_file)
            return wb
        except Exception as exc:
            logger.warning(
                "[WRITE] Could not open existing %s as a workbook (%s) — "
                "starting a new one instead", out_file, exc,
            )
    wb = openpyxl.Workbook()
    wb.remove(wb.active)  # drop the auto-created default sheet; we name our own
    return wb


def _get_or_create_sheet(wb, title: str, columns: list[str]):
    """Return (worksheet, is_new). A new sheet gets its header row + styling
    (an existing one already has both, so _style_header is skipped for it —
    re-styling on every run would be wasted work and risks fighting any
    manual tweaks the header row picked up)."""
    if title in wb.sheetnames:
        return wb[title], False
    ws = wb.create_sheet(title)
    ws.append(columns)
    _style_header(ws)
    return ws, True


def _existing_row_keys(ws, col_indices: list[int]) -> set[tuple]:
    """
    Read every existing data row (skipping the header) and return the set of
    dedup keys built from the given 0-based column indices, so writing
    doesn't re-add a row that's already in the sheet.
    """
    keys: set[tuple] = set()
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row is None:
            continue
        key = tuple(str(row[i] or "").strip().lower() for i in col_indices)
        keys.add(key)
    return keys


def _clear_matching_rows(ws, col_indices: list[int], keys_to_clear: set[tuple]) -> int:
    """
    Delete every data row (skipping the header) whose values at
    *col_indices* form a key in *keys_to_clear*. Used to refresh a single
    niche+location's rows in place on a re-run — without this, either
    stale rows from an old run would never be corrected (a business that
    dropped a tier, changed phone, etc. would keep showing outdated data
    forever), or a blunt "clear everything" would also wipe out rows
    belonging to a *different* niche/location sharing the same output file.
    Rows are deleted bottom-to-top so earlier deletions don't shift the
    row indices of ones still to be checked. Returns the count deleted.
    """
    rows_to_delete = [
        row_idx
        for row_idx in range(2, ws.max_row + 1)
        if tuple(
            str(ws.cell(row=row_idx, column=col_idx + 1).value or "").strip().lower()
            for col_idx in col_indices
        ) in keys_to_clear
    ]
    for row_idx in reversed(rows_to_delete):
        ws.delete_rows(row_idx, 1)
    return len(rows_to_delete)


def _is_file_locked(path: Path) -> bool:
    """True if *path* exists and is currently open/locked elsewhere (e.g.
    open in Excel, which holds an exclusive lock on Windows)."""
    if not path.exists():
        return False
    try:
        with open(path, "r+b"):
            pass
        return False
    except PermissionError:
        return True


def _next_available_path(path: Path) -> Path:
    """Return the first "<stem>_N<suffix>" path (N = 1, 2, 3, ...) that
    isn't locked, starting from *path* itself."""
    if not _is_file_locked(path):
        return path
    n = 1
    while True:
        candidate = path.with_name(f"{path.stem}_{n}{path.suffix}")
        if not _is_file_locked(candidate):
            return candidate
        n += 1


def _resolve_output_path(path: Path) -> Path:
    """
    If *path* doesn't exist or isn't locked, return it unchanged. If it's
    open/locked elsewhere, return the next available "<stem>_N<suffix>"
    path instead (checked proactively via _is_file_locked, since a plain
    'r' open — what load_workbook itself would do — doesn't reliably fail
    just because Excel has the file open for writing).
    """
    resolved = _next_available_path(path)
    if resolved != path:
        logger.warning("[WRITE] %s is open/locked — will save to %s instead", path, resolved)
    return resolved


def run_write(
    db_path: str,
    niche: str,
    location: str,
    out_path: Optional[str],
    suppress_path: Optional[str],
) -> str:
    """
    Stage 4b: query DB, apply suppression, dedupe, write Excel.
    Returns the output file path.
    """
    logger.info("=== STAGE 4b: WRITE ===")

    sup_emails, sup_domains = _load_suppression(suppress_path)

    # Build output filename if not provided
    if not out_path:
        date_str = datetime.now(timezone.utc).strftime("%Y%m%d")
        safe_niche = re.sub(r"[^\w]+", "_", niche).strip("_").lower()
        safe_loc = re.sub(r"[^\w]+", "_", location).strip("_").lower()
        out_path = f"leads_{safe_niche}_{safe_loc}_{date_str}.xlsx"

    with get_conn(db_path) as conn:
        rows = conn.execute(
            """
            SELECT
                b.business_name,
                b.phone,
                b.category,
                b.address,
                b.normalized_url AS website,
                b.website_url    AS website_raw,
                b.facebook_url,
                b.instagram_url,
                b.linkedin_url,
                b.source,
                e.email,
                e.extract_method,
                e.is_role,
                e.is_freemail,
                e.mx_status,
                e.tier
            FROM emails e
            JOIN businesses b ON b.id = e.business_id
            WHERE b.niche = ? AND b.location = ?
              AND e.tier IN ('acceptable', 'risky')
            ORDER BY e.tier ASC, b.business_name ASC
            """,
            (niche, location),
        ).fetchall()

    logger.info("[WRITE] %d raw rows from DB (before suppression + dedupe)", len(rows))

    # Apply suppression
    unsuppressed = [
        r for r in rows
        if not _is_suppressed(r["email"], sup_emails, sup_domains)
    ]
    suppressed_count = len(rows) - len(unsuppressed)
    logger.info("[WRITE] Suppressed %d rows", suppressed_count)

    # Deduplicate
    deduped = _deduplicate(unsuppressed)
    logger.info("[WRITE] After dedup: %d rows", len(deduped))

    # Count tiers for funnel log
    tier_counts: dict[str, int] = {}
    for r in deduped:
        tier_counts[r["tier"]] = tier_counts.get(r["tier"], 0) + 1
    logger.info(
        "[WRITE] Tiers — acceptable=%d  risky=%d",
        tier_counts.get("acceptable", 0),
        tier_counts.get("risky", 0),
    )

    # Build (or re-open) the Excel workbook. If out_path is already open in
    # another program (e.g. Excel), write to "<stem>_1<suffix>" instead
    # (then _2, _3, ... if those are locked too) rather than failing.
    out_file = _resolve_output_path(Path(out_path))
    out_file.parent.mkdir(parents=True, exist_ok=True)
    wb = _load_or_create_workbook(out_file)
    ws, is_new_leads_sheet = _get_or_create_sheet(wb, "Leads", COLUMNS)

    # Refresh THIS niche+location's rows in place: clear whatever rows
    # already belong to it (looked up by email, via the businesses table —
    # not just "is this email in today's result set", so a row stays
    # correctly identified as this niche/location's even if e.g. its tier
    # changed and it no longer qualifies this run) before rewriting them
    # below. Rows belonging to any OTHER niche/location already in this
    # same output file are left completely untouched.
    if not is_new_leads_sheet:
        with get_conn(db_path) as conn:
            owned_emails = {
                r["email"].strip().lower()
                for r in conn.execute(
                    """SELECT e.email FROM emails e
                       JOIN businesses b ON b.id = e.business_id
                       WHERE b.niche = ? AND b.location = ?""",
                    (niche, location),
                ).fetchall()
            }
        if owned_emails:
            cleared = _clear_matching_rows(ws, [4], {(e,) for e in owned_emails})
            if cleared:
                logger.info(
                    "[WRITE] Leads sheet — cleared %d existing row(s) for "
                    "niche=%r location=%r before rewriting",
                    cleared, niche, location,
                )

    # Safety net against cross-niche email collisions — after the clear
    # above, this is empty for the common case (this niche/location no
    # longer has any rows left to collide with).
    existing_emails = set() if is_new_leads_sheet else _existing_row_keys(ws, [4])

    risky_fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")

    appended = 0
    skipped_existing = 0
    for row in deduped:
        email_key = (row["email"].strip().lower(),)
        if email_key in existing_emails:
            skipped_existing += 1
            continue

        website = row["website"] or row["website_raw"] or ""
        comment = _build_comment(row)

        ws.append(
            [
                row["business_name"],
                "",                      # Owner Name — intentionally blank
                row["phone"] or "",
                row["category"] or "",
                row["email"],
                website,
                row["facebook_url"] or "",
                row["instagram_url"] or "",
                row["linkedin_url"] or "",
                row["address"] or "",
                comment,
            ]
        )

        # Highlight risky rows in light yellow
        if row["tier"] == "risky":
            for cell in ws[ws.max_row]:
                cell.fill = risky_fill

        existing_emails.add(email_key)
        appended += 1

    logger.info(
        "[WRITE] Leads sheet — appended %d new row(s), skipped %d already present",
        appended, skipped_existing,
    )

    _autofit_columns(ws, COLUMNS)
    ws.freeze_panes = "A2"

    # Sheet 2 — businesses with no website at all. CRAWL/VERIFY had nothing
    # to work with for these, so they have no email; SOCIAL's Bing-search
    # fallback (_find_social_no_website) is their only source of contact
    # info besides the phone number already on file.
    with get_conn(db_path) as conn:
        no_website_rows = conn.execute(
            """
            SELECT business_name, phone, category, address,
                   facebook_url, instagram_url, linkedin_url
            FROM businesses
            WHERE website_url IS NULL AND normalized_url IS NULL
              AND niche = ? AND location = ?
            ORDER BY business_name ASC
            """,
            (niche, location),
        ).fetchall()

        # Debug aid: if this sheet is unexpectedly empty, this line makes it
        # obvious whether the SQL genuinely found nothing for this niche/
        # location (e.g. every discovered business really does have a site)
        # versus a bug elsewhere (e.g. a scraper stashing a non-website URL
        # in website_url, as the Yelp scraper used to do with its own
        # listing-page link — see discover.py's _parse_yelp_listing).
        total_for_query = conn.execute(
            "SELECT COUNT(*) FROM businesses WHERE niche = ? AND location = ?",
            (niche, location),
        ).fetchone()[0]
        with_site = conn.execute(
            """SELECT COUNT(*) FROM businesses
               WHERE niche = ? AND location = ?
                 AND (website_url IS NOT NULL OR normalized_url IS NOT NULL)""",
            (niche, location),
        ).fetchone()[0]
        logger.info(
            "[WRITE] No Website query — %d raw row(s) before dedup "
            "(of %d total businesses for this niche/location; %d have a site)",
            len(no_website_rows), total_for_query, with_site,
        )

    ws2, is_new_no_website_sheet = _get_or_create_sheet(wb, "No Website", NO_WEBSITE_COLUMNS)

    # Same refresh-in-place approach as the Leads sheet above: clear rows
    # already belonging to this niche+location (identified by every
    # business currently on record for it, not just today's no-website
    # ones — so a business that picked up a website since the last run
    # correctly drops off this sheet instead of lingering as a stale row)
    # before rewriting; other niches/locations in the file are untouched.
    if not is_new_no_website_sheet:
        with get_conn(db_path) as conn:
            owned_no_website = {
                (str(r["business_name"] or "").strip().lower(), str(r["phone"] or "").strip().lower())
                for r in conn.execute(
                    "SELECT business_name, phone FROM businesses WHERE niche = ? AND location = ?",
                    (niche, location),
                ).fetchall()
            }
        if owned_no_website:
            cleared2 = _clear_matching_rows(ws2, [0, 1], owned_no_website)
            if cleared2:
                logger.info(
                    "[WRITE] No Website sheet — cleared %d existing row(s) for "
                    "niche=%r location=%r before rewriting",
                    cleared2, niche, location,
                )

    # (Company Name, Phone) identifies a business here — there's no email to
    # key on for this sheet.
    existing_no_website = set() if is_new_no_website_sheet else _existing_row_keys(ws2, [0, 1])

    no_website_appended = 0
    no_website_skipped = 0
    for row in no_website_rows:
        key = (str(row["business_name"] or "").strip().lower(), str(row["phone"] or "").strip().lower())
        if key in existing_no_website:
            no_website_skipped += 1
            continue

        ws2.append(
            [
                row["business_name"],
                row["phone"] or "",
                row["category"] or "",
                row["address"] or "",
                row["facebook_url"] or "",
                row["instagram_url"] or "",
                row["linkedin_url"] or "",
            ]
        )
        existing_no_website.add(key)
        no_website_appended += 1

    _autofit_columns(ws2, NO_WEBSITE_COLUMNS)
    ws2.freeze_panes = "A2"

    logger.info(
        "[WRITE] No Website sheet — appended %d new row(s), skipped %d already present",
        no_website_appended, no_website_skipped,
    )

    try:
        wb.save(str(out_file))
    except PermissionError:
        # Race-condition safety net: the file passed the proactive
        # _resolve_output_path check earlier but got opened in Excel (or
        # similar) in the time since. All the scraping/verification work
        # for this run is done and expensive to redo — don't throw it away
        # over a locked file. Retry against "_1", "_2", ... using the save
        # attempt itself as the lock check (more reliable here than the
        # r+b pre-check, since that's exactly what just disagreed with it).
        n = 1
        while True:
            fallback = out_file.with_name(f"{out_file.stem}_{n}{out_file.suffix}")
            try:
                wb.save(str(fallback))
                break
            except PermissionError:
                n += 1
        logger.warning(
            "[WRITE] %s is open/locked — saved to %s instead", out_file, fallback,
        )
        out_file = fallback

    logger.info(
        "[WRITE] DONE — %d new lead(s) in Leads (%d total now) to %s  (suppressed=%d)",
        appended, ws.max_row - 1, out_file, suppressed_count,
    )
    return str(out_file)
