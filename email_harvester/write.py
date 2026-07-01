"""
Stage 4b — WRITE.

Reads verified emails from SQLite, applies suppression list,
deduplicates (by email then by domain), and writes an Excel file.

Output columns (exact order):
  Company Name | Owner Name | Phone | Category | Email | Website | Address | Comment

Comment format:
  mx_status=acceptable; role=false; source=yellowpages
"""

import csv
import logging
import re
import sqlite3
import urllib.parse
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
    "Address",
    "Comment",
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

    # Build Excel
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Leads"

    ws.append(COLUMNS)
    _style_header(ws)

    risky_fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")

    for row in deduped:
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
                row["address"] or "",
                comment,
            ]
        )

        # Highlight risky rows in light yellow
        if row["tier"] == "risky":
            for cell in ws[ws.max_row]:
                cell.fill = risky_fill

    # Auto-fit column widths (approximate)
    for col_idx, col_name in enumerate(COLUMNS, 1):
        col_letter = get_column_letter(col_idx)
        max_len = len(col_name)
        for cell in ws[col_letter]:
            try:
                if cell.value:
                    max_len = max(max_len, len(str(cell.value)))
            except Exception:
                pass
        ws.column_dimensions[col_letter].width = min(max_len + 4, 60)

    ws.freeze_panes = "A2"

    out_file = Path(out_path)
    out_file.parent.mkdir(parents=True, exist_ok=True)
    wb.save(str(out_file))

    logger.info(
        "[WRITE] DONE — wrote %d leads to %s  (suppressed=%d)",
        len(deduped), out_file, suppressed_count,
    )
    return str(out_file)
