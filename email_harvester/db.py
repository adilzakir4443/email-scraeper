"""SQLite schema and database helpers. All state lives here — stages check this to resume."""

import sqlite3
import logging
from contextlib import contextmanager
from pathlib import Path
from typing import Generator

logger = logging.getLogger(__name__)

SCHEMA = """
PRAGMA journal_mode=WAL;
PRAGMA foreign_keys=ON;

CREATE TABLE IF NOT EXISTS businesses (
    id               INTEGER PRIMARY KEY AUTOINCREMENT,
    niche            TEXT    NOT NULL,
    location         TEXT    NOT NULL,
    business_name    TEXT    NOT NULL,
    website_url      TEXT,                   -- raw URL from directory
    normalized_url   TEXT,                   -- cleaned URL after RESOLVE stage
    phone            TEXT,
    address          TEXT,
    category         TEXT,
    source           TEXT    NOT NULL,       -- yellowpages | bing | yelp
    discovered_at    TEXT    DEFAULT (datetime('now')),
    resolve_status   TEXT    DEFAULT 'pending',  -- pending | done | no_site | failed
    crawl_status     TEXT    DEFAULT 'pending',  -- pending | done | failed | no_emails_static
    playwright_tried INTEGER DEFAULT 0,
    facebook_url     TEXT,
    instagram_url    TEXT,
    linkedin_url     TEXT,
    social_status    TEXT    DEFAULT 'pending',  -- pending | done
    UNIQUE(niche, location, business_name, source)
);

CREATE TABLE IF NOT EXISTS emails (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    business_id     INTEGER NOT NULL REFERENCES businesses(id) ON DELETE CASCADE,
    email           TEXT    NOT NULL,
    source_url      TEXT,
    extract_method  TEXT,   -- mailto | regex | obfuscated | cfemail
    syntax_valid    INTEGER,
    is_disposable   INTEGER,
    is_role         INTEGER,
    is_freemail     INTEGER,
    mx_status       TEXT,   -- acceptable | risky | invalid | error
    tier            TEXT,   -- acceptable | risky | invalid
    verified_at     TEXT,
    UNIQUE(business_id, email)
);

CREATE INDEX IF NOT EXISTS idx_businesses_crawl   ON businesses(crawl_status);
CREATE INDEX IF NOT EXISTS idx_businesses_resolve ON businesses(resolve_status);
CREATE INDEX IF NOT EXISTS idx_emails_business    ON emails(business_id);
CREATE INDEX IF NOT EXISTS idx_emails_tier        ON emails(tier);
"""


# Columns added to `businesses` after it was first created. CREATE TABLE IF
# NOT EXISTS does not retrofit new columns onto an existing table, so a DB
# from before a given column was added needs an explicit ALTER TABLE — this
# keeps init_db safe to run against old databases without breaking them.
_BUSINESS_COLUMN_MIGRATIONS: dict[str, str] = {
    "facebook_url":  "ALTER TABLE businesses ADD COLUMN facebook_url TEXT",
    "instagram_url": "ALTER TABLE businesses ADD COLUMN instagram_url TEXT",
    "linkedin_url":  "ALTER TABLE businesses ADD COLUMN linkedin_url TEXT",
    "social_status": "ALTER TABLE businesses ADD COLUMN social_status TEXT DEFAULT 'pending'",
}


def _migrate_business_columns(conn: sqlite3.Connection) -> None:
    existing = {row[1] for row in conn.execute("PRAGMA table_info(businesses)")}
    for column, ddl in _BUSINESS_COLUMN_MIGRATIONS.items():
        if column not in existing:
            conn.execute(ddl)
            logger.info("Migrated businesses table: added column %s", column)


def init_db(db_path: str) -> None:
    """Create tables if they don't exist. Safe to call on every run."""
    path = Path(db_path)
    path.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(db_path)
    try:
        conn.executescript(SCHEMA)
        _migrate_business_columns(conn)
        conn.commit()
    finally:
        # sqlite3.Connection's context-manager protocol only commits/rolls
        # back on exit — it does not close the connection. Without an
        # explicit close(), this connection (and its WAL file lock) leaks
        # for the life of the process.
        conn.close()
    logger.info("Database initialised at %s", db_path)


@contextmanager
def get_conn(db_path: str) -> Generator[sqlite3.Connection, None, None]:
    """Context manager yielding an open connection with row_factory set."""
    conn = sqlite3.connect(db_path, timeout=30)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys=ON")
    try:
        yield conn
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


# ---------------------------------------------------------------------------
# Upsert helpers
# ---------------------------------------------------------------------------

def upsert_business(
    conn: sqlite3.Connection,
    *,
    niche: str,
    location: str,
    business_name: str,
    website_url: str | None,
    phone: str | None,
    address: str | None,
    category: str | None,
    source: str,
) -> int:
    """Insert or ignore a business row; return its id."""
    conn.execute(
        """
        INSERT OR IGNORE INTO businesses
            (niche, location, business_name, website_url, phone, address, category, source)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
        """,
        (niche, location, business_name, website_url, phone, address, category, source),
    )
    row = conn.execute(
        "SELECT id FROM businesses WHERE niche=? AND location=? AND business_name=? AND source=?",
        (niche, location, business_name, source),
    ).fetchone()
    return row["id"]


def upsert_email(
    conn: sqlite3.Connection,
    *,
    business_id: int,
    email: str,
    source_url: str | None,
    extract_method: str,
) -> None:
    """Insert email if not already present for this business."""
    conn.execute(
        """
        INSERT OR IGNORE INTO emails (business_id, email, source_url, extract_method)
        VALUES (?, ?, ?, ?)
        """,
        (business_id, email, source_url, extract_method),
    )


def upsert_social(
    conn: sqlite3.Connection,
    *,
    business_id: int,
    facebook_url: str | None,
    instagram_url: str | None,
    linkedin_url: str | None,
) -> None:
    """
    Set social links on a business, only filling in fields that are still
    NULL (never overwrites an already-found link), then marks social_status
    as 'done' regardless of whether anything new was found — a business
    that was searched and came up empty should not be retried forever.
    """
    conn.execute(
        """
        UPDATE businesses
        SET facebook_url  = COALESCE(facebook_url, ?),
            instagram_url = COALESCE(instagram_url, ?),
            linkedin_url  = COALESCE(linkedin_url, ?),
            social_status = 'done'
        WHERE id = ?
        """,
        (facebook_url, instagram_url, linkedin_url, business_id),
    )


# ---------------------------------------------------------------------------
# Stage-resume queries
# ---------------------------------------------------------------------------

def count_businesses(conn: sqlite3.Connection, niche: str, location: str) -> dict[str, int]:
    """Return funnel counts for a given niche+location run."""
    total = conn.execute(
        "SELECT COUNT(*) FROM businesses WHERE niche=? AND location=?",
        (niche, location),
    ).fetchone()[0]
    with_site = conn.execute(
        "SELECT COUNT(*) FROM businesses WHERE niche=? AND location=? AND website_url IS NOT NULL",
        (niche, location),
    ).fetchone()[0]
    crawled = conn.execute(
        "SELECT COUNT(*) FROM businesses WHERE niche=? AND location=? AND crawl_status='done'",
        (niche, location),
    ).fetchone()[0]
    return {"total": total, "with_site": with_site, "crawled": crawled}
