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


def init_db(db_path: str) -> None:
    """Create tables if they don't exist. Safe to call on every run."""
    path = Path(db_path)
    path.parent.mkdir(parents=True, exist_ok=True)
    with sqlite3.connect(db_path) as conn:
        conn.executescript(SCHEMA)
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
