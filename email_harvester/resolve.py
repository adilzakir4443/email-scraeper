"""
Stage 2 — RESOLVE.

For each business with a raw website_url:
  - Strip common tracking params (utm_*, fbclid, gclid, etc.)
  - Force https://
  - Follow exactly one redirect to get the canonical URL
  - Mark resolve_status as done/failed/no_site

Businesses without a website_url are marked no_site and left in the DB.
Already-resolved rows (resolve_status != 'pending') are skipped — idempotent.
"""

import logging
import re
import sqlite3
import urllib.parse
from typing import Optional

import httpx

from .db import get_conn
from .proxy import get_pool

logger = logging.getLogger(__name__)

# Tracking params to strip
_STRIP_PARAMS: frozenset[str] = frozenset(
    [
        "utm_source", "utm_medium", "utm_campaign", "utm_term", "utm_content",
        "fbclid", "gclid", "msclkid", "yclid", "ref", "referrer",
        "_ga", "_gac", "mc_cid", "mc_eid",
    ]
)

_SCHEME_RE = re.compile(r"^https?://", re.IGNORECASE)


def _strip_tracking(url: str) -> str:
    """Remove known tracking query parameters from url."""
    try:
        parsed = urllib.parse.urlparse(url)
        qs = urllib.parse.parse_qs(parsed.query, keep_blank_values=True)
        clean = {k: v for k, v in qs.items() if k.lower() not in _STRIP_PARAMS}
        new_query = urllib.parse.urlencode(clean, doseq=True)
        cleaned = parsed._replace(query=new_query)
        return urllib.parse.urlunparse(cleaned)
    except Exception:
        return url


def _force_https(url: str) -> str:
    """Upgrade http:// to https://; add https:// if scheme missing."""
    url = url.strip()
    if not _SCHEME_RE.match(url):
        url = "https://" + url
    return re.sub(r"^http://", "https://", url, flags=re.IGNORECASE)


def _resolve_one(url: str) -> Optional[str]:
    """
    Follow at most one redirect and return the final URL.
    Returns None on network/DNS failure.
    """
    pool = get_pool()
    proxy = pool.get()
    proxies = {"http://": proxy, "https://": proxy} if proxy else None

    for attempt in range(3):
        try:
            with httpx.Client(
                proxies=proxies,
                follow_redirects=True,
                max_redirects=5,
                timeout=httpx.Timeout(15.0),
                verify=True,
            ) as client:
                resp = client.head(url, headers={
                    "User-Agent": "Mozilla/5.0 (compatible; EmailHarvester/1.0)",
                })
                # HEAD sometimes returns 405; fall back to GET
                if resp.status_code == 405:
                    resp = client.get(url, headers={
                        "User-Agent": "Mozilla/5.0 (compatible; EmailHarvester/1.0)",
                    })
                final_url = str(resp.url)
                if proxy:
                    pool.report_success(proxy)
                return final_url
        except (httpx.ConnectError, httpx.TimeoutException, httpx.TooManyRedirects) as exc:
            logger.debug("Resolve attempt %d failed for %s: %s", attempt, url, exc)
            if proxy:
                pool.report_failure(proxy)
            pool.backoff_sleep(attempt + 1)
        except Exception as exc:
            logger.warning("Unexpected error resolving %s: %s", url, exc)
            break

    return None


def run_resolve(db_path: str) -> None:
    """
    Stage 2: normalise all pending website URLs.
    Already-resolved rows are skipped automatically.
    """
    logger.info("=== STAGE 2: RESOLVE ===")

    with get_conn(db_path) as conn:
        # Mark businesses with no website as no_site upfront
        conn.execute(
            "UPDATE businesses SET resolve_status='no_site' WHERE website_url IS NULL AND resolve_status='pending'"
        )

    with get_conn(db_path) as conn:
        rows = conn.execute(
            "SELECT id, website_url FROM businesses WHERE resolve_status='pending'"
        ).fetchall()

    total = len(rows)
    logger.info("[RESOLVE] %d URLs to process", total)

    done = failed = 0

    for row in rows:
        biz_id: int = row["id"]
        raw_url: str = row["website_url"]

        cleaned = _strip_tracking(_force_https(raw_url))
        final = _resolve_one(cleaned)

        if final:
            with get_conn(db_path) as conn:
                conn.execute(
                    "UPDATE businesses SET normalized_url=?, resolve_status='done' WHERE id=?",
                    (final, biz_id),
                )
            done += 1
        else:
            with get_conn(db_path) as conn:
                conn.execute(
                    "UPDATE businesses SET normalized_url=?, resolve_status='failed' WHERE id=?",
                    (cleaned, biz_id),  # keep cleaned URL even if resolve failed
                )
            failed += 1
            logger.debug("[RESOLVE] Failed to resolve biz_id=%d url=%s", biz_id, raw_url)

    logger.info("[RESOLVE] DONE — resolved=%d  failed=%d  total=%d", done, failed, total)
