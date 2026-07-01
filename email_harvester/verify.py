"""
Stage 4a — VERIFY.

For each unverified email in the DB:
  1. Syntax check via email-validator
  2. Disposable-domain blocklist check
  3. Role-account flagging (info@, sales@, etc.)
  4. Free-mail domain detection (gmail, yahoo, outlook, etc.)
  5. MX record lookup via dnspython

Assigns each email a tier:
  - invalid  → drop from output (bad syntax, no MX, or disposable)
  - risky    → include with warning (role account or free-mail provider)
  - acceptable → passes all checks

Never performs SMTP probing or catch-all detection.
Resumable: already-verified rows (verified_at IS NOT NULL) are skipped.
"""

import logging
import re
import sqlite3
from datetime import datetime, timezone
from typing import Optional

import dns.resolver
import dns.exception
from email_validator import validate_email, EmailNotValidError

from .db import get_conn

logger = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Role-account prefixes (FLAG, do not drop)
# ---------------------------------------------------------------------------
ROLE_PREFIXES: frozenset[str] = frozenset(
    [
        "info", "sales", "admin", "support", "contact", "hello",
        "help", "noreply", "no-reply", "postmaster", "webmaster",
        "billing", "accounts", "accounting", "enquiries", "enquiry",
        "office", "general", "mail", "email",
    ]
)

# ---------------------------------------------------------------------------
# Free-mail providers (risky tier, not dropped)
# ---------------------------------------------------------------------------
FREEMAIL_DOMAINS: frozenset[str] = frozenset(
    [
        "gmail.com", "yahoo.com", "yahoo.co.uk", "yahoo.com.au",
        "outlook.com", "hotmail.com", "hotmail.co.uk", "live.com",
        "icloud.com", "me.com", "mac.com",
        "aol.com", "protonmail.com", "proton.me",
        "zoho.com", "yandex.com", "mail.com", "gmx.com",
    ]
)

# DNS resolver with a 5-second timeout
_RESOLVER = dns.resolver.Resolver()
_RESOLVER.lifetime = 5.0
_RESOLVER.timeout = 5.0

# In-process MX cache to avoid redundant lookups within a run
_mx_cache: dict[str, str] = {}


def _check_syntax(email: str) -> bool:
    """Return True if email passes strict RFC syntax validation."""
    try:
        validate_email(email, check_deliverability=False)
        return True
    except EmailNotValidError:
        return False


def _is_disposable(domain: str) -> bool:
    """Return True if domain appears in the disposable-email-domains list."""
    try:
        import disposable_email_domains  # type: ignore
        return domain in disposable_email_domains.blocklist
    except ImportError:
        logger.debug("disposable-email-domains not installed; skipping disposable check")
        return False


def _is_role(local: str) -> bool:
    """Return True if the local part is a known role-account prefix."""
    return local.lower() in ROLE_PREFIXES


def _is_freemail(domain: str) -> bool:
    return domain.lower() in FREEMAIL_DOMAINS


def _mx_lookup(domain: str) -> str:
    """
    Check MX records. Returns:
      'acceptable' — valid MX found
      'invalid'    — no MX records (DNS NXDOMAIN, NoAnswer)
      'error'      — transient DNS error (don't reject permanently)
    """
    if domain in _mx_cache:
        return _mx_cache[domain]

    try:
        answers = _RESOLVER.resolve(domain, "MX")
        status = "acceptable" if answers else "invalid"
    except (dns.resolver.NXDOMAIN, dns.resolver.NoAnswer, dns.resolver.NoNameservers):
        status = "invalid"
    except dns.exception.Timeout:
        logger.debug("DNS timeout for %s", domain)
        status = "error"
    except Exception as exc:
        logger.debug("MX lookup error for %s: %s", domain, exc)
        status = "error"

    _mx_cache[domain] = status
    return status


def _compute_tier(
    *,
    syntax_valid: bool,
    is_disposable: bool,
    is_role: bool,
    is_freemail: bool,
    mx_status: str,
) -> str:
    """Compute final tier from individual check results."""
    if not syntax_valid:
        return "invalid"
    if is_disposable:
        return "invalid"
    if mx_status == "invalid":
        return "invalid"
    if is_role or is_freemail:
        return "risky"
    return "acceptable"


def run_verify(db_path: str) -> None:
    """
    Stage 4a: verify all emails not yet verified.
    Updates syntax_valid, is_disposable, is_role, is_freemail, mx_status, tier, verified_at.
    """
    logger.info("=== STAGE 4a: VERIFY ===")

    with get_conn(db_path) as conn:
        rows = conn.execute(
            "SELECT id, email FROM emails WHERE verified_at IS NULL"
        ).fetchall()

    total = len(rows)
    logger.info("[VERIFY] %d emails to verify", total)

    counts: dict[str, int] = {"acceptable": 0, "risky": 0, "invalid": 0, "error": 0}

    for row in rows:
        email_id: int = row["id"]
        email: str = row["email"]

        try:
            local, domain = email.rsplit("@", 1)
        except ValueError:
            # Malformed — mark invalid immediately
            with get_conn(db_path) as conn:
                conn.execute(
                    """UPDATE emails SET syntax_valid=0, tier='invalid',
                       verified_at=? WHERE id=?""",
                    (datetime.now(timezone.utc).isoformat(), email_id),
                )
            counts["invalid"] += 1
            continue

        syntax_ok = _check_syntax(email)
        disposable = _is_disposable(domain) if syntax_ok else False
        role = _is_role(local)
        freemail = _is_freemail(domain)
        mx = _mx_lookup(domain) if syntax_ok and not disposable else "invalid"

        tier = _compute_tier(
            syntax_valid=syntax_ok,
            is_disposable=disposable,
            is_role=role,
            is_freemail=freemail,
            mx_status=mx,
        )

        with get_conn(db_path) as conn:
            conn.execute(
                """UPDATE emails
                   SET syntax_valid=?, is_disposable=?, is_role=?, is_freemail=?,
                       mx_status=?, tier=?, verified_at=?
                   WHERE id=?""",
                (
                    int(syntax_ok), int(disposable), int(role), int(freemail),
                    mx, tier, datetime.now(timezone.utc).isoformat(),
                    email_id,
                ),
            )

        if tier == "error":
            counts["error"] += 1
        else:
            counts[tier] += 1

    logger.info(
        "[VERIFY] DONE — acceptable=%d  risky=%d  invalid=%d  dns_error=%d  total=%d",
        counts["acceptable"], counts["risky"], counts["invalid"], counts["error"], total,
    )
