"""Email extraction from HTML: mailto hrefs, regex, obfuscation, Cloudflare cfemail."""

import logging
import re

from bs4 import BeautifulSoup

logger = logging.getLogger(__name__)

# RFC-5321-ish pattern — deliberately broad; syntax validation happens in verify.py
_EMAIL_RE = re.compile(
    r"[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}",
    re.IGNORECASE,
)

# "[at]" / "(at)" / " at " etc. obfuscation patterns — [dot] is optional
_AT_RE = re.compile(
    r"[a-zA-Z0-9._%+\-]+\s*(?:\[at\]|\(at\)|\bat\b)\s*[a-zA-Z0-9.\-]+"
    r"(?:\s*(?:\[dot\]|\(dot\)|\bdot\b)\s*[a-zA-Z]{2,}|\.[a-zA-Z]{2,})",
    re.IGNORECASE,
)

# Matches HTML entity–encoded @
_ENTITY_AT_RE = re.compile(r"&#(?:64|x40);", re.IGNORECASE)

# Emails we never want to surface regardless of content
_SKIP_DOMAINS = frozenset(
    [
        "example.com",
        "domain.com",
        "yourdomain.com",
        "email.com",
        "sentry.io",
        "w3.org",
        "schema.org",
        "google.com",
        "googleapis.com",
        "facebook.com",
        "instagram.com",
        "twitter.com",
        "x.com",
        "linkedin.com",
        "wixpress.com",
        "squarespace.com",
        "shopify.com",
        "wordpress.org",
        "wordpress.com",
    ]
)


def _normalise(email: str) -> str:
    return email.strip().lower()


# A ZIP code (optional) immediately followed by a US-style phone number, with
# or without separators. Pages sometimes concatenate "...address, ZIP, phone,
# email..." with no whitespace between them — e.g. an auto-generated SEO meta
# description rendered as "...Austin TX 78702512.355.1557info@site.com" — so
# the ZIP/phone digits get greedily absorbed into the email's local part by
# _EMAIL_RE. This anchors to the *start* of the local part only, so it can't
# affect a real local part that merely contains digits elsewhere in it.
_GLUED_PHONE_PREFIX_RE = re.compile(
    r"^(?:\d{5}(?:-\d{4})?)?"        # optional ZIP or ZIP+4
    r"(?:\(\d{3}\)\s*|\d{3}[.\-]?)"  # area code, with or without parens
    r"\d{3}[.\-]?\d{4}"              # exchange + line number
)


def _strip_glued_phone_prefix(local: str) -> str:
    """Strip a ZIP/phone-number prefix accidentally glued onto a local part.
    Returns *local* unchanged if no such prefix is found, or if stripping it
    would leave nothing behind (safer to keep the odd-looking original than
    to produce an empty local part)."""
    stripped = _GLUED_PHONE_PREFIX_RE.sub("", local, count=1)
    return stripped if stripped else local


def _skip(email: str) -> bool:
    """Return True if the email should be discarded before even reaching verify."""
    try:
        domain = email.split("@", 1)[1].lower()
    except IndexError:
        return True
    if email.startswith("//"):
        return True
    # Match the domain itself or any subdomain of it (e.g. "sentry.wixpress.com"
    # and "sentry-next.wixpress.com" — Sentry DSN keys embedded by Wix sites,
    # which look exactly like emails — must be caught by the "wixpress.com" entry).
    return any(domain == d or domain.endswith("." + d) for d in _SKIP_DOMAINS)


# ---------------------------------------------------------------------------
# Cloudflare __cf_email__ decoding
# ---------------------------------------------------------------------------

def _decode_cfemail(encoded: str) -> str | None:
    """
    Decode a Cloudflare-protected email.
    encoded is a hex string; first byte is XOR key.
    """
    try:
        data = bytes.fromhex(encoded)
    except ValueError:
        return None
    if len(data) < 2:
        return None
    key = data[0]
    result = "".join(chr(b ^ key) for b in data[1:])
    return result if "@" in result else None


# ---------------------------------------------------------------------------
# Public extraction entry-point
# ---------------------------------------------------------------------------

def extract_emails(html: str, page_url: str) -> list[dict[str, str]]:
    """
    Parse *html* and return a list of dicts:
        {"email": str, "method": str}

    method is one of: mailto | regex | obfuscated | cfemail
    """
    results: dict[str, str] = {}  # email -> method (first win)

    soup = BeautifulSoup(html, "lxml")

    # 1. mailto: hrefs -------------------------------------------------------
    for tag in soup.find_all("a", href=True):
        href: str = tag["href"]
        if href.lower().startswith("mailto:"):
            raw = href[7:].split("?")[0].strip()
            email = _normalise(raw)
            if email and not _skip(email):
                results.setdefault(email, "mailto")

    # 2. Cloudflare cfemail --------------------------------------------------
    # <a href="/cdn-cgi/l/email-protection" class="__cf_email__" data-cfemail="...">
    for tag in soup.find_all(attrs={"data-cfemail": True}):
        decoded = _decode_cfemail(tag["data-cfemail"])
        if decoded:
            email = _normalise(decoded)
            if not _skip(email):
                results.setdefault(email, "cfemail")

    # Also catch inline script: __cf_email__ hex literals
    for script in soup.find_all("script"):
        text = script.get_text()
        for match in re.finditer(r'__cf_email__.*?"([0-9a-f]{6,})"', text, re.IGNORECASE):
            decoded = _decode_cfemail(match.group(1))
            if decoded:
                email = _normalise(decoded)
                if not _skip(email):
                    results.setdefault(email, "cfemail")

    # 3. Normalise HTML entities then run obfuscation patterns ---------------
    text_blob = soup.get_text(" ")
    # Replace &#64; / &#x40; with @
    text_blob = _ENTITY_AT_RE.sub("@", text_blob)

    for match in _AT_RE.finditer(text_blob):
        raw = match.group(0)
        cleaned = (
            raw.replace("[at]", "@")
            .replace("(at)", "@")
            .replace("[dot]", ".")
            .replace("(dot)", ".")
            .replace(" at ", "@")
            .replace(" dot ", ".")
        )
        # Collapse any remaining whitespace
        cleaned = re.sub(r"\s+", "", cleaned)
        email = _normalise(cleaned)
        if "@" in email:
            local, _, domain_part = email.partition("@")
            email = f"{_strip_glued_phone_prefix(local)}@{domain_part}"
        if _EMAIL_RE.fullmatch(email) and not _skip(email):
            results.setdefault(email, "obfuscated")

    # 4. Raw regex over full page text + all attribute values ----------------
    # Include attribute values to catch emails in data-* or title attributes
    # (e.g. auto-generated SEO <meta name="description"> content, which is
    # also where address/phone digits most often end up glued directly onto
    # an email with no separator — see _strip_glued_phone_prefix).
    full_text = html
    for match in _EMAIL_RE.finditer(full_text):
        email = _normalise(match.group(0))
        local, _, domain_part = email.partition("@")
        email = f"{_strip_glued_phone_prefix(local)}@{domain_part}"
        if not _skip(email):
            results.setdefault(email, "regex")

    found = [{"email": e, "method": m} for e, m in results.items()]
    logger.debug("Extracted %d email(s) from %s", len(found), page_url)
    return found
