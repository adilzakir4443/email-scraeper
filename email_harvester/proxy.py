"""Proxy rotation from PROXY_POOL env var. One proxy per request, random selection."""

import base64
import logging
import os
import random
import re
import time
import urllib.parse
from typing import Optional

logger = logging.getLogger(__name__)

# Backoff config for 403/429/CAPTCHA responses
BACKOFF_BASE = 2.0
BACKOFF_MAX = 120.0

# A proxy that returns 407 this many times has invalid credentials, not a
# transient/rotation-fixable problem — re-trying it with a different proxy
# from the pool won't help if EVERY proxy shares the same bad credential
# format, so this is surfaced as an error rather than a warning.
_407_THRESHOLD = 3

# Deliberately NOT parsed with urllib.parse.urlparse: the whole point of
# normalizing is to handle a RAW, not-yet-encoded password that may itself
# contain '@', '/' or '#' — characters urlparse treats as URL structure
# (userinfo delimiter, path start, fragment start) and would mis-split on
# before we ever get a chance to re-encode them. A greedy regex split on the
# LAST '@' before the host — the one place '@' is unambiguous in a proxy
# URL, since host:port never contains '@' — sidesteps that chicken-and-egg
# problem entirely.
_PROXY_URL_RE = re.compile(r"^(?P<scheme>[a-zA-Z][\w+.-]*)://(?:(?P<userinfo>.*)@)?(?P<hostport>[^@/]+)/?$")


class ProxyPool:
    """Thread-safe rotating proxy pool loaded from PROXY_POOL env var."""

    def __init__(self, proxies: list[str] | None = None) -> None:
        if proxies is not None:
            raw_pool = [p.strip() for p in proxies if p.strip()]
        else:
            raw = os.environ.get("PROXY_POOL", "")
            raw_pool = [p.strip() for p in raw.split(",") if p.strip()]

        self._pool = [self._normalize_proxy_url(p) for p in raw_pool]

        if not self._pool:
            logger.warning(
                "PROXY_POOL is empty — running without proxies. "
                "Set PROXY_POOL=http://user:pass@host:port,... to enable rotation."
            )
        else:
            logger.info("Loaded %d proxies from PROXY_POOL", len(self._pool))
            for proxy in self._pool:
                logger.debug("[PROXY] Loaded: %s", self._mask(proxy))

        self._failure_counts: dict[str, int] = {}
        self._407_counts: dict[str, int] = {}

    @staticmethod
    def _normalize_proxy_url(proxy: str) -> str:
        """Ensure proxy URL is correctly formatted with encoded credentials.

        A raw, not-yet-encoded password containing '@', '/' or '#' breaks
        urlparse's extraction of host/port — re-encoding here (rather than
        trusting the .env value verbatim) is what actually prevents a whole
        class of 407s caused by a mis-parsed proxy URL. See _PROXY_URL_RE's
        comment for why this can't just call urlparse() to get the pieces.
        """
        match = _PROXY_URL_RE.match(proxy.strip())
        if not match:
            return proxy

        scheme = match.group("scheme") or "http"
        userinfo = match.group("userinfo")
        hostport = match.group("hostport")
        if not userinfo or ":" not in userinfo:
            return proxy

        username, _, password = userinfo.partition(":")
        username = urllib.parse.quote(username, safe="")
        password = urllib.parse.quote(password, safe="")
        return f"{scheme}://{username}:{password}@{hostport}"

    @staticmethod
    def _mask(proxy: str) -> str:
        """Render a proxy URL for logging with its password hidden."""
        try:
            parsed = urllib.parse.urlparse(proxy)
            if parsed.username and parsed.password:
                return (
                    f"{parsed.scheme}://{parsed.username}:***@{parsed.hostname}:{parsed.port}"
                )
        except Exception:
            pass
        return proxy

    @staticmethod
    def _get_proxy_auth_header(proxy: str) -> dict[str, str]:
        """Extract a Basic Proxy-Authorization header from proxy's embedded
        credentials, for manual injection alongside httpx's own handling."""
        try:
            parsed = urllib.parse.urlparse(proxy)
            if parsed.username and parsed.password:
                credentials = f"{urllib.parse.unquote(parsed.username)}:{urllib.parse.unquote(parsed.password)}"
                encoded = base64.b64encode(credentials.encode()).decode()
                return {"Proxy-Authorization": f"Basic {encoded}"}
        except Exception:
            pass
        return {}

    def get_auth_header(self, proxy: Optional[str]) -> dict[str, str]:
        """Auth header for a SPECIFIC proxy (the one actually in use for
        this request) — never re-selects a random proxy internally, since
        that would attach one proxy's credentials to a connection made
        through a different one and manufacture 407s rather than fix them."""
        if not proxy:
            return {}
        return self._get_proxy_auth_header(proxy)

    @property
    def available(self) -> bool:
        return bool(self._pool)

    def get(self) -> Optional[str]:
        """Return a random proxy URL, or None if pool is empty."""
        if not self._pool:
            return None
        return random.choice(self._pool)

    def get_httpx_proxies(self) -> dict[str, str] | None:
        """Return httpx-compatible proxy dict, or None."""
        proxy = self.get()
        if proxy is None:
            return None
        return {"http://": proxy, "https://": proxy}

    def report_failure(self, proxy: str) -> None:
        """Track consecutive failures; log at threshold."""
        self._failure_counts[proxy] = self._failure_counts.get(proxy, 0) + 1
        count = self._failure_counts[proxy]
        if count >= 3:
            logger.warning("Proxy %s has failed %d times consecutively", proxy, count)

    def report_success(self, proxy: str) -> None:
        self._failure_counts.pop(proxy, None)

    def mark_407(self, proxy: Optional[str]) -> None:
        """Track proxies returning 407 — may need a credential fix, not
        rotation, since a bad PROXY_POOL entry returns 407 no matter which
        proxy in the pool it's rotated to next."""
        if not proxy:
            return
        self._407_counts[proxy] = self._407_counts.get(proxy, 0) + 1
        if self._407_counts[proxy] >= _407_THRESHOLD:
            logger.error(
                "Proxy %s has returned 407 %d times. "
                "Check proxy credentials format: should be "
                "http://username:password@host:port",
                self._mask(proxy), self._407_counts[proxy],
            )

    def size(self) -> int:
        return len(self._pool)

    def backoff_sleep(self, attempt: int) -> None:
        """Exponential backoff sleep, capped at BACKOFF_MAX seconds."""
        delay = min(BACKOFF_BASE ** attempt + random.uniform(0, 1), BACKOFF_MAX)
        logger.debug("Backoff sleep %.1fs (attempt %d)", delay, attempt)
        time.sleep(delay)


# Module-level singleton; populated lazily or by CLI.
_pool: ProxyPool | None = None


def get_pool() -> ProxyPool:
    global _pool
    if _pool is None:
        _pool = ProxyPool()
    return _pool


def init_pool(proxies: list[str] | None = None) -> ProxyPool:
    global _pool
    _pool = ProxyPool(proxies)
    return _pool
