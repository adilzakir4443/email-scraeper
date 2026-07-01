"""Proxy rotation from PROXY_POOL env var. One proxy per request, random selection."""

import logging
import os
import random
import time
from typing import Optional

logger = logging.getLogger(__name__)

# Backoff config for 403/429/CAPTCHA responses
BACKOFF_BASE = 2.0
BACKOFF_MAX = 120.0


class ProxyPool:
    """Thread-safe rotating proxy pool loaded from PROXY_POOL env var."""

    def __init__(self, proxies: list[str] | None = None) -> None:
        if proxies is not None:
            self._pool = [p.strip() for p in proxies if p.strip()]
        else:
            raw = os.environ.get("PROXY_POOL", "")
            self._pool = [p.strip() for p in raw.split(",") if p.strip()]

        if not self._pool:
            logger.warning(
                "PROXY_POOL is empty — running without proxies. "
                "Set PROXY_POOL=http://user:pass@host:port,... to enable rotation."
            )
        else:
            logger.info("Loaded %d proxies from PROXY_POOL", len(self._pool))

        self._failure_counts: dict[str, int] = {}

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
