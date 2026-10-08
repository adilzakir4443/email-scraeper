"""
Central Playwright browser manager, shared by every directory scraper that
needs a real rendered browser instead of a plain httpx GET — sites that
block httpx/proxy traffic outright (YellowPages, Yelp, BBB, Manta) are
generally still reachable through a real Chromium session with human-like
pacing, since their blocking is keyed off TLS/HTTP fingerprints and request
timing, not just the IP.

Headless by default; pass headless=False (wired to --headed on the CLI) to
watch a session work, e.g. while debugging a selector that stopped matching.
"""

import concurrent.futures
import logging
import random
import time
import urllib.parse

logger = logging.getLogger(__name__)

# How long a single browser-mode scrape is allowed to run in its worker
# thread before giving up — generous, since a slow proxy + multiple page
# loads + human-paced delays can legitimately take a while, but still finite
# so a hung browser can't block the whole DISCOVER/CRAWL stage forever.
_THREAD_TIMEOUT = 120

DEFAULT_VIEWPORT = {"width": 1366, "height": 768}

REALISTIC_VIEWPORTS = [
    {"width": 1366, "height": 768},
    {"width": 1920, "height": 1080},
    {"width": 1440, "height": 900},
    {"width": 1536, "height": 864},
]

USER_AGENTS = [
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/123.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Windows NT 11.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/17.4 Safari/605.1.15",
]

_STEALTH_INIT_SCRIPT = """
    Object.defineProperty(navigator, 'webdriver', {get: () => undefined});
    Object.defineProperty(navigator, 'plugins', {get: () => [1,2,3,4,5]});
    Object.defineProperty(navigator, 'languages', {get: () => ['en-US', 'en']});
    window.chrome = {runtime: {}};
"""


class BrowserSession:
    """Context-manager wrapper around a single Playwright Chromium browser
    + context, pre-configured with a randomized realistic fingerprint
    (viewport, user agent) and basic stealth patches."""

    def __init__(self, headless: bool = True, proxy: dict | None = None, slow_mo: int = 0) -> None:
        self.headless = headless
        self.proxy = proxy  # {"server": "http://host:port", "username": "...", "password": "..."}
        self.slow_mo = slow_mo  # ms delay between actions (0=fast, 50-100=human-like)
        self._pw = None
        self._browser = None
        self._context = None

    def __enter__(self) -> "BrowserSession":
        self.start()
        return self

    def __exit__(self, exc_type, exc_val, exc_tb) -> bool:
        self.stop()
        return False

    def start(self) -> None:
        from playwright.sync_api import sync_playwright

        self._pw = sync_playwright().start()
        launch_args = {
            "headless": self.headless,
            "args": [
                "--no-sandbox",
                "--disable-dev-shm-usage",
                "--disable-blink-features=AutomationControlled",
                "--disable-infobars",
                "--window-size=1366,768",
            ],
            "slow_mo": self.slow_mo,
        }
        if self.proxy:
            launch_args["proxy"] = self.proxy
        self._browser = self._pw.chromium.launch(**launch_args)

        viewport = random.choice(REALISTIC_VIEWPORTS)
        ua = random.choice(USER_AGENTS)

        self._context = self._browser.new_context(
            user_agent=ua,
            viewport=viewport,
            locale="en-US",
            timezone_id="America/Chicago",
            java_script_enabled=True,
            ignore_https_errors=False,
            extra_http_headers={
                "Accept-Language": "en-US,en;q=0.9",
                "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8",
            },
        )

        # Stealth: hide the automation flags headless Chromium exposes by
        # default, which several of these sites check for before blocking.
        self._context.add_init_script(_STEALTH_INIT_SCRIPT)

    def new_page(self):
        return self._context.new_page()

    def stop(self) -> None:
        try:
            if self._context:
                self._context.close()
            if self._browser:
                self._browser.close()
            if self._pw:
                self._pw.stop()
        except Exception:
            pass


def human_delay(min_ms: int = 800, max_ms: int = 2500) -> None:
    """Sleep a random human-like delay in milliseconds."""
    time.sleep(random.uniform(min_ms / 1000, max_ms / 1000))


def human_scroll(page, times: int = 3) -> None:
    """Scroll down the page gradually, like a human reading, rather than
    jumping straight to the bottom (a common bot tell)."""
    for _ in range(times):
        scroll_amount = random.randint(300, 700)
        try:
            page.evaluate(f"window.scrollBy(0, {scroll_amount})")
        except Exception:
            return
        time.sleep(random.uniform(0.3, 0.8))


def human_type(page, selector: str, text: str) -> None:
    """Type text character by character with random per-keystroke speed,
    instead of Playwright's instant fill(), which some sites flag."""
    page.click(selector)
    for char in text:
        page.keyboard.type(char)
        time.sleep(random.uniform(0.05, 0.18))


def safe_goto(page, url: str, timeout: int = 30000) -> bool:
    """Navigate to url; return True on success, False on error."""
    try:
        page.goto(url, timeout=timeout, wait_until="domcontentloaded")
        return True
    except Exception as exc:
        logger.debug("Navigation failed %s: %s", url, exc)
        return False


# Cookie-consent banners that would otherwise sit on top of the content
# these scrapers need to read. Tried in order; the first that matches and
# is visible gets clicked.
_COOKIE_CONSENT_SELECTORS = [
    ".cookie-accept",
    "#accept-cookies",
    "[data-accept]",
    'button:has-text("Accept")',
]


def dismiss_cookie_banner(page) -> bool:
    """Click through a cookie-consent banner if one is present. Returns
    True if something was clicked, False if no known banner was found."""
    for selector in _COOKIE_CONSENT_SELECTORS:
        try:
            locator = page.locator(selector).first
            if locator.count() > 0 and locator.is_visible():
                locator.click(timeout=2000)
                return True
        except Exception:
            continue
    return False


def parse_proxy_for_playwright(proxy_url: str | None) -> dict | None:
    """Convert an http://user:pass@host:port proxy URL (this project's
    PROXY_POOL format) into the {"server", "username", "password"} dict
    Playwright's launch(proxy=...) expects."""
    if not proxy_url:
        return None
    try:
        parsed = urllib.parse.urlparse(proxy_url)
        result = {"server": f"{parsed.scheme}://{parsed.hostname}:{parsed.port}"}
        if parsed.username:
            result["username"] = urllib.parse.unquote(parsed.username)
        if parsed.password:
            result["password"] = urllib.parse.unquote(parsed.password)
        return result
    except Exception:
        return None


def run_in_thread(fn, *args, **kwargs):
    """
    Run a sync-Playwright function in a fresh worker thread.

    Playwright's sync API refuses to start if the calling thread already
    has a running asyncio event loop ("It looks like you are using
    Playwright Sync API inside the asyncio loop") — confirmed live: calling
    BrowserSession directly from inside `asyncio.run(...)` raises exactly
    that error, and running the same call through this helper instead does
    not. A fresh thread has no event loop of its own, so Playwright's
    own internal loop (which it also runs in a dedicated thread) never
    collides with one the caller happens to be running.

    Every BrowserSession usage must be fully self-contained inside fn —
    created, used and torn down in the one thread fn runs in — since
    Playwright's sync objects aren't safe to hand across threads.
    """
    with concurrent.futures.ThreadPoolExecutor(max_workers=1) as executor:
        future = executor.submit(fn, *args, **kwargs)
        return future.result(timeout=_THREAD_TIMEOUT)
