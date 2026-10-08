"""
Tests for the Bing scraper and obfuscated-email regex fixes:

1. email_harvester.discover._scrape_bing originally never incremented
   `pages_fetched` (fixed in 37cbb86), so the `_BING_MAX_PAGES` cap was
   never enforced.
2. _scrape_bing / _parse_bing_listing targeted Bing's classic web-search
   local pack (div.b_localList / b_title / b_phone / etc.), which no
   longer exists — bing.com/search + filters=local_listing:true now just
   returns plain organic results with no structured business data at all.
   The scraper was rewritten to hit bing.com/maps/overlaybfpr instead,
   which embeds each listing's full structured data as JSON in a
   `data-entity` attribute on div.b_maglistcard.
3. email_harvester.extract._AT_RE required a "[dot]"/"(dot)"/"dot" separator
   before the TLD, so obfuscated addresses like "name at example.com"
   (obfuscated "at", plain "." before the TLD) were missed.
4. email_harvester.extract._skip did an exact-match against _SKIP_DOMAINS, so
   Sentry DSN keys shaped like emails (e.g. "<hex>@sentry.wixpress.com",
   embedded by Wix sites) slipped through even though "wixpress.com" was
   already on the skip list — the subdomain never matched exactly. Fixed to
   match the domain or any subdomain of a skipped entry. Also added
   "domain.com"/"yourdomain.com" (common hardcoded form placeholders like
   "user@domain.com") to the skip list.
5. email_harvester.verify._compute_tier had no branch for mx_status=="error"
   (a DNS timeout/failure), so it fell through to "acceptable" — and because
   verified_at got stamped regardless, a transient DNS error permanently
   misclassified the email as verified-good. Fixed so run_verify leaves
   verified_at NULL on an MX error, so the row is retried on the next run.
6. email_harvester.verify._mx_lookup used only the system-configured DNS
   resolver. On some networks (observed on a real machine: a router/ISP
   that silently fails dnspython's raw queries even though normal apps
   resolve fine, and identically in this dev sandbox) EVERY lookup times
   out, so nothing can ever verify as acceptable/risky. Fixed to retry via
   public resolvers (8.8.8.8/1.1.1.1) when the system resolver errors.
7. email_harvester.write.run_write crashed with an unhandled PermissionError
   if the output .xlsx was open elsewhere (e.g. in Excel — a very common
   real workflow: review the last run's output, forget to close it, run
   again), throwing away an entire pipeline run's work. Fixed to fall back
   to a timestamped filename instead of crashing.
8. email_harvester.extract's raw-regex pass scans full HTML including
   attribute values, so when a page concatenates address/phone/email with
   no separator (observed for real: a WordPress SEO plugin's auto-generated
   <meta name="description"> rendered "...Austin TX 78702512.355.1557
   info@ruralrooster.com"), the ZIP+phone digits get greedily absorbed into
   the email's local part by _EMAIL_RE. Added _strip_glued_phone_prefix to
   strip a recognizable ZIP/phone prefix off the start of a local part.
   Also added "email.com" to _SKIP_DOMAINS (observed for real: a newsletter
   form's placeholder="your@email.com" being extracted as a real address).
9. email_harvester.social — new SOCIAL stage that finds Facebook/Instagram/
   LinkedIn links (from a business's own site, or via Bing search fallback).
   Bing wraps every organic-result href in a bing.com/ck/a redirect whose
   real destination is base64-encoded in its "u" parameter (verified live:
   even fully relevant results are never a bare facebook.com/instagram.com
   href), so pattern-matching hrefs directly would never find anything even
   when Bing returns a correct result — _unwrap_bing_redirect decodes it
   first.
10. email_harvester.write.run_write always built a brand-new
    openpyxl.Workbook() and overwrote out_path unconditionally, so
    re-running the tool against the same output file (a common workflow:
    build one master leads list across several niche/location runs)
    silently discarded every row written by earlier runs. Fixed to load
    the existing workbook and append below the old data when the output
    file already exists, deduping by email (Leads) / (name, phone)
    (No Website) so re-running the *same* niche/location doesn't pile up
    duplicates every time.
11. email_harvester.discover._parse_yelp_listing stored the Yelp *listing
    page* URL (e.g. https://www.yelp.com/biz/...) as website_url — not the
    business's actual website, which Yelp's search card rarely shows at
    all. Every Yelp-sourced business therefore looked like it had a site,
    so write.py's "No Website" sheet silently excluded businesses that
    truly have none. Fixed to store None when no real site is found.
12. email_harvester.discover gained 4 more sources — Google Maps (real,
    live-verified via Playwright), YellowPages Canada, Yell.com (UK), and
    Thomson Local (UK) — plus region detection (_detect_region) so
    --location picks only the sources that cover that country. Each source
    now runs in isolation (run_discover catches per-source exceptions) so
    one failing/blocked scraper can't take down the whole DISCOVER stage.
13. email_harvester.write.run_write's append-mode (from the previous fix)
    only ever added rows and never refreshed stale ones — re-running the
    same niche+location wouldn't correct a business whose data had changed
    since the last run. Changed to clear-then-rewrite, but scoped to just
    the niche+location currently being written (via _clear_matching_rows,
    keyed by email for Leads / (name, phone) for No Website, looked up
    through the businesses table so a row stays correctly attributed even
    if it no longer qualifies this run) — rows for any OTHER niche/location
    already in the same output file are left untouched, preserving the
    original append-mode fix's guarantee. Locked-file handling also
    changed from a timestamp suffix to numbered "_1", "_2", ... (checked
    proactively via _resolve_output_path/_is_file_locked before writing,
    with a reactive retry-cascade at save time as a race-condition net).
14. Three data-accuracy fixes:
    - discover.py: added _is_relevant, a keyword-based niche/synonym-map
      check run on every scraped listing (business_name + category)
      before it's persisted — a category that contradicts the niche is
      dropped; no category at all (too little info to judge) is kept but
      flagged relevance_checked=0 rather than silently trusted.
    - resolve.py: a business with no website_url now gets one more check
      — a Bing search for '"<name>" "<city>" official site' — before
      being marked resolve_status='no_site', since many directory
      listings simply omit a site the business actually has. Reuses
      social.py's Bing-search plumbing (_unwrap_bing_redirect etc.).
    - write.py: Leads (Sheet 1) now guards against an empty/NULL email
      row; "No Website" (Sheet 2) now requires resolve_status='no_site'
      (a confirmed absence, not just "not yet checked" — matters once
      resolve.py's search-fallback above means website_url can still be
      NULL while resolve is mid-flight); new Sheet 3 ("Has Website, No
      Email") lists businesses with a site that CRAWL actually ran
      against but found zero emails for — excluding crawl_status=
      'pending' ones, which just haven't been processed yet.
15. New --enrich flag / email_harvester/enrich.py: fills in missing columns
    on an existing Excel file from real sources (website crawl, Bing
    search, YellowPages) — never guesses. Two issues found via live
    testing while building it:
    - Each field's search (website/email/phone/social/address) now runs
      in its own try/except inside _enrich_row. Originally one exception
      (a proxy dropping out mid-run, observed live) inside the function
      discarded every field already found for that business, not just the
      one that failed.
    - _bing_search_social (reused from social.py) has no way to verify a
      social-platform result the way _find_website does (fetch + title
      match) — confirmed live that this let an entirely unrelated
      celebrity's Instagram get attached to a small business. Added
      _social_url_plausible: a cheap local check that the profile URL's
      slug actually contains a word from the business name, applied to
      every Bing-search-fallback social result (not to ones found via the
      business's own site crawl, which don't need it).
16. New --api-fallback flag / email_harvester/google_api.py: Google Places
    (DISCOVER) and Custom Search Engine (RESOLVE website search, SOCIAL
    platform search) fallbacks, activated ONLY when the proxy-based path
    comes back with nothing AND the user opted in with API keys in .env —
    verified live (real HTTP calls with a fake key, showing REQUIRED_DENIED/
    400 handled gracefully) that every function fails soft (returns None/[]
    rather than raising) so a bad/expired key can't crash the pipeline.
    Confirmed via integration test that a successful proxy-based DISCOVER
    run never even imports google_api (API quota only spent once the free
    path has genuinely failed).

No real network calls are made: httpx fetching and DNS resolution are
monkeypatched out.
"""

import base64
import json
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch, MagicMock

import dns.exception
import httpx
from bs4 import BeautifulSoup

from email_harvester import browser, crawl, discover, enrich, google_api, resolve, social, verify, write
from email_harvester.db import init_db, get_conn, upsert_business, upsert_email, upsert_social
from email_harvester.extract import extract_emails


# ---------------------------------------------------------------------------
# Bing Maps listing cards
# ---------------------------------------------------------------------------

def _maglistcard(entity: dict) -> str:
    """Build a `div.b_maglistcard` snippet the way bing.com/maps renders one."""
    payload = json.dumps({"entity": entity}).replace('"', "&quot;")
    return f'<div class="b_maglistcard" data-entity="{payload}"></div>'


_REAL_ENTITY = {
    "title": "Radiant Plumbing, Air Conditioning, & Electrical",
    "id": "ypid:YNEACCB75817F133D6",
    "address": "901 Reinli St, Austin, TX 78751",
    "primaryCategoryName": "HVAC services",
    "phone": "(512) 690-4935",
    "website": "https://radiantplumbing.com/austin/",
}

# A card with no "title" in its entity — _parse_bing_listing must reject it,
# same as a card with no business name in the old markup would have.
_NAMELESS_ENTITY = {
    "address": "123 Nowhere Ln, Austin, TX 78701",
    "phone": "(512) 000-0000",
}


class FakeResponse:
    def __init__(self, text: str, status_code: int = 200):
        self.text = text
        self.status_code = status_code


class ParseBingListingTests(unittest.TestCase):
    def _card(self, html: str):
        from bs4 import BeautifulSoup
        return BeautifulSoup(html, "lxml").select_one("div.b_maglistcard")

    def test_extracts_all_fields_from_data_entity_json(self):
        card = self._card(_maglistcard(_REAL_ENTITY))
        biz = discover._parse_bing_listing(card)
        self.assertEqual(biz, {
            "business_name": "Radiant Plumbing, Air Conditioning, & Electrical",
            "website_url": "https://radiantplumbing.com/austin/",
            "phone": "(512) 690-4935",
            "address": "901 Reinli St, Austin, TX 78751",
            "category": "HVAC services",
        })

    def test_returns_none_when_title_missing(self):
        card = self._card(_maglistcard(_NAMELESS_ENTITY))
        self.assertIsNone(discover._parse_bing_listing(card))

    def test_returns_none_for_malformed_json(self):
        card = self._card('<div class="b_maglistcard" data-entity="not json"></div>')
        self.assertIsNone(discover._parse_bing_listing(card))

    def test_old_local_pack_markup_yields_no_cards(self):
        """The classic local-pack selectors (div.b_localList li, b_title h2,
        ...) no longer correspond to anything Bing renders. Confirms the
        scraper doesn't accidentally still depend on that old structure."""
        from bs4 import BeautifulSoup
        old_style_html = """
        <html><body>
          <div class="b_localList">
            <li><div class="b_title"><h2>Joe's Plumbing</h2></div></li>
          </div>
        </body></html>
        """
        soup = BeautifulSoup(old_style_html, "lxml")
        self.assertEqual(soup.select("div.b_maglistcard[data-entity]"), [])


# ---------------------------------------------------------------------------
# Bing pagination cap
# ---------------------------------------------------------------------------

# A page with a real card container but no parseable business (no "title"
# in the entity), so `collected` never advances and only `pages_fetched`
# can stop the loop.
_EMPTY_CARD_PAGE = f"<html><body>{_maglistcard(_NAMELESS_ENTITY)}</body></html>"


class BingPaginationTests(unittest.TestCase):
    def test_stops_at_max_pages_when_results_never_satisfy_max_results(self):
        """
        Regression test for the missing `pages_fetched += 1`: with an
        unbounded max_results and pages that never yield a business, the
        old code looped forever (or until cards ran out). It must now stop
        after exactly _BING_MAX_PAGES fetches.
        """
        fetch_mock = MagicMock(return_value=FakeResponse(_EMPTY_CARD_PAGE))
        with patch.object(discover, "_fetch_with_retry", fetch_mock), \
             patch.object(discover, "_politeness_sleep", lambda: None), \
             patch.object(discover, "_make_client") as make_client_mock:
            make_client_mock.return_value.__enter__.return_value = MagicMock()
            make_client_mock.return_value.__exit__.return_value = False

            results = list(discover._scrape_bing("plumber", "Austin, TX", max_results=1000))

        self.assertEqual(results, [])
        self.assertEqual(fetch_mock.call_count, discover._BING_MAX_PAGES)

    def test_stops_early_once_max_results_reached(self):
        """Sanity check: the max_results cap still short-circuits before
        the page cap when there are enough businesses on the first page."""
        page_with_business = (
            "<html><body>"
            + _maglistcard(_REAL_ENTITY)
            + _maglistcard({**_REAL_ENTITY, "title": "Austin Pipe Co"})
            + "</body></html>"
        )
        fetch_mock = MagicMock(return_value=FakeResponse(page_with_business))
        with patch.object(discover, "_fetch_with_retry", fetch_mock), \
             patch.object(discover, "_politeness_sleep", lambda: None), \
             patch.object(discover, "_make_client") as make_client_mock:
            make_client_mock.return_value.__enter__.return_value = MagicMock()
            make_client_mock.return_value.__exit__.return_value = False

            results = list(discover._scrape_bing("plumber", "Austin, TX", max_results=1))

        self.assertEqual(len(results), 1)
        self.assertEqual(fetch_mock.call_count, 1)


# ---------------------------------------------------------------------------
# Obfuscated email regex
# ---------------------------------------------------------------------------

class ObfuscatedEmailTests(unittest.TestCase):
    def _extract(self, text: str) -> list[str]:
        html = f"<html><body><p>{text}</p></body></html>"
        return [r["email"] for r in extract_emails(html, "https://example.test")]

    def test_bracket_at_and_bracket_dot(self):
        emails = self._extract("Contact john [at] widgets [dot] com for a quote")
        self.assertIn("john@widgets.com", emails)

    def test_paren_at_and_paren_dot(self):
        emails = self._extract("email jane(at)widgets(dot)com")
        self.assertIn("jane@widgets.com", emails)

    def test_word_at_and_word_dot(self):
        emails = self._extract("reach bob at widgets dot org anytime")
        self.assertIn("bob@widgets.org", emails)

    def test_obfuscated_at_with_plain_dot_tld(self):
        """This is the fix: obfuscated 'at' but a literal '.' before the
        TLD (no '[dot]'/'(dot)'/'dot' token) previously wasn't matched."""
        emails = self._extract("email jane at widgets.com for details")
        self.assertIn("jane@widgets.com", emails)

    def test_bracket_at_with_plain_dot_tld(self):
        emails = self._extract("email jane[at]widgets.com for details")
        self.assertIn("jane@widgets.com", emails)

    def test_plain_email_still_found_via_regex(self):
        emails = self._extract("Reach us at jane@widgets.com directly.")
        self.assertIn("jane@widgets.com", emails)

    def test_mailto_link_found(self):
        html = '<html><body><a href="mailto:info@widgets.com">Email us</a></body></html>'
        emails = [r["email"] for r in extract_emails(html, "https://example.test")]
        self.assertIn("info@widgets.com", emails)

    def test_skip_domains_excluded(self):
        emails = self._extract("test at example.com is a placeholder")
        self.assertNotIn("test@example.com", emails)

    def test_skip_domains_matches_subdomains(self):
        """Regression test for the fix: a Sentry DSN key formatted like an
        email (embedded by Wix sites in <script> tags) must be caught by the
        "wixpress.com" skip entry even though the actual domain is a
        subdomain (sentry.wixpress.com / sentry-next.wixpress.com)."""
        html = (
            "<html><body><script>"
            'sentryConfig.dsn = "https://605a7baede844d278b89dc95ae0a9123'
            '@sentry-next.wixpress.com/12345";'
            'otherConfig.dsn = "https://dd0a55ccb8124b9c9d938e3acf41f8aa'
            '@sentry.wixpress.com/67890";'
            "</script></body></html>"
        )
        emails = [r["email"] for r in extract_emails(html, "https://example.test")]
        self.assertEqual(emails, [])

    def test_domain_com_placeholder_skipped(self):
        emails = self._extract("contact user at domain.com for a quote")
        self.assertNotIn("user@domain.com", emails)

    def test_email_com_placeholder_skipped(self):
        """Regression test: a form input's placeholder="your@email.com"
        (a real newsletter signup form) was being extracted as if it were
        the business's actual contact email."""
        html = (
            '<html><body><input placeholder="your@email.com" '
            'name="emailaddr"></body></html>'
        )
        emails = [r["email"] for r in extract_emails(html, "https://example.test")]
        self.assertNotIn("your@email.com", emails)

    def test_strips_glued_zip_and_phone_from_regex_match(self):
        """Regression test: a WordPress SEO plugin's auto-generated
        <meta name="description"> concatenated the page's address, ZIP,
        phone, and email with no separators at all, e.g.
        "...Austin TX 78702512.355.1557info@ruralrooster.com" — the ZIP and
        phone digits must not end up glued onto the email's local part."""
        html = (
            '<html><head><meta name="description" content='
            '"Contact UsRural Rooster Print &amp; Design3504 E. 4th St'
            'Unit CAustin TX 78702512.355.1557info@ruralrooster.com" />'
            "</head><body></body></html>"
        )
        emails = [r["email"] for r in extract_emails(html, "https://ruralrooster.com/contact")]
        self.assertIn("info@ruralrooster.com", emails)
        self.assertNotIn("78702512.355.1557info@ruralrooster.com", emails)

    def test_strips_glued_phone_without_zip(self):
        emails = self._extract("Call 512.355.1557info@ruralrooster.com now")
        self.assertIn("info@ruralrooster.com", emails)

    def test_does_not_mangle_local_part_with_digits(self):
        """Guard against over-stripping: a legitimate local part that merely
        contains digits (too short to look like a real phone number) must
        pass through unchanged."""
        emails = self._extract("Email: sales2024@realbusiness.com")
        self.assertIn("sales2024@realbusiness.com", emails)

    def test_skip_does_not_over_match_similar_domains(self):
        """Guard against a naive substring fix: "notgoogle.com" must NOT be
        treated as a subdomain of the skipped "google.com"."""
        emails = self._extract("Reach us at jane@notgoogle.com directly.")
        self.assertIn("jane@notgoogle.com", emails)


# ---------------------------------------------------------------------------
# MX-error handling in verify.py
# ---------------------------------------------------------------------------

class VerifyMxErrorTests(unittest.TestCase):
    """DNS resolution is monkeypatched — no real network/DNS calls are made,
    which also matches this environment's actual DNS being unreachable."""

    def setUp(self):
        self._tmpdir = tempfile.TemporaryDirectory()
        self.db_path = str(Path(self._tmpdir.name) / "test.db")
        init_db(self.db_path)
        with get_conn(self.db_path) as conn:
            biz_id = upsert_business(
                conn, niche="plumber", location="Austin, TX",
                business_name="Joe's Plumbing", website_url="https://joesplumbing.com",
                phone=None, address=None, category=None, source="yellowpages",
            )
            upsert_email(conn, business_id=biz_id, email="joe@joesplumbing.com",
                         source_url=None, extract_method="mailto")
            self.email_id = conn.execute(
                "SELECT id FROM emails WHERE email=?", ("joe@joesplumbing.com",)
            ).fetchone()["id"]

    def tearDown(self):
        self._tmpdir.cleanup()

    def _row(self):
        with get_conn(self.db_path) as conn:
            return dict(conn.execute(
                "SELECT * FROM emails WHERE id=?", (self.email_id,)
            ).fetchone())

    def test_mx_error_leaves_row_unverified_for_retry(self):
        """Regression test for the fix: a DNS timeout/error must NOT be
        recorded as a passing tier, and must NOT stamp verified_at, so the
        row is picked up again (WHERE verified_at IS NULL) on the next run."""
        with patch.object(verify, "_mx_lookup", return_value="error"):
            verify.run_verify(self.db_path)

        row = self._row()
        self.assertEqual(row["mx_status"], "error")
        self.assertIsNone(row["tier"])
        self.assertIsNone(row["verified_at"])

    def test_mx_error_then_success_on_retry(self):
        """After a transient DNS error, a later run_verify call (once DNS is
        reachable again) must still be able to verify the same row."""
        with patch.object(verify, "_mx_lookup", return_value="error"):
            verify.run_verify(self.db_path)
        self.assertIsNone(self._row()["verified_at"])

        with patch.object(verify, "_mx_lookup", return_value="acceptable"):
            verify.run_verify(self.db_path)

        row = self._row()
        self.assertEqual(row["tier"], "acceptable")
        self.assertIsNotNone(row["verified_at"])

    def test_mx_acceptable_sets_tier_and_verified_at(self):
        with patch.object(verify, "_mx_lookup", return_value="acceptable"):
            verify.run_verify(self.db_path)

        row = self._row()
        self.assertEqual(row["tier"], "acceptable")
        self.assertIsNotNone(row["verified_at"])

    def test_mx_invalid_sets_tier_invalid(self):
        with patch.object(verify, "_mx_lookup", return_value="invalid"):
            verify.run_verify(self.db_path)

        row = self._row()
        self.assertEqual(row["tier"], "invalid")
        self.assertIsNotNone(row["verified_at"])


# ---------------------------------------------------------------------------
# Public-DNS fallback in _mx_lookup
# ---------------------------------------------------------------------------

class MxFallbackResolverTests(unittest.TestCase):
    def setUp(self):
        verify._mx_cache.clear()

    def tearDown(self):
        verify._mx_cache.clear()

    def test_falls_back_to_public_dns_when_system_resolver_times_out(self):
        """Regression test for the fix: on a network where the system
        resolver silently fails dnspython's queries (observed for real —
        even gmail.com timed out), _mx_lookup must retry via the public-DNS
        fallback instead of giving up immediately."""
        with patch.object(verify._RESOLVER, "resolve", side_effect=dns.exception.Timeout()), \
             patch.object(verify._FALLBACK_RESOLVER, "resolve", return_value=[MagicMock()]):
            status = verify._mx_lookup("example-fallback-test.com")

        self.assertEqual(status, "acceptable")

    def test_reports_error_only_when_both_resolvers_fail(self):
        with patch.object(verify._RESOLVER, "resolve", side_effect=dns.exception.Timeout()), \
             patch.object(verify._FALLBACK_RESOLVER, "resolve", side_effect=dns.exception.Timeout()):
            status = verify._mx_lookup("example-both-fail-test.com")

        self.assertEqual(status, "error")

    def test_does_not_use_fallback_when_system_resolver_succeeds(self):
        """The fallback resolver should only be tried when the system one
        fails — not on every lookup."""
        fallback_resolve = MagicMock()
        with patch.object(verify._RESOLVER, "resolve", return_value=[MagicMock()]), \
             patch.object(verify._FALLBACK_RESOLVER, "resolve", fallback_resolve):
            status = verify._mx_lookup("example-system-ok-test.com")

        self.assertEqual(status, "acceptable")
        fallback_resolve.assert_not_called()


# ---------------------------------------------------------------------------
# Locked-output-file fallback in write.py
# ---------------------------------------------------------------------------

class WritePermissionErrorTests(unittest.TestCase):
    def setUp(self):
        self._tmpdir = tempfile.TemporaryDirectory()
        self.db_path = str(Path(self._tmpdir.name) / "test.db")
        init_db(self.db_path)
        with get_conn(self.db_path) as conn:
            biz_id = upsert_business(
                conn, niche="plumber", location="Austin, TX",
                business_name="Joe's Plumbing", website_url="https://joesplumbing.com",
                phone="555-1234", address="1 Main St", category="Plumbing",
                source="yellowpages",
            )
            upsert_email(conn, business_id=biz_id, email="joe@joesplumbing.com",
                         source_url=None, extract_method="mailto")
            conn.execute(
                "UPDATE emails SET tier='acceptable', mx_status='acceptable' WHERE email=?",
                ("joe@joesplumbing.com",),
            )

    def tearDown(self):
        self._tmpdir.cleanup()

    def test_falls_back_to_numbered_suffix_when_output_is_locked(self):
        """Regression test for the fix: a locked output file (e.g. open in
        Excel — the file that's open when the previous run's results are
        being reviewed) must not crash the whole pipeline and discard a
        completed run's data; it must save under an alternate "_N" name."""
        out_path = str(Path(self._tmpdir.name) / "leads.xlsx")
        real_save = None
        call_count = {"n": 0}

        def fake_save(self_wb, path):
            call_count["n"] += 1
            if call_count["n"] == 1:
                raise PermissionError(13, "Permission denied", path)
            return real_save(self_wb, path)

        import openpyxl
        real_save = openpyxl.Workbook.save
        with patch("openpyxl.Workbook.save", fake_save):
            result_path = write.run_write(
                db_path=self.db_path, niche="plumber", location="Austin, TX",
                out_path=out_path, suppress_path=None,
            )

        self.assertNotEqual(result_path, out_path)
        self.assertTrue(Path(result_path).exists())
        self.assertFalse(Path(out_path).exists())
        self.assertEqual(call_count["n"], 2)

    def test_saves_normally_when_not_locked(self):
        out_path = str(Path(self._tmpdir.name) / "leads.xlsx")
        result_path = write.run_write(
            db_path=self.db_path, niche="plumber", location="Austin, TX",
            out_path=out_path, suppress_path=None,
        )
        self.assertEqual(result_path, out_path)
        self.assertTrue(Path(result_path).exists())


class WriteNoWebsiteSheetTests(unittest.TestCase):
    def setUp(self):
        self._tmpdir = tempfile.TemporaryDirectory()
        self.db_path = str(Path(self._tmpdir.name) / "test.db")
        init_db(self.db_path)

    def tearDown(self):
        self._tmpdir.cleanup()

    def test_no_website_business_appears_only_on_sheet_2(self):
        import openpyxl

        with get_conn(self.db_path) as conn:
            # Has a website + a verified email -> belongs on sheet 1 only.
            biz_with_site = upsert_business(
                conn, niche="plumber", location="Austin, TX",
                business_name="Joe Plumbing", website_url="https://joe.com",
                phone="555-1111", address="1 Main St", category="Plumbing",
                source="yellowpages",
            )
            conn.execute("UPDATE businesses SET normalized_url=? WHERE id=?",
                         ("https://joe.com", biz_with_site))
            upsert_email(conn, business_id=biz_with_site, email="joe@joe.com",
                         source_url=None, extract_method="mailto")
            conn.execute(
                "UPDATE emails SET tier='acceptable', mx_status='acceptable' WHERE business_id=?",
                (biz_with_site,),
            )

            # No website at all -> belongs on sheet 2 only.
            biz_no_site = upsert_business(
                conn, niche="plumber", location="Austin, TX",
                business_name="No Site Plumbing", website_url=None,
                phone="555-2222", address="2 Side St", category="Plumbing",
                source="bing",
            )
            conn.execute(
                "UPDATE businesses SET resolve_status='no_site' WHERE id=?", (biz_no_site,)
            )
            upsert_social(
                conn, business_id=biz_no_site,
                facebook_url="https://facebook.com/nosite",
                instagram_url="https://instagram.com/nosite",
                linkedin_url=None,
            )

        out_path = str(Path(self._tmpdir.name) / "leads.xlsx")
        write.run_write(
            db_path=self.db_path, niche="plumber", location="Austin, TX",
            out_path=out_path, suppress_path=None,
        )

        wb = openpyxl.load_workbook(out_path)
        self.assertEqual(wb.sheetnames, ["Leads", "No Website", "Has Website, No Email"])

        leads_names = [r[0] for r in wb["Leads"].iter_rows(min_row=2, values_only=True)]
        self.assertIn("Joe Plumbing", leads_names)
        self.assertNotIn("No Site Plumbing", leads_names)

        no_site_rows = list(wb["No Website"].iter_rows(min_row=2, values_only=True))
        self.assertEqual(len(no_site_rows), 1)
        row = no_site_rows[0]
        self.assertEqual(
            row,
            ("No Site Plumbing", "555-2222", "Plumbing", "2 Side St",
             "https://facebook.com/nosite", "https://instagram.com/nosite", None),
        )
        self.assertEqual(wb["No Website"][1][0].value, "Company Name")

    def test_no_website_sheet_header_styling_matches_sheet_1(self):
        import openpyxl

        out_path = str(Path(self._tmpdir.name) / "leads.xlsx")
        write.run_write(
            db_path=self.db_path, niche="plumber", location="Austin, TX",
            out_path=out_path, suppress_path=None,
        )
        wb = openpyxl.load_workbook(out_path)
        leads_header = wb["Leads"]["A1"]
        no_site_header = wb["No Website"]["A1"]
        self.assertEqual(no_site_header.fill.start_color.rgb, leads_header.fill.start_color.rgb)
        self.assertEqual(no_site_header.font.bold, leads_header.font.bold)
        self.assertEqual(no_site_header.font.color.rgb, leads_header.font.color.rgb)

    def test_sheet_2_created_and_empty_when_no_such_businesses(self):
        import openpyxl

        out_path = str(Path(self._tmpdir.name) / "leads.xlsx")
        write.run_write(
            db_path=self.db_path, niche="plumber", location="Austin, TX",
            out_path=out_path, suppress_path=None,
        )
        wb = openpyxl.load_workbook(out_path)
        self.assertIn("No Website", wb.sheetnames)
        ws2 = wb["No Website"]
        self.assertEqual(ws2.max_row, 1)  # header only


class WriteAppendModeTests(unittest.TestCase):
    """Regression tests for run_write's in-place-update behavior: it must
    never blindly overwrite/discard out_path. Re-running the *same*
    niche+location refreshes that niche's rows in place (cleared, then
    rewritten fresh from the DB — so stale/changed data doesn't linger);
    a *different* niche+location written to the same file is left alone,
    so a shared output file accumulates across genuinely different runs."""

    def setUp(self):
        self._tmpdir = tempfile.TemporaryDirectory()
        self.db_path = str(Path(self._tmpdir.name) / "test.db")
        self.out_path = str(Path(self._tmpdir.name) / "leads.xlsx")
        init_db(self.db_path)

    def tearDown(self):
        self._tmpdir.cleanup()

    def _add_lead(self, name: str, email: str) -> int:
        with get_conn(self.db_path) as conn:
            biz_id = upsert_business(
                conn, niche="plumber", location="Austin, TX",
                business_name=name, website_url=f"https://{name.lower()}.com",
                phone="555-0000", address="1 Main St", category="Plumbing",
                source="yellowpages",
            )
            conn.execute("UPDATE businesses SET normalized_url=website_url WHERE id=?", (biz_id,))
            upsert_email(conn, business_id=biz_id, email=email,
                         source_url=None, extract_method="mailto")
            conn.execute(
                "UPDATE emails SET tier='acceptable', mx_status='acceptable' WHERE business_id=?",
                (biz_id,),
            )
        return biz_id

    def _add_no_website(self, name: str) -> int:
        with get_conn(self.db_path) as conn:
            biz_id = upsert_business(
                conn, niche="plumber", location="Austin, TX",
                business_name=name, website_url=None,
                phone="555-1111", address="2 Side St", category="Plumbing",
                source="bing",
            )
            conn.execute(
                "UPDATE businesses SET resolve_status='no_site' WHERE id=?", (biz_id,)
            )
            upsert_social(conn, business_id=biz_id,
                          facebook_url=f"https://facebook.com/{name}",
                          instagram_url=None, linkedin_url=None)
        return biz_id

    def test_second_run_preserves_old_rows_and_appends_new_ones(self):
        import openpyxl

        self._add_lead("Alpha Plumbing", "alpha@alpha.com")
        write.run_write(db_path=self.db_path, niche="plumber", location="Austin, TX",
                         out_path=self.out_path, suppress_path=None)

        self._add_lead("Beta Plumbing", "beta@beta.com")
        write.run_write(db_path=self.db_path, niche="plumber", location="Austin, TX",
                         out_path=self.out_path, suppress_path=None)

        wb = openpyxl.load_workbook(self.out_path)
        names = [r[0] for r in wb["Leads"].iter_rows(min_row=2, values_only=True)]
        self.assertEqual(names, ["Alpha Plumbing", "Beta Plumbing"])

    def test_rerunning_same_query_does_not_duplicate_rows(self):
        import openpyxl

        self._add_lead("Alpha Plumbing", "alpha@alpha.com")
        write.run_write(db_path=self.db_path, niche="plumber", location="Austin, TX",
                         out_path=self.out_path, suppress_path=None)
        # Re-run with no new data at all.
        write.run_write(db_path=self.db_path, niche="plumber", location="Austin, TX",
                         out_path=self.out_path, suppress_path=None)

        wb = openpyxl.load_workbook(self.out_path)
        names = [r[0] for r in wb["Leads"].iter_rows(min_row=2, values_only=True)]
        self.assertEqual(names, ["Alpha Plumbing"])  # not duplicated

    def test_no_website_sheet_also_preserves_and_appends(self):
        import openpyxl

        self._add_no_website("Old Co")
        write.run_write(db_path=self.db_path, niche="plumber", location="Austin, TX",
                         out_path=self.out_path, suppress_path=None)

        self._add_no_website("New Co")
        write.run_write(db_path=self.db_path, niche="plumber", location="Austin, TX",
                         out_path=self.out_path, suppress_path=None)

        wb = openpyxl.load_workbook(self.out_path)
        names = [r[0] for r in wb["No Website"].iter_rows(min_row=2, values_only=True)]
        # Order isn't preserved across runs any more (each run rewrites this
        # niche/location's rows fresh, sorted by business_name) — both rows
        # being present at all is what this test actually guards.
        self.assertEqual(set(names), {"Old Co", "New Co"})

    def test_same_niche_rerun_refreshes_stale_row(self):
        """Regression test for the fix: a business that no longer qualifies
        on a re-run (e.g. its email dropped below acceptable/risky tier)
        must not linger in the sheet forever — the niche's rows are cleared
        and rewritten fresh each run, not just added to."""
        import openpyxl

        biz_id = self._add_lead("Alpha Plumbing", "alpha@alpha.com")
        write.run_write(db_path=self.db_path, niche="plumber", location="Austin, TX",
                         out_path=self.out_path, suppress_path=None)

        with get_conn(self.db_path) as conn:
            conn.execute("UPDATE emails SET tier='invalid' WHERE business_id=?", (biz_id,))
        write.run_write(db_path=self.db_path, niche="plumber", location="Austin, TX",
                         out_path=self.out_path, suppress_path=None)

        wb = openpyxl.load_workbook(self.out_path)
        names = [r[0] for r in wb["Leads"].iter_rows(min_row=2, values_only=True)]
        self.assertEqual(names, [])

    def test_different_niche_in_same_file_is_not_touched(self):
        """Regression test for the exact scenario that motivated the
        append fix in the first place: running a DIFFERENT niche into the
        same output file must not erase the previous niche's rows, even
        though each niche's own rows now get cleared-and-rewritten on
        re-run of THAT niche specifically."""
        import openpyxl

        with get_conn(self.db_path) as conn:
            biz_id = upsert_business(
                conn, niche="screen printing", location="Austin, TX",
                business_name="Beta Printing", website_url="https://beta.com",
                phone="555-9999", address="9 Elm St", category="Printing",
                source="yellowpages",
            )
            conn.execute("UPDATE businesses SET normalized_url=website_url WHERE id=?", (biz_id,))
            upsert_email(conn, business_id=biz_id, email="beta@beta.com",
                         source_url=None, extract_method="mailto")
            conn.execute(
                "UPDATE emails SET tier='acceptable', mx_status='acceptable' WHERE business_id=?",
                (biz_id,),
            )
        write.run_write(db_path=self.db_path, niche="screen printing", location="Austin, TX",
                         out_path=self.out_path, suppress_path=None)

        self._add_lead("Alpha Plumbing", "alpha@alpha.com")
        write.run_write(db_path=self.db_path, niche="plumber", location="Austin, TX",
                         out_path=self.out_path, suppress_path=None)

        # Re-run the plumber niche again — must not disturb screen printing.
        write.run_write(db_path=self.db_path, niche="plumber", location="Austin, TX",
                         out_path=self.out_path, suppress_path=None)

        wb = openpyxl.load_workbook(self.out_path)
        names = [r[0] for r in wb["Leads"].iter_rows(min_row=2, values_only=True)]
        self.assertEqual(set(names), {"Alpha Plumbing", "Beta Printing"})

    def test_corrupt_existing_file_falls_back_to_fresh_workbook(self):
        """If out_path exists but isn't a valid workbook, run_write must not
        crash — it should log a warning and start fresh rather than losing
        the run's results."""
        Path(self.out_path).write_text("not a real xlsx file", encoding="utf-8")
        self._add_lead("Alpha Plumbing", "alpha@alpha.com")

        result_path = write.run_write(
            db_path=self.db_path, niche="plumber", location="Austin, TX",
            out_path=self.out_path, suppress_path=None,
        )

        import openpyxl
        wb = openpyxl.load_workbook(result_path)
        names = [r[0] for r in wb["Leads"].iter_rows(min_row=2, values_only=True)]
        self.assertEqual(names, ["Alpha Plumbing"])


class ClearMatchingRowsTests(unittest.TestCase):
    """Direct tests for the niche/location-scoped row-clearing helper."""

    def _sheet_with_rows(self, rows: list[tuple]):
        import openpyxl
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(["Company Name", "Email"])
        for r in rows:
            ws.append(list(r))
        return ws

    def test_clears_only_matching_rows(self):
        ws = self._sheet_with_rows([
            ("Alpha", "alpha@alpha.com"),
            ("Beta", "beta@beta.com"),
            ("Gamma", "gamma@gamma.com"),
        ])
        cleared = write._clear_matching_rows(ws, [1], {("alpha@alpha.com",), ("gamma@gamma.com",)})
        self.assertEqual(cleared, 2)
        remaining = [r[0] for r in ws.iter_rows(min_row=2, values_only=True)]
        self.assertEqual(remaining, ["Beta"])

    def test_header_row_never_touched(self):
        ws = self._sheet_with_rows([("Alpha", "alpha@alpha.com")])
        write._clear_matching_rows(ws, [1], {("alpha@alpha.com",)})
        self.assertEqual(ws.cell(row=1, column=1).value, "Company Name")

    def test_no_matches_clears_nothing(self):
        ws = self._sheet_with_rows([("Alpha", "alpha@alpha.com")])
        cleared = write._clear_matching_rows(ws, [1], {("nobody@nowhere.com",)})
        self.assertEqual(cleared, 0)
        self.assertEqual(ws.max_row, 2)


class ResolveOutputPathTests(unittest.TestCase):
    """Direct tests for the proactive lock-check / numbered-suffix fallback."""

    def setUp(self):
        self._tmpdir = tempfile.TemporaryDirectory()

    def tearDown(self):
        self._tmpdir.cleanup()

    def _patch_locked(self, *locked_paths: str):
        real_open = open
        locked = {str(p) for p in locked_paths}

        def fake_open(path, mode="r", *a, **kw):
            if str(path) in locked and mode == "r+b":
                raise PermissionError(13, "Permission denied", str(path))
            return real_open(path, mode, *a, **kw)

        return patch("builtins.open", side_effect=fake_open)

    def test_nonexistent_path_is_never_locked(self):
        p = Path(self._tmpdir.name) / "leads.xlsx"
        self.assertFalse(write._is_file_locked(p))
        self.assertEqual(write._resolve_output_path(p), p)

    def test_existing_unlocked_path_returned_unchanged(self):
        p = Path(self._tmpdir.name) / "leads.xlsx"
        p.write_text("x")
        self.assertFalse(write._is_file_locked(p))
        self.assertEqual(write._resolve_output_path(p), p)

    def test_locked_path_falls_back_to_suffix_1(self):
        p = Path(self._tmpdir.name) / "leads.xlsx"
        p.write_text("x")
        with self._patch_locked(str(p)):
            resolved = write._resolve_output_path(p)
        self.assertEqual(resolved, p.with_name("leads_1.xlsx"))

    def test_locked_path_and_locked_suffix_1_falls_back_to_suffix_2(self):
        """Regression test matching the spec's worked example: if
        leads_1.xlsx is also locked, fall through to leads_2.xlsx."""
        p = Path(self._tmpdir.name) / "leads.xlsx"
        p1 = p.with_name("leads_1.xlsx")
        p.write_text("x")
        p1.write_text("x")
        with self._patch_locked(str(p), str(p1)):
            resolved = write._resolve_output_path(p)
        self.assertEqual(resolved, p.with_name("leads_2.xlsx"))


# ---------------------------------------------------------------------------
# social.py: HTML scanning, Bing redirect decoding, DB upsert
# ---------------------------------------------------------------------------

class SocialExtractFromHtmlTests(unittest.TestCase):
    def test_finds_real_links_and_ignores_fb_noise(self):
        html = (
            '<html><body>'
            '<a href="https://www.facebook.com/sharer/sharer.php?u=x">Share</a>'
            '<a href="https://www.facebook.com/plugins/like.php?href=x">Like</a>'
            '<a href="https://www.facebook.com/tr?id=123">pixel</a>'
            '<a href="https://www.facebook.com/RuralRoosterAustin">Follow us</a>'
            '<a href="https://www.instagram.com/ruralrooster/">Instagram</a>'
            '<a href="https://www.linkedin.com/company/rural-rooster">LinkedIn</a>'
            '</body></html>'
        )
        result = social._extract_social_from_html(html)
        self.assertEqual(result, {
            "facebook": "https://www.facebook.com/RuralRoosterAustin",
            "instagram": "https://www.instagram.com/ruralrooster/",
            "linkedin": "https://www.linkedin.com/company/rural-rooster",
        })

    def test_missing_platforms_are_none(self):
        html = '<html><body><p>No social links here</p></body></html>'
        result = social._extract_social_from_html(html)
        self.assertEqual(result, {"facebook": None, "instagram": None, "linkedin": None})


class FbIgnorePathTests(unittest.TestCase):
    def test_ignored_paths_are_filtered_by_segment(self):
        for path in ["sharer", "share", "dialog", "plugins", "tr", "hashtag"]:
            url = f"https://www.facebook.com/{path}?x=1"
            self.assertTrue(social._is_ignored_fb_path(url), f"{path} should be ignored")

    def test_does_not_over_match_as_substring(self):
        """Regression guard: a real page named e.g. "SharersDelightBakery"
        must NOT be filtered just because "sharer" is a substring of it."""
        self.assertFalse(social._is_ignored_fb_path("https://www.facebook.com/SharersDelightBakery"))
        self.assertFalse(social._is_ignored_fb_path("https://www.facebook.com/hashtagheroes"))

    def test_nested_ignored_path_still_caught(self):
        self.assertTrue(social._is_ignored_fb_path("https://www.facebook.com/sharer.php?u=x"))


class BingRedirectDecodeTests(unittest.TestCase):
    def test_decodes_real_destination_from_redirect(self):
        """Regression test: Bing wraps every organic-result href in a
        bing.com/ck/a redirect — verified live that even a fully relevant
        result is never a bare destination href. Without decoding this,
        _bing_search_social could never find anything."""
        real_url = "https://www.facebook.com/RuralRoosterAustin"
        encoded = "a1" + base64.urlsafe_b64encode(real_url.encode()).decode().rstrip("=")
        wrapped = f"https://www.bing.com/ck/a?!&&p=abc123&u={encoded}&ntb=1"
        self.assertEqual(social._unwrap_bing_redirect(wrapped), real_url)

    def test_non_redirect_url_passed_through_unchanged(self):
        url = "https://example.com/foo"
        self.assertEqual(social._unwrap_bing_redirect(url), url)

    def test_malformed_redirect_falls_back_to_original(self):
        bad = "https://www.bing.com/ck/a?u=not-valid-base64!!!"
        # Must not raise — either decodes to something or returns the input.
        result = social._unwrap_bing_redirect(bad)
        self.assertIsInstance(result, str)


class CiteReconstructionTests(unittest.TestCase):
    def test_reconstructs_url_from_breadcrumb_with_scheme(self):
        cite = "https://www.facebook.com › RuralRoosterAustin"
        result = social._reconstruct_from_cite(cite, "facebook.com")
        self.assertEqual(result, "https://www.facebook.com/RuralRoosterAustin")

    def test_reconstructs_url_from_breadcrumb_without_scheme(self):
        cite = "facebook.com › pagename"
        result = social._reconstruct_from_cite(cite, "facebook.com")
        self.assertEqual(result, "https://facebook.com/pagename")

    def test_returns_none_for_unrelated_domain(self):
        cite = "https://dictionary.cambridge.org › dictionary › english › rural"
        self.assertIsNone(social._reconstruct_from_cite(cite, "facebook.com"))


class BingSearchSocialTests(unittest.TestCase):
    """_bing_search_social exercised against synthetic Bing-shaped HTML —
    no real network calls."""

    def _fake_result_page(self, real_url: str) -> str:
        encoded = "a1" + base64.urlsafe_b64encode(real_url.encode()).decode().rstrip("=")
        wrapped = f"https://www.bing.com/ck/a?!&&p=abc123&u={encoded}&ntb=1"
        return f'''
        <html><body><ol id="b_results">
          <li class="b_algo">
            <h2><a href="{wrapped}">Rural Rooster on Facebook</a></h2>
            <div class="b_caption"><cite>{real_url}</cite></div>
          </li>
        </ol></body></html>
        '''

    def test_finds_and_decodes_real_result(self):
        real_url = "https://www.facebook.com/RuralRoosterAustin"
        html = self._fake_result_page(real_url)
        with patch.object(social, "_fetch", return_value=html), \
             patch.object(social, "_make_client") as make_client_mock:
            make_client_mock.return_value.__enter__.return_value = MagicMock()
            make_client_mock.return_value.__exit__.return_value = False
            result = social._bing_search_social("Rural Rooster", "Austin, TX", "facebook")

        self.assertEqual(result, real_url)

    def test_returns_none_when_no_results(self):
        with patch.object(social, "_fetch", return_value="<html><body>no results</body></html>"), \
             patch.object(social, "_make_client") as make_client_mock:
            make_client_mock.return_value.__enter__.return_value = MagicMock()
            make_client_mock.return_value.__exit__.return_value = False
            result = social._bing_search_social("Rural Rooster", "Austin, TX", "facebook")

        self.assertIsNone(result)

    def test_returns_none_when_fetch_fails(self):
        with patch.object(social, "_fetch", return_value=None), \
             patch.object(social, "_make_client") as make_client_mock:
            make_client_mock.return_value.__enter__.return_value = MagicMock()
            make_client_mock.return_value.__exit__.return_value = False
            result = social._bing_search_social("Rural Rooster", "Austin, TX", "facebook")

        self.assertIsNone(result)


class UpsertSocialTests(unittest.TestCase):
    def setUp(self):
        self._tmpdir = tempfile.TemporaryDirectory()
        self.db_path = str(Path(self._tmpdir.name) / "test.db")
        init_db(self.db_path)
        with get_conn(self.db_path) as conn:
            self.biz_id = upsert_business(
                conn, niche="plumber", location="Austin, TX",
                business_name="Joe's Plumbing", website_url="https://joesplumbing.com",
                phone=None, address=None, category=None, source="yellowpages",
            )

    def tearDown(self):
        self._tmpdir.cleanup()

    def _row(self):
        with get_conn(self.db_path) as conn:
            return dict(conn.execute(
                "SELECT * FROM businesses WHERE id=?", (self.biz_id,)
            ).fetchone())

    def test_sets_fields_and_marks_done(self):
        with get_conn(self.db_path) as conn:
            upsert_social(
                conn, business_id=self.biz_id,
                facebook_url="https://facebook.com/joesplumbing",
                instagram_url=None, linkedin_url=None,
            )
        row = self._row()
        self.assertEqual(row["facebook_url"], "https://facebook.com/joesplumbing")
        self.assertIsNone(row["instagram_url"])
        self.assertEqual(row["social_status"], "done")

    def test_does_not_overwrite_existing_value(self):
        """Regression test: upsert_social must only fill NULL fields, never
        overwrite a link already found (e.g. by CRAWL) with a later None."""
        with get_conn(self.db_path) as conn:
            upsert_social(
                conn, business_id=self.biz_id,
                facebook_url="https://facebook.com/joesplumbing",
                instagram_url=None, linkedin_url=None,
            )
        with get_conn(self.db_path) as conn:
            upsert_social(
                conn, business_id=self.biz_id,
                facebook_url=None,
                instagram_url="https://instagram.com/joesplumbing",
                linkedin_url=None,
            )
        row = self._row()
        self.assertEqual(row["facebook_url"], "https://facebook.com/joesplumbing")
        self.assertEqual(row["instagram_url"], "https://instagram.com/joesplumbing")

    def test_marks_done_even_when_nothing_found(self):
        """A business that was searched and came up empty must not be
        retried forever."""
        with get_conn(self.db_path) as conn:
            upsert_social(
                conn, business_id=self.biz_id,
                facebook_url=None, instagram_url=None, linkedin_url=None,
            )
        row = self._row()
        self.assertEqual(row["social_status"], "done")


class RunSocialSkipTests(unittest.TestCase):
    def setUp(self):
        self._tmpdir = tempfile.TemporaryDirectory()
        self.db_path = str(Path(self._tmpdir.name) / "test.db")
        init_db(self.db_path)

    def tearDown(self):
        self._tmpdir.cleanup()

    def test_skips_businesses_already_marked_done(self):
        with get_conn(self.db_path) as conn:
            biz_id = upsert_business(
                conn, niche="plumber", location="Austin, TX",
                business_name="Already Done Co", website_url="https://example-abc.com",
                phone=None, address=None, category=None, source="yellowpages",
            )
            upsert_social(conn, business_id=biz_id, facebook_url=None,
                          instagram_url=None, linkedin_url=None)

        with patch.object(social, "_extract_from_website") as extract_mock, \
             patch.object(social, "_find_social_no_website") as search_mock:
            social.run_social(self.db_path)

        extract_mock.assert_not_called()
        search_mock.assert_not_called()

    def test_processes_pending_business_with_website(self):
        with get_conn(self.db_path) as conn:
            biz_id = upsert_business(
                conn, niche="plumber", location="Austin, TX",
                business_name="Pending Co", website_url="https://example-xyz.com",
                phone=None, address=None, category=None, source="yellowpages",
            )

        with patch.object(social, "_extract_from_website",
                           return_value={"facebook": "https://facebook.com/pendingco",
                                         "instagram": None, "linkedin": None}) as extract_mock, \
             patch.object(social, "_find_social_no_website") as search_mock, \
             patch.object(social, "time"):
            social.run_social(self.db_path)

        extract_mock.assert_called_once()
        search_mock.assert_not_called()

        with get_conn(self.db_path) as conn:
            row = dict(conn.execute(
                "SELECT * FROM businesses WHERE id=?", (biz_id,)
            ).fetchone())
        self.assertEqual(row["facebook_url"], "https://facebook.com/pendingco")
        self.assertEqual(row["social_status"], "done")

    def test_falls_back_to_search_when_website_yields_nothing(self):
        with get_conn(self.db_path) as conn:
            biz_id = upsert_business(
                conn, niche="plumber", location="Austin, TX",
                business_name="No Social Co", website_url="https://example-qrs.com",
                phone=None, address=None, category=None, source="yellowpages",
            )

        with patch.object(social, "_extract_from_website",
                           return_value={"facebook": None, "instagram": None, "linkedin": None}), \
             patch.object(social, "_find_social_no_website",
                           return_value={"facebook": "https://facebook.com/foundviabing",
                                         "instagram": None, "linkedin": None}) as search_mock, \
             patch.object(social, "time"):
            social.run_social(self.db_path)

        search_mock.assert_called_once()

        with get_conn(self.db_path) as conn:
            row = dict(conn.execute(
                "SELECT facebook_url FROM businesses WHERE id=?", (biz_id,)
            ).fetchone())
        self.assertEqual(row["facebook_url"], "https://facebook.com/foundviabing")


# ---------------------------------------------------------------------------
# crawl.py: _crawl_site_static's 3-tuple return, conditional upsert_social
# ---------------------------------------------------------------------------

class CrawlStaticSocialTests(unittest.TestCase):
    def test_merges_social_links_across_pages(self):
        homepage_html = (
            '<html><body><a href="https://www.facebook.com/joesplumbing">FB</a></body></html>'
        )
        contact_html = (
            '<html><body><a href="https://www.instagram.com/joesplumbing/">IG</a></body></html>'
        )

        def fake_fetch(client, url, proxy=None):
            if url == "https://joesplumbing.com":
                return homepage_html
            if url == "https://joesplumbing.com/contact":
                return contact_html
            return None

        with patch.object(crawl, "_fetch_page", side_effect=fake_fetch), \
             patch.object(crawl, "_discover_contact_links", return_value=[]), \
             patch.object(crawl, "_make_client") as make_client_mock, \
             patch.object(crawl, "time"):
            make_client_mock.return_value.__enter__.return_value = MagicMock()
            make_client_mock.return_value.__exit__.return_value = False
            emails, pages, social_links = crawl._crawl_site_static("https://joesplumbing.com")

        self.assertEqual(social_links["facebook"], "https://www.facebook.com/joesplumbing")
        self.assertEqual(social_links["instagram"], "https://www.instagram.com/joesplumbing/")
        self.assertIsNone(social_links["linkedin"])

    def test_returns_empty_social_when_homepage_unreachable(self):
        with patch.object(crawl, "_fetch_page", return_value=None), \
             patch.object(crawl, "_make_client") as make_client_mock:
            make_client_mock.return_value.__enter__.return_value = MagicMock()
            make_client_mock.return_value.__exit__.return_value = False
            emails, pages, social_links = crawl._crawl_site_static("https://unreachable.example")

        self.assertEqual(emails, [])
        self.assertEqual(pages, [])
        self.assertEqual(social_links, {"facebook": None, "instagram": None, "linkedin": None})


class RunCrawlSocialTests(unittest.TestCase):
    def setUp(self):
        self._tmpdir = tempfile.TemporaryDirectory()
        self.db_path = str(Path(self._tmpdir.name) / "test.db")
        init_db(self.db_path)

    def tearDown(self):
        self._tmpdir.cleanup()

    def _make_business(self, name: str, url: str) -> int:
        with get_conn(self.db_path) as conn:
            biz_id = upsert_business(
                conn, niche="plumber", location="Austin, TX",
                business_name=name, website_url=url,
                phone=None, address=None, category=None, source="yellowpages",
            )
            conn.execute(
                "UPDATE businesses SET resolve_status='done', normalized_url=? WHERE id=?",
                (url, biz_id),
            )
        return biz_id

    def test_upsert_social_called_when_social_links_found(self):
        biz_id = self._make_business("Joe's Plumbing", "https://joesplumbing.com")
        social_found = {"facebook": "https://facebook.com/joesplumbing",
                         "instagram": None, "linkedin": None}
        fake_email = {"email": "joe@joesplumbing.com", "method": "mailto",
                      "source_url": "https://joesplumbing.com"}

        with patch.object(crawl, "_crawl_site_static",
                           return_value=([fake_email], ["https://joesplumbing.com"], social_found)), \
             patch.object(crawl, "time"):
            crawl.run_crawl(self.db_path)

        with get_conn(self.db_path) as conn:
            row = dict(conn.execute(
                "SELECT facebook_url, social_status FROM businesses WHERE id=?", (biz_id,)
            ).fetchone())
        self.assertEqual(row["facebook_url"], "https://facebook.com/joesplumbing")
        self.assertEqual(row["social_status"], "done")

    def test_upsert_social_not_called_when_nothing_found(self):
        """Regression test: when static crawl finds no social links, the
        business must stay social_status='pending' (not prematurely marked
        'done') so the dedicated SOCIAL stage can still try the Bing-search
        fallback for it later."""
        biz_id = self._make_business("No Social Co", "https://nosocial.com")
        social_empty = {"facebook": None, "instagram": None, "linkedin": None}
        fake_email = {"email": "info@nosocial.com", "method": "mailto",
                      "source_url": "https://nosocial.com"}

        with patch.object(crawl, "_crawl_site_static",
                           return_value=([fake_email], ["https://nosocial.com"], social_empty)), \
             patch.object(crawl, "time"):
            crawl.run_crawl(self.db_path)

        with get_conn(self.db_path) as conn:
            row = dict(conn.execute(
                "SELECT facebook_url, social_status FROM businesses WHERE id=?", (biz_id,)
            ).fetchone())
        self.assertIsNone(row["facebook_url"])
        self.assertEqual(row["social_status"], "pending")


# ---------------------------------------------------------------------------
# discover.py: region detection, Google Maps parsing, Yelp website_url fix,
# and per-source failure isolation in run_discover
# ---------------------------------------------------------------------------

class DetectRegionTests(unittest.TestCase):
    def test_us_state_abbreviation(self):
        self.assertEqual(discover._detect_region("Austin, TX"), "us")
        self.assertEqual(discover._detect_region("New York, NY"), "us")

    def test_uk_by_abbreviation_and_keyword(self):
        self.assertEqual(discover._detect_region("London, UK"), "uk")
        self.assertEqual(discover._detect_region("Manchester, England"), "uk")

    def test_canada_by_province_and_keyword(self):
        self.assertEqual(discover._detect_region("Toronto, ON"), "ca")
        self.assertEqual(discover._detect_region("Vancouver, BC"), "ca")
        self.assertEqual(discover._detect_region("Toronto, Canada"), "ca")

    def test_unrecognized_location_is_unknown(self):
        self.assertEqual(discover._detect_region("Paris, France"), "unknown")
        self.assertEqual(discover._detect_region("Some Village, XY"), "unknown")


class ParseYelpListingTests(unittest.TestCase):
    def test_website_url_is_none_not_yelp_listing_page(self):
        """Regression test for the fix: a Yelp search card only ever links
        to Yelp's own listing page, which must never be stored as the
        business's website_url — that made every Yelp business look like it
        had a site, hiding genuinely-site-less ones from the "No Website"
        sheet."""
        from bs4 import BeautifulSoup
        html = '''
        <li class="border-color">
          <a class="css-19v1rkv" href="/biz/joes-plumbing-austin">Joe's Plumbing</a>
          <p class="phone">(555) 123-4567</p>
          <address>123 Main St, Austin, TX</address>
          <span class="category">Plumbing</span>
        </li>
        '''
        card = BeautifulSoup(html, "lxml").select_one("li")
        result = discover._parse_yelp_listing(card)
        self.assertEqual(result["business_name"], "Joe's Plumbing")
        self.assertIsNone(result["website_url"])


class ParseGoogleMapsCardTests(unittest.TestCase):
    def _card(self, html: str):
        from bs4 import BeautifulSoup
        return BeautifulSoup(html, "lxml").select_one('div[role="article"]')

    def test_parses_normal_card(self):
        html = '''
        <div role="article">
          <a class="hfpxzc" aria-label="Beyond Wow Plumbing &amp; Drains" href="#"></a>
          <a aria-label="Visit Beyond Wow Plumbing &amp; Drains's website" href="https://beyondwow.com/"></a>
          <div class="W4Efsd"><span aria-label="4.9 stars" role="img"><span>4.9</span></span></div>
          <div class="W4Efsd">
            <div class="W4Efsd"><span><span>Plumber</span></span><span><span aria-hidden="true">·</span><span>3432 Greystone Dr</span></span></div>
            <div class="W4Efsd"><span><span>Closes soon · 6 PM · Opens 7 AM Thu</span></span><span><span aria-hidden="true">·</span><span class="UsdlK">+1 512-601-6173</span></span></div>
          </div>
        </div>
        '''
        result = discover._parse_google_maps_card(self._card(html))
        self.assertEqual(result, {
            "business_name": "Beyond Wow Plumbing & Drains",
            "website_url": "https://beyondwow.com/",
            "phone": "+1 512-601-6173",
            "address": "3432 Greystone Dr",
            "category": "Plumber",
        })

    def test_sponsored_card_excluded(self):
        """Regression test: a paid-ad card's "Visit ... website" link
        points at a /aclk Google Ads redirect, not the real business site
        (confirmed live) — sponsored cards must be dropped entirely."""
        html = '''
        <div role="article">
          <h1 aria-label="Sponsored"></h1>
          <a class="hfpxzc" aria-label="Heritage Handyman LLC" href="#"></a>
          <a aria-label="Visit Heritage Handyman LLC's website" href="/aclk?sa=L&amp;ai=abc"></a>
        </div>
        '''
        self.assertIsNone(discover._parse_google_maps_card(self._card(html)))

    def test_bare_category_with_no_address(self):
        """Regression test: some cards show only a category with no address
        row at all (no "·"); the old code required a "·" to extract
        anything, so this returned category=None too even though the
        category text was right there."""
        html = '''
        <div role="article">
          <a class="hfpxzc" aria-label="Plumb Masters, Inc." href="#"></a>
          <a aria-label="Visit Plumb Masters's website" href="https://m.facebook.com/plumb.masters/"></a>
          <div class="W4Efsd"><span aria-label="4.8 stars" role="img"></span></div>
          <div class="W4Efsd">
            <div class="W4Efsd"><span><span>Plumber</span></span></div>
            <div class="W4Efsd"><span><span>Open 24 hours</span></span><span><span aria-hidden="true">·</span><span class="UsdlK">+1 512-960-0044</span></span></div>
          </div>
        </div>
        '''
        result = discover._parse_google_maps_card(self._card(html))
        self.assertEqual(result["category"], "Plumber")
        self.assertIsNone(result["address"])

    def test_icon_glyph_between_dots_filtered_out(self):
        """Regression test: some cards interleave a private-use-area icon
        glyph character between "·"-joined segments (confirmed live), which
        must not end up glued onto the address text."""
        html = '''
        <div role="article">
          <a class="hfpxzc" aria-label="Rooter-Man Plumbing Austin TX" href="#"></a>
          <a aria-label="Visit Rooter-Man's website" href="http://rooterman.com/austin/"></a>
          <div class="W4Efsd">
            <div class="W4Efsd"><span><span>Plumber  15503 Patrica St</span></span></div>
            <div class="W4Efsd"><span><span class="UsdlK">+1 512-720-7092</span></span></div>
          </div>
        </div>
        '''
        # Rebuild with literal middle-dots the way BeautifulSoup would see them
        html = html.replace("Plumber  15503 Patrica St", "Plumber ·  · 15503 Patrica St")
        result = discover._parse_google_maps_card(self._card(html))
        self.assertEqual(result["category"], "Plumber")
        self.assertEqual(result["address"], "15503 Patrica St")

    def test_no_website_link_found(self):
        html = '''
        <div role="article">
          <a class="hfpxzc" aria-label="No Site Co" href="#"></a>
        </div>
        '''
        result = discover._parse_google_maps_card(self._card(html))
        self.assertEqual(result["business_name"], "No Site Co")
        self.assertIsNone(result["website_url"])

    def test_no_name_returns_none(self):
        html = '<div role="article"><span>nothing here</span></div>'
        self.assertIsNone(discover._parse_google_maps_card(self._card(html)))


class RunDiscoverSourceIsolationTests(unittest.TestCase):
    """run_discover must isolate each source: one raising or returning
    nothing must not stop the others, matching the task's explicit
    requirement that a failing/blocked scraper can't crash the pipeline."""

    def setUp(self):
        self._tmpdir = tempfile.TemporaryDirectory()
        self.db_path = str(Path(self._tmpdir.name) / "test.db")
        init_db(self.db_path)

    def tearDown(self):
        self._tmpdir.cleanup()

    def _biz(self, name: str, source: str) -> dict:
        return {
            "business_name": name, "website_url": None, "phone": "555-0000",
            "address": "1 St", "category": "Plumbing", "source": source,
        }

    def test_one_failing_source_does_not_stop_the_others(self):
        def failing_source(niche, location, budget):
            yield self._biz("Should Not Persist", "yellowpages")
            raise RuntimeError("simulated scraper crash")

        def working_source(niche, location, budget):
            yield self._biz("Good Plumbing", "bing")

        with patch.dict(discover._SOURCE_FACTORIES, {
            "yellowpages": failing_source,
            "bing": working_source,
        }), patch.object(discover, "_REGION_SOURCES", {"us": ["yellowpages", "bing"]}), \
             patch.object(discover, "_politeness_sleep"):
            discover.run_discover(self.db_path, "plumbers", "Austin, TX", max_results=30)

        with get_conn(self.db_path) as conn:
            names = {r["business_name"] for r in conn.execute("SELECT business_name FROM businesses").fetchall()}
        # The row yielded before the crash is still persisted (each row is
        # committed as it's inserted), and the next source still ran.
        self.assertIn("Should Not Persist", names)
        self.assertIn("Good Plumbing", names)

    def test_source_returning_zero_results_does_not_stop_others(self):
        def empty_source(niche, location, budget):
            return iter([])

        def working_source(niche, location, budget):
            yield self._biz("Good Plumbing", "bing")

        with patch.dict(discover._SOURCE_FACTORIES, {
            "yellowpages": empty_source,
            "bing": working_source,
        }), patch.object(discover, "_REGION_SOURCES", {"us": ["yellowpages", "bing"]}), \
             patch.object(discover, "_politeness_sleep"):
            discover.run_discover(self.db_path, "plumbers", "Austin, TX", max_results=30)

        with get_conn(self.db_path) as conn:
            names = {r["business_name"] for r in conn.execute("SELECT business_name FROM businesses").fetchall()}
        self.assertIn("Good Plumbing", names)


class RunDiscoverRegionSelectionTests(unittest.TestCase):
    def setUp(self):
        self._tmpdir = tempfile.TemporaryDirectory()
        self.db_path = str(Path(self._tmpdir.name) / "test.db")
        init_db(self.db_path)

    def tearDown(self):
        self._tmpdir.cleanup()

    def test_uk_location_only_calls_uk_sources(self):
        called = []

        def make_tracker(name):
            def factory(niche, location, budget):
                called.append(name)
                return iter([])
            return factory

        patches = {name: make_tracker(name) for name in discover._SOURCE_FACTORIES}
        with patch.dict(discover._SOURCE_FACTORIES, patches), \
             patch.object(discover, "_politeness_sleep"):
            discover.run_discover(self.db_path, "plumbers", "London, UK", max_results=30)

        self.assertEqual(set(called), {"yell_uk", "thomson_local", "google_maps"})

    def test_unrecognized_location_calls_all_sources(self):
        called = []

        def make_tracker(name):
            def factory(niche, location, budget):
                called.append(name)
                return iter([])
            return factory

        patches = {name: make_tracker(name) for name in discover._SOURCE_FACTORIES}
        with patch.dict(discover._SOURCE_FACTORIES, patches), \
             patch.object(discover, "_politeness_sleep"):
            discover.run_discover(self.db_path, "plumbers", "Paris, France", max_results=30)

        self.assertEqual(set(called), set(discover._ALL_SOURCES))


# ---------------------------------------------------------------------------
# discover.py: niche relevance filtering (_is_relevant / _niche_keywords)
# ---------------------------------------------------------------------------

class NicheKeywordsTests(unittest.TestCase):
    def test_synonym_map_matches_singular_and_plural(self):
        self.assertEqual(
            discover._niche_keywords("plumbers"),
            discover._niche_keywords("plumber"),
        )
        self.assertIn("pipe", discover._niche_keywords("plumbers"))

    def test_unmapped_niche_falls_back_to_stem(self):
        self.assertEqual(discover._niche_keywords("screen printing"), ["scree"])


class IsRelevantTests(unittest.TestCase):
    def test_match_via_business_name_is_confidently_relevant(self):
        self.assertEqual(
            discover._is_relevant("Joe's Plumbing", None, "plumbers"),
            (True, True),
        )
        self.assertEqual(
            discover._is_relevant("Joe's Pipe & Drain Co", None, "plumbers"),
            (True, True),
        )

    def test_match_via_category_is_confidently_relevant(self):
        self.assertEqual(
            discover._is_relevant("Joe & Co", "Plumbing Contractor", "plumbers"),
            (True, True),
        )

    def test_no_category_and_no_name_match_is_kept_but_unchecked(self):
        self.assertEqual(discover._is_relevant("Joe & Co", None, "plumbers"), (True, False))
        self.assertEqual(discover._is_relevant("Joe & Co", "", "plumbers"), (True, False))

    def test_contradicting_category_is_dropped(self):
        """Regression test for the fix: niche=plumbers but category=
        restaurant must be dropped, not silently kept."""
        self.assertEqual(discover._is_relevant("Some Diner", "Restaurant", "plumbers"), (False, True))


class RunDiscoverRelevanceFilterTests(unittest.TestCase):
    def setUp(self):
        self._tmpdir = tempfile.TemporaryDirectory()
        self.db_path = str(Path(self._tmpdir.name) / "test.db")
        init_db(self.db_path)

    def tearDown(self):
        self._tmpdir.cleanup()

    def test_irrelevant_results_filtered_and_logged(self):
        def mixed_source(niche, location, budget):
            yield {"business_name": "Joe's Plumbing", "category": "Plumbing",
                   "website_url": None, "phone": None, "address": None, "source": "yellowpages"}
            yield {"business_name": "Some Diner", "category": "Restaurant",
                   "website_url": None, "phone": None, "address": None, "source": "yellowpages"}

        with patch.dict(discover._SOURCE_FACTORIES, {"yellowpages": mixed_source}), \
             patch.object(discover, "_REGION_SOURCES", {"us": ["yellowpages"]}), \
             patch.object(discover, "_politeness_sleep"), \
             self.assertLogs(discover.logger, level="INFO") as log_capture:
            discover.run_discover(self.db_path, "plumbers", "Austin, TX", max_results=30)

        with get_conn(self.db_path) as conn:
            names = {r["business_name"] for r in conn.execute("SELECT business_name FROM businesses").fetchall()}
        self.assertEqual(names, {"Joe's Plumbing"})
        self.assertTrue(any(
            "Filtered 1 irrelevant results from yellowpages" in msg for msg in log_capture.output
        ))

    def test_relevance_checked_flag_persisted(self):
        def uncertain_source(niche, location, budget):
            yield {"business_name": "Joe & Co", "category": None,
                   "website_url": None, "phone": None, "address": None, "source": "yellowpages"}

        with patch.dict(discover._SOURCE_FACTORIES, {"yellowpages": uncertain_source}), \
             patch.object(discover, "_REGION_SOURCES", {"us": ["yellowpages"]}), \
             patch.object(discover, "_politeness_sleep"):
            discover.run_discover(self.db_path, "plumbers", "Austin, TX", max_results=30)

        with get_conn(self.db_path) as conn:
            row = conn.execute("SELECT relevance_checked FROM businesses").fetchone()
        self.assertEqual(row["relevance_checked"], 0)


# ---------------------------------------------------------------------------
# resolve.py: website search-fallback before marking no_site
# ---------------------------------------------------------------------------

class SearchWebsiteViaBingTests(unittest.TestCase):
    @staticmethod
    def _wrap(real_url: str) -> str:
        encoded = "a1" + base64.urlsafe_b64encode(real_url.encode()).decode().rstrip("=")
        return f"https://www.bing.com/ck/a?!&&p=abc&u={encoded}&ntb=1"

    def test_finds_real_site_and_skips_directory_result(self):
        html = f'''
        <html><body><ol id="b_results">
          <li class="b_algo">
            <h2><a href="{self._wrap("https://www.yelp.com/biz/radiant-plumbing")}">Radiant Plumbing - Yelp</a></h2>
            <div class="b_caption"><p>Reviews on Yelp.</p></div>
            <cite>https://www.yelp.com &rsaquo; biz</cite>
          </li>
          <li class="b_algo">
            <h2><a href="{self._wrap("https://radiantplumbing.com/austin/")}">Radiant Plumbing - Official Site</a></h2>
            <div class="b_caption"><p>Welcome to radiantplumbing.com, Austin's plumbing company.</p></div>
            <cite>https://radiantplumbing.com</cite>
          </li>
        </ol></body></html>
        '''
        with patch.object(resolve, "_fetch", return_value=html), \
             patch.object(resolve, "_make_client") as make_client_mock:
            make_client_mock.return_value.__enter__.return_value = MagicMock()
            make_client_mock.return_value.__exit__.return_value = False
            result = resolve._search_website_via_bing("Radiant Plumbing", "Austin, TX")

        self.assertEqual(result, "https://radiantplumbing.com/austin/")

    def test_returns_none_when_only_directory_results(self):
        html = f'''
        <html><body><ol id="b_results">
          <li class="b_algo">
            <h2><a href="{self._wrap("https://www.facebook.com/radiantplumbing")}">Radiant Plumbing - Facebook</a></h2>
            <div class="b_caption"><p>Follow us on Facebook.</p></div>
            <cite>https://www.facebook.com</cite>
          </li>
        </ol></body></html>
        '''
        with patch.object(resolve, "_fetch", return_value=html), \
             patch.object(resolve, "_make_client") as make_client_mock:
            make_client_mock.return_value.__enter__.return_value = MagicMock()
            make_client_mock.return_value.__exit__.return_value = False
            result = resolve._search_website_via_bing("Radiant Plumbing", "Austin, TX")

        self.assertIsNone(result)

    def test_returns_none_when_domain_not_in_title_or_snippet(self):
        """A found domain that doesn't actually appear in the title/snippet
        text is treated as an unconfirmed/coincidental match, not a hit."""
        real_url = "https://unrelated-site.example/"
        html = f'''
        <html><body><ol id="b_results">
          <li class="b_algo">
            <h2><a href="{self._wrap(real_url)}">Some Other Page</a></h2>
            <div class="b_caption"><p>Nothing about the business name here.</p></div>
            <cite>https://unrelated-site.example</cite>
          </li>
        </ol></body></html>
        '''
        with patch.object(resolve, "_fetch", return_value=html), \
             patch.object(resolve, "_make_client") as make_client_mock:
            make_client_mock.return_value.__enter__.return_value = MagicMock()
            make_client_mock.return_value.__exit__.return_value = False
            result = resolve._search_website_via_bing("Radiant Plumbing", "Austin, TX")

        self.assertIsNone(result)

    def test_returns_none_when_fetch_fails(self):
        with patch.object(resolve, "_fetch", return_value=None), \
             patch.object(resolve, "_make_client") as make_client_mock:
            make_client_mock.return_value.__enter__.return_value = MagicMock()
            make_client_mock.return_value.__exit__.return_value = False
            result = resolve._search_website_via_bing("Radiant Plumbing", "Austin, TX")

        self.assertIsNone(result)


class RunResolveSearchFallbackTests(unittest.TestCase):
    def setUp(self):
        self._tmpdir = tempfile.TemporaryDirectory()
        self.db_path = str(Path(self._tmpdir.name) / "test.db")
        init_db(self.db_path)

    def tearDown(self):
        self._tmpdir.cleanup()

    def _make_business(self, name: str) -> int:
        with get_conn(self.db_path) as conn:
            return upsert_business(
                conn, niche="plumber", location="Austin, TX",
                business_name=name, website_url=None,
                phone="555-0000", address="1 St", category="Plumbing",
                source="bing",
            )

    def test_found_via_search_sets_website_and_falls_through_to_resolve(self):
        biz_id = self._make_business("Radiant Plumbing")

        with patch.object(resolve, "_search_website_via_bing",
                           return_value="https://radiantplumbing.com/austin/"), \
             patch.object(resolve, "_resolve_one",
                           return_value="https://radiantplumbing.com/austin/"), \
             patch.object(resolve, "time"):
            resolve.run_resolve(self.db_path)

        with get_conn(self.db_path) as conn:
            row = dict(conn.execute("SELECT * FROM businesses WHERE id=?", (biz_id,)).fetchone())
        self.assertEqual(row["website_url"], "https://radiantplumbing.com/austin/")
        self.assertEqual(row["normalized_url"], "https://radiantplumbing.com/austin/")
        self.assertEqual(row["resolve_status"], "done")

    def test_not_found_via_search_marks_no_site(self):
        biz_id = self._make_business("No Site Co")

        with patch.object(resolve, "_search_website_via_bing", return_value=None), \
             patch.object(resolve, "time"):
            resolve.run_resolve(self.db_path)

        with get_conn(self.db_path) as conn:
            row = dict(conn.execute("SELECT * FROM businesses WHERE id=?", (biz_id,)).fetchone())
        self.assertIsNone(row["website_url"])
        self.assertEqual(row["resolve_status"], "no_site")

    def test_business_with_website_url_never_hits_search(self):
        with get_conn(self.db_path) as conn:
            upsert_business(
                conn, niche="plumber", location="Austin, TX",
                business_name="Has Site Co", website_url="https://hassite.com",
                phone="555-0000", address="1 St", category="Plumbing",
                source="yellowpages",
            )

        with patch.object(resolve, "_search_website_via_bing") as search_mock, \
             patch.object(resolve, "_resolve_one", return_value="https://hassite.com/"), \
             patch.object(resolve, "time"):
            resolve.run_resolve(self.db_path)

        search_mock.assert_not_called()


# ---------------------------------------------------------------------------
# write.py: stricter Sheet 1/2 filters + new Sheet 3
# ---------------------------------------------------------------------------

class WriteThreeSheetPlacementTests(unittest.TestCase):
    def setUp(self):
        self._tmpdir = tempfile.TemporaryDirectory()
        self.db_path = str(Path(self._tmpdir.name) / "test.db")
        self.out_path = str(Path(self._tmpdir.name) / "leads.xlsx")
        init_db(self.db_path)

    def tearDown(self):
        self._tmpdir.cleanup()

    def _write(self):
        write.run_write(db_path=self.db_path, niche="plumber", location="Austin, TX",
                         out_path=self.out_path, suppress_path=None)

    def _sheet_names_present(self, sheet: str) -> list[str]:
        import openpyxl
        wb = openpyxl.load_workbook(self.out_path)
        return [r[0] for r in wb[sheet].iter_rows(min_row=2, values_only=True)]

    def test_confirmed_no_site_goes_to_sheet_2_unconfirmed_does_not(self):
        with get_conn(self.db_path) as conn:
            confirmed = upsert_business(
                conn, niche="plumber", location="Austin, TX",
                business_name="Confirmed No Site", website_url=None,
                phone="555-1", address="1 St", category="Plumbing", source="bing",
            )
            conn.execute("UPDATE businesses SET resolve_status='no_site' WHERE id=?", (confirmed,))

            upsert_business(
                conn, niche="plumber", location="Austin, TX",
                business_name="Unconfirmed Pending", website_url=None,
                phone="555-2", address="2 St", category="Plumbing", source="bing",
            )
            # resolve_status stays 'pending' — not yet confirmed.

        self._write()
        self.assertEqual(self._sheet_names_present("No Website"), ["Confirmed No Site"])

    def test_crawled_with_zero_emails_goes_to_sheet_3_pending_crawl_does_not(self):
        with get_conn(self.db_path) as conn:
            crawled = upsert_business(
                conn, niche="plumber", location="Austin, TX",
                business_name="Crawled No Email", website_url="https://crawled.com",
                phone="555-3", address="3 St", category="Plumbing", source="yellowpages",
            )
            conn.execute(
                "UPDATE businesses SET normalized_url=website_url, crawl_status='no_emails_static' WHERE id=?",
                (crawled,),
            )

            upsert_business(
                conn, niche="plumber", location="Austin, TX",
                business_name="Pending Crawl", website_url="https://pending.com",
                phone="555-4", address="4 St", category="Plumbing", source="yellowpages",
            )
            # crawl_status stays 'pending' — hasn't been processed yet.

        self._write()
        self.assertEqual(self._sheet_names_present("Has Website, No Email"), ["Crawled No Email"])

    def test_business_with_email_never_appears_in_sheet_3(self):
        with get_conn(self.db_path) as conn:
            biz_id = upsert_business(
                conn, niche="plumber", location="Austin, TX",
                business_name="Has Email Co", website_url="https://hasemail.com",
                phone="555-5", address="5 St", category="Plumbing", source="yellowpages",
            )
            conn.execute(
                "UPDATE businesses SET normalized_url=website_url, crawl_status='done' WHERE id=?",
                (biz_id,),
            )
            upsert_email(conn, business_id=biz_id, email="info@hasemail.com",
                         source_url=None, extract_method="mailto")
            conn.execute(
                "UPDATE emails SET tier='acceptable', mx_status='acceptable' WHERE business_id=?",
                (biz_id,),
            )

        self._write()
        self.assertEqual(self._sheet_names_present("Has Website, No Email"), [])
        self.assertEqual(self._sheet_names_present("Leads"), ["Has Email Co"])

    def test_row_with_empty_email_excluded_from_leads(self):
        """Guard test for the fix: a row with a NULL/empty email must never
        appear in Sheet 1, even if it somehow has an acceptable/risky tier."""
        with get_conn(self.db_path) as conn:
            biz_id = upsert_business(
                conn, niche="plumber", location="Austin, TX",
                business_name="Bad Row Co", website_url="https://badrow.com",
                phone="555-6", address="6 St", category="Plumbing", source="yellowpages",
            )
            conn.execute("UPDATE businesses SET normalized_url=website_url WHERE id=?", (biz_id,))
            conn.execute(
                "INSERT INTO emails (business_id, email, tier, mx_status) VALUES (?, '', 'acceptable', 'acceptable')",
                (biz_id,),
            )

        self._write()
        self.assertEqual(self._sheet_names_present("Leads"), [])

    def test_summary_log_line_reports_all_three_counts(self):
        with get_conn(self.db_path) as conn:
            b1 = upsert_business(
                conn, niche="plumber", location="Austin, TX",
                business_name="Leads Co", website_url="https://leads.com",
                phone="555-7", address="7 St", category="Plumbing", source="yellowpages",
            )
            conn.execute("UPDATE businesses SET normalized_url=website_url WHERE id=?", (b1,))
            upsert_email(conn, business_id=b1, email="info@leads.com",
                         source_url=None, extract_method="mailto")
            conn.execute(
                "UPDATE emails SET tier='acceptable', mx_status='acceptable' WHERE business_id=?", (b1,)
            )

        with self.assertLogs(write.logger, level="INFO") as log_capture:
            self._write()

        self.assertTrue(any(
            "Sheet1 leads=1 Sheet2 no_website=0 Sheet3 has_site_no_email=0" in msg
            for msg in log_capture.output
        ))


# ---------------------------------------------------------------------------
# enrich.py (rewritten: no db_path/suppress_path; Google API fallback
# threaded through every field via api_fallback/google_api_key/google_cx)
# ---------------------------------------------------------------------------

class DetectColumnsTests(unittest.TestCase):
    def test_detects_by_header_name_case_insensitive_any_order(self):
        import openpyxl
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(["WEBSITE", "company name", "Phone", "Email"])
        header = next(ws.iter_rows(min_row=1, max_row=1))
        columns = enrich._detect_columns(header)
        self.assertEqual(columns, {"website": 0, "business_name": 1, "phone": 2, "email": 3})

    def test_name_alias_also_matches_business_name(self):
        import openpyxl
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(["Name"])
        header = next(ws.iter_rows(min_row=1, max_row=1))
        self.assertEqual(enrich._detect_columns(header), {"business_name": 0})

    def test_unrecognized_header_ignored(self):
        import openpyxl
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(["Company Name", "Notes"])
        header = next(ws.iter_rows(min_row=1, max_row=1))
        self.assertEqual(enrich._detect_columns(header), {"business_name": 0})


class IsDirectoryDomainTests(unittest.TestCase):
    def test_known_directory_domains_rejected(self):
        for url in [
            "https://www.yelp.com/biz/joe-plumbing",
            "https://www.instagram.com/joeplumbing",
            "https://www.bbb.org/us/tx/austin/profile/plumber/joe",
        ]:
            self.assertTrue(enrich._is_directory_domain(url), url)

    def test_real_business_site_accepted(self):
        self.assertFalse(enrich._is_directory_domain("https://joeplumbing.com/"))


class FuzzyMatchTests(unittest.TestCase):
    def test_identical_strings_ratio_one(self):
        self.assertEqual(enrich._fuzzy_match("plumbing", "plumbing"), 1.0)

    def test_unrelated_strings_low_ratio(self):
        self.assertLess(enrich._fuzzy_match("plumbing", "xyz"), 0.4)


class SocialUrlPlausibleTests(unittest.TestCase):
    def test_unrelated_celebrity_profile_rejected(self):
        """Regression test for the bug found live: an Instagram result for
        a random celebrity must not be accepted for an unrelated business."""
        self.assertFalse(
            enrich._social_url_plausible("https://www.instagram.com/harrykane/", "Joe Plumbing")
        )

    def test_matching_slug_accepted(self):
        self.assertTrue(
            enrich._social_url_plausible("https://www.facebook.com/JoePlumbingAustin", "Joe Plumbing")
        )

    def test_underscored_slug_accepted(self):
        self.assertTrue(
            enrich._social_url_plausible("https://www.instagram.com/joe_plumbing_atx/", "Joe Plumbing")
        )


class ExtractionHelperTests(unittest.TestCase):
    def test_extract_phone_formats_to_standard_pattern(self):
        result = enrich._extract_phone_from_results(
            [{"title": "", "snippet": "Call us at (512) 555-1234 today"}]
        )
        self.assertEqual(result, "(512) 555-1234")

    def test_format_phone_handles_11_digit_with_country_code(self):
        self.assertEqual(enrich._format_phone("15125551234"), "(512) 555-1234")

    def test_extract_address_matches_street_pattern(self):
        result = enrich._extract_address_from_results(
            [{"title": "", "snippet": "Visit us at 123 Main St Suite 200 Austin TX"}]
        )
        self.assertEqual(result, "123 Main St Suite 200 Austin TX")

    def test_extract_email_filters_false_positives(self):
        self.assertIsNone(enrich._extract_emails_from_results(
            [{"title": "", "snippet": "noreply@joeplumbing.com sent this"}]
        ))

    def test_extract_email_finds_real_after_false_positive(self):
        result = enrich._extract_emails_from_results([
            {"title": "", "snippet": "noreply@joeplumbing.com"},
            {"title": "", "snippet": "contact: sales@joeplumbing.com"},
        ])
        self.assertEqual(result, "sales@joeplumbing.com")


class BingSearchTests(unittest.TestCase):
    @staticmethod
    def _wrap(real_url: str) -> str:
        encoded = "a1" + base64.urlsafe_b64encode(real_url.encode()).decode().rstrip("=")
        return f"https://www.bing.com/ck/a?!&&p=abc&u={encoded}&ntb=1"

    def test_decodes_redirect_and_extracts_snippet(self):
        html = f'''
        <html><body><ol id="b_results">
          <li class="b_algo">
            <h2><a href="{self._wrap("https://joeplumbing.com/")}">Joe Plumbing - Official Site</a></h2>
            <div class="b_caption"><p>Welcome to Joe Plumbing.</p></div>
          </li>
        </ol></body></html>
        '''
        with patch.object(enrich, "_fetch", return_value=html), \
             patch.object(enrich, "_make_client") as mc:
            mc.return_value.__enter__.return_value = MagicMock()
            mc.return_value.__exit__.return_value = False
            results = enrich._bing_search("Joe Plumbing Austin", None)

        self.assertEqual(results, [{
            "url": "https://joeplumbing.com/",
            "title": "Joe Plumbing - Official Site",
            "snippet": "Welcome to Joe Plumbing.",
        }])

    def test_returns_empty_list_on_fetch_failure(self):
        with patch.object(enrich, "_fetch", side_effect=Exception("boom")), \
             patch.object(enrich, "_make_client") as mc:
            mc.return_value.__enter__.return_value = MagicMock()
            mc.return_value.__exit__.return_value = False
            self.assertEqual(enrich._bing_search("q", None), [])

    def test_returns_empty_list_when_no_html(self):
        with patch.object(enrich, "_fetch", return_value=None), \
             patch.object(enrich, "_make_client") as mc:
            mc.return_value.__enter__.return_value = MagicMock()
            mc.return_value.__exit__.return_value = False
            self.assertEqual(enrich._bing_search("q", None), [])


class VerifyWebsiteTests(unittest.TestCase):
    def test_requires_200_and_title_match(self):
        class FakeResp:
            status_code = 200
            text = "<html><title>Joe Plumbing - Home</title></html>"

        with patch.object(enrich.httpx, "get", return_value=FakeResp()):
            self.assertTrue(enrich._verify_website("https://joeplumbing.com/", "Joe Plumbing"))

    def test_rejects_non_200(self):
        class FakeResp:
            status_code = 404
            text = ""

        with patch.object(enrich.httpx, "get", return_value=FakeResp()):
            self.assertFalse(enrich._verify_website("https://joeplumbing.com/", "Joe Plumbing"))

    def test_rejects_unrelated_title(self):
        class FakeResp:
            status_code = 200
            text = "<html><title>Wikipedia - Free Encyclopedia</title></html>"

        with patch.object(enrich.httpx, "get", return_value=FakeResp()):
            self.assertFalse(enrich._verify_website("https://example.com/", "Radiant Plumbing"))


class FindWebsiteTests(unittest.TestCase):
    def test_skips_directory_domain_verifies_real_site(self):
        with patch.object(enrich, "_bing_search", return_value=[
            {"url": "https://www.yelp.com/biz/joe-plumbing", "title": "Yelp", "snippet": ""},
            {"url": "https://joeplumbing.com/", "title": "Joe Plumbing", "snippet": ""},
        ]), patch.object(enrich, "_verify_website", side_effect=lambda url, name: "joeplumbing" in url):
            result = enrich._find_website("Joe Plumbing", "Austin", None, False, None, None)

        self.assertEqual(result, "https://joeplumbing.com/")

    def test_returns_none_when_nothing_verified(self):
        with patch.object(enrich, "_bing_search", return_value=[]):
            self.assertIsNone(enrich._find_website("Joe Plumbing", "Austin", None, False, None, None))

    def test_api_fallback_not_tried_when_bing_succeeds(self):
        with patch.object(enrich, "_bing_search", return_value=[
            {"url": "https://joeplumbing.com/", "title": "Joe Plumbing", "snippet": ""},
        ]), patch.object(enrich, "_verify_website", return_value=True), \
             patch("email_harvester.google_api.find_website_via_cse") as cse_mock:
            enrich._find_website("Joe Plumbing", "Austin", None, True, "key", "cx")
        cse_mock.assert_not_called()

    def test_api_fallback_tried_when_bing_finds_nothing(self):
        with patch.object(enrich, "_bing_search", return_value=[]), \
             patch("email_harvester.google_api.find_website_via_cse",
                   return_value="https://joeplumbing.com/") as cse_mock, \
             patch.object(enrich, "_verify_website", return_value=True):
            result = enrich._find_website("Joe Plumbing", "Austin", None, True, "key", "cx")
        cse_mock.assert_called_once()
        self.assertEqual(result, "https://joeplumbing.com/")

    def test_api_fallback_not_tried_without_flag(self):
        with patch.object(enrich, "_bing_search", return_value=[]), \
             patch("email_harvester.google_api.find_website_via_cse") as cse_mock:
            enrich._find_website("Joe Plumbing", "Austin", None, False, "key", "cx")
        cse_mock.assert_not_called()


class GetCrawlResultTests(unittest.TestCase):
    def test_caches_result_across_calls(self):
        crawl_calls = []

        def fake_crawl(url):
            crawl_calls.append(url)
            return ([{"email": "info@joeplumbing.com"}], [], {"facebook": "https://facebook.com/joe"})

        with patch.object(enrich, "_crawl_site_static", side_effect=fake_crawl):
            cache = {}
            r1 = enrich._get_crawl_result("https://joeplumbing.com/", cache)
            r2 = enrich._get_crawl_result("https://joeplumbing.com/", cache)

        self.assertEqual(len(crawl_calls), 1)
        self.assertEqual(r1, r2)

    def test_crawl_exception_returns_empty_result_not_raise(self):
        with patch.object(enrich, "_crawl_site_static", side_effect=Exception("boom")):
            emails, social_links = enrich._get_crawl_result("https://joeplumbing.com/", {})
        self.assertEqual(emails, [])
        self.assertEqual(social_links, {})


class FindEmailTests(unittest.TestCase):
    def test_email_and_social_share_one_crawl(self):
        """Regression test: without a shared crawl cache, email + each of
        the 3 social platforms would each independently re-crawl the same
        site — up to 4x redundant fetches per business."""
        crawl_calls = []

        def fake_crawl(url):
            crawl_calls.append(url)
            return (
                [{"email": "info@joeplumbing.com"}], [],
                {"facebook": "https://facebook.com/joeplumbing", "instagram": None, "linkedin": None},
            )

        with patch.object(enrich, "_crawl_site_static", side_effect=fake_crawl):
            cache = {}
            email = enrich._find_email(
                "Joe Plumbing", "Austin", "https://joeplumbing.com/", None, False, None, None, cache
            )
            fb = enrich._find_social(
                "Joe Plumbing", "Austin", "https://joeplumbing.com/", "facebook", None,
                False, None, None, cache,
            )

        self.assertEqual(email, "info@joeplumbing.com")
        self.assertEqual(fb, "https://facebook.com/joeplumbing")
        self.assertEqual(len(crawl_calls), 1)

    def test_prefers_non_role_email_from_crawl(self):
        with patch.object(enrich, "_crawl_site_static", return_value=(
            [{"email": "info@joeplumbing.com"}, {"email": "sarah@joeplumbing.com"}], [], {},
        )):
            result = enrich._find_email(
                "Joe Plumbing", "Austin", "https://joeplumbing.com/", None, False, None, None,
            )
        self.assertEqual(result, "sarah@joeplumbing.com")

    def test_falls_back_to_bing_search_when_no_website(self):
        with patch.object(enrich, "_bing_search", return_value=[
            {"title": "", "snippet": "Contact: info@joeplumbing.com"},
        ]):
            result = enrich._find_email("Joe Plumbing", "Austin", None, None, False, None, None)
        self.assertEqual(result, "info@joeplumbing.com")

    def test_api_fallback_only_tried_when_bing_empty(self):
        with patch.object(enrich, "_bing_search", return_value=[
            {"title": "", "snippet": "info@joeplumbing.com"},
        ]), patch("email_harvester.google_api.google_custom_search") as cse_mock:
            enrich._find_email("Joe Plumbing", "Austin", None, None, True, "key", "cx")
        cse_mock.assert_not_called()


class FindPhoneTests(unittest.TestCase):
    def test_finds_via_bing(self):
        with patch.object(enrich, "_bing_search", return_value=[
            {"title": "", "snippet": "Call (512) 555-1234"},
        ]):
            self.assertEqual(
                enrich._find_phone("Joe Plumbing", "Austin", None, False, None, None),
                "(512) 555-1234",
            )

    def test_api_fallback_used_when_bing_empty(self):
        with patch.object(enrich, "_bing_search", return_value=[]), \
             patch("email_harvester.google_api.google_custom_search",
                   return_value=[{"title": "", "snippet": "(512) 555-9999"}]) as cse_mock:
            result = enrich._find_phone("Joe Plumbing", "Austin", None, True, "key", "cx")
        cse_mock.assert_called_once()
        self.assertEqual(result, "(512) 555-9999")


class FindSocialTests(unittest.TestCase):
    def test_falls_through_bing_then_cse_with_plausibility_check(self):
        """Regression test for the fix: a CSE fallback result for social
        must also pass the plausibility check, same as the Bing path."""
        with patch.object(enrich, "_crawl_site_static", return_value=([], [], {})), \
             patch.object(enrich, "_bing_search_social", return_value=None), \
             patch("email_harvester.google_api.google_custom_search", return_value=[
                 {"title": "", "link": "https://www.instagram.com/harrykane/", "snippet": ""},
             ]):
            result = enrich._find_social(
                "Joe Plumbing", "Austin", "https://joeplumbing.com/", "instagram", None,
                True, "key", "cx",
            )
        self.assertIsNone(result)  # implausible slug rejected even from CSE

    def test_cse_result_accepted_when_plausible(self):
        with patch.object(enrich, "_crawl_site_static", return_value=([], [], {})), \
             patch.object(enrich, "_bing_search_social", return_value=None), \
             patch("email_harvester.google_api.google_custom_search", return_value=[
                 {"title": "", "link": "https://www.instagram.com/joeplumbingatx/", "snippet": ""},
             ]):
            result = enrich._find_social(
                "Joe Plumbing", "Austin", "https://joeplumbing.com/", "instagram", None,
                True, "key", "cx",
            )
        self.assertEqual(result, "https://www.instagram.com/joeplumbingatx/")


class FindAddressTests(unittest.TestCase):
    def test_finds_via_bing(self):
        with patch.object(enrich, "_bing_search", return_value=[
            {"title": "", "snippet": "Located at 123 Main St Austin TX"},
        ]):
            result = enrich._find_address("Joe Plumbing", "Austin", None, False, None, None)
        self.assertEqual(result, "123 Main St Austin TX")

    def test_api_fallback_used_when_bing_empty(self):
        with patch.object(enrich, "_bing_search", return_value=[]), \
             patch("email_harvester.google_api.google_custom_search",
                   return_value=[{"title": "", "snippet": "456 Elm Ave Austin TX"}]) as cse_mock:
            result = enrich._find_address("Joe Plumbing", "Austin", None, True, "key", "cx")
        cse_mock.assert_called_once()
        self.assertEqual(result, "456 Elm Ave Austin TX")


class RunEnrichTests(unittest.TestCase):
    def setUp(self):
        self._tmpdir = tempfile.TemporaryDirectory()
        self.xlsx_path = str(Path(self._tmpdir.name) / "leads.xlsx")

    def tearDown(self):
        self._tmpdir.cleanup()

    def _make_sheet(self, rows: list[list]):
        import openpyxl
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(["Company Name", "Phone", "Category", "Email", "Website",
                   "Facebook", "Instagram", "LinkedIn", "Address"])
        for row in rows:
            ws.append(row)
        wb.save(self.xlsx_path)

    def test_row_with_empty_company_name_is_skipped(self):
        self._make_sheet([["", "", "", "", "", "", "", "", ""]])
        with patch.object(enrich, "_find_website") as fw_mock:
            result_path = enrich.run_enrich(self.xlsx_path, "plumbers", "Austin, TX")
        fw_mock.assert_not_called()
        self.assertEqual(result_path, self.xlsx_path)

    def test_merged_cell_in_target_column_is_skipped_not_fatal(self):
        """Regression test: MergedCell.value is read-only in openpyxl —
        writing a found value into a column that happens to be merged for
        this row used to raise AttributeError and abort the whole run."""
        import openpyxl
        self._make_sheet([[
            "Joe Plumbing", "", "Plumbing", "", "", "", "", "", "",
        ]])
        wb = openpyxl.load_workbook(self.xlsx_path)
        ws = wb.active
        ws.merge_cells("C2:D2")  # Email column (D) is the merge's 2nd cell, not its anchor
        wb.save(self.xlsx_path)

        with patch.object(enrich, "_find_website", return_value=None), \
             patch.object(enrich, "_find_email", return_value="found@joeplumbing.com"), \
             patch.object(enrich, "_find_phone", return_value="(512) 555-9999"), \
             patch.object(enrich, "_find_social", return_value=None), \
             patch.object(enrich, "_find_address", return_value=None), \
             patch.object(enrich.time, "sleep"):
            result_path = enrich.run_enrich(self.xlsx_path, "plumbers", "Austin, TX")

        wb2 = openpyxl.load_workbook(result_path)
        ws2 = wb2.active
        row = {ws2.cell(row=1, column=c).value: ws2.cell(row=2, column=c).value for c in range(1, ws2.max_column + 1)}
        self.assertIsNone(row["Email"])  # skipped — merged cell, can't write
        self.assertEqual(row["Phone"], "(512) 555-9999")  # unrelated field still filled

    def test_row_with_nothing_missing_is_not_searched(self):
        self._make_sheet([[
            "Joe Plumbing", "555-1234", "Plumbing", "info@joe.com", "https://joe.com",
            "https://facebook.com/joe", "https://instagram.com/joe", "https://linkedin.com/joe", "1 St",
        ]])
        with patch.object(enrich, "_find_website") as fw_mock, \
             patch.object(enrich, "_find_email") as fe_mock:
            enrich.run_enrich(self.xlsx_path, "plumbers", "Austin, TX")
        fw_mock.assert_not_called()
        fe_mock.assert_not_called()

    def test_filled_fields_highlighted_others_untouched(self):
        import openpyxl
        self._make_sheet([[
            "Joe Plumbing", "", "Plumbing", "", "https://joeplumbing.com",
            "", "", "", "123 Main St",
        ]])
        with patch.object(enrich, "_find_email", return_value="info@joeplumbing.com"), \
             patch.object(enrich, "_find_phone", return_value="(512) 555-9999"), \
             patch.object(enrich, "_find_social", return_value=None):
            result_path = enrich.run_enrich(self.xlsx_path, "plumbers", "Austin, TX")

        wb = openpyxl.load_workbook(result_path)
        ws = wb.active
        row = {ws.cell(row=1, column=c).value: ws.cell(row=2, column=c) for c in range(1, ws.max_column + 1)}

        self.assertEqual(row["Phone"].value, "(512) 555-9999")
        self.assertEqual(row["Phone"].fill.start_color.rgb, "00E2EFDA")
        self.assertEqual(row["Email"].value, "info@joeplumbing.com")
        self.assertEqual(row["Email"].fill.start_color.rgb, "00E2EFDA")
        # Pre-existing values must not be re-highlighted.
        self.assertEqual(row["Website"].value, "https://joeplumbing.com")
        self.assertEqual(row["Website"].fill.start_color.rgb, "00000000")
        self.assertEqual(row["Category"].value, "Plumbing")
        self.assertEqual(row["Category"].fill.start_color.rgb, "00000000")

    def test_summary_log_line(self):
        self._make_sheet([[
            "Joe Plumbing", "", "Plumbing", "", "", "", "", "", "",
        ]])
        with patch.object(enrich, "_find_website", return_value=None), \
             patch.object(enrich, "_find_email", return_value=None), \
             patch.object(enrich, "_find_phone", return_value="(512) 555-9999"), \
             patch.object(enrich, "_find_social", return_value=None), \
             patch.object(enrich, "_find_address", return_value=None), \
             patch.object(enrich.time, "sleep"), \
             self.assertLogs(enrich.logger, level="INFO") as log_capture:
            enrich.run_enrich(self.xlsx_path, "plumbers", "Austin, TX")

        self.assertTrue(any(
            "processed=1 enriched=1 fields_filled=1" in msg for msg in log_capture.output
        ))

    def test_missing_company_name_column_raises_value_error(self):
        import openpyxl
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(["Phone", "Email"])
        ws.append(["555-1234", "x@y.com"])
        wb.save(self.xlsx_path)

        with self.assertRaises(ValueError) as ctx:
            enrich.run_enrich(self.xlsx_path, "plumbers", "Austin, TX")
        self.assertIn("Company Name", str(ctx.exception))

    def test_api_fallback_flows_through_to_every_field_helper(self):
        self._make_sheet([[
            "Joe Plumbing", "", "Plumbing", "", "",
            "", "", "", "",
        ]])
        with patch.object(enrich, "_find_website", return_value=None) as fw, \
             patch.object(enrich, "_find_email", return_value=None) as fe, \
             patch.object(enrich, "_find_phone", return_value=None) as fp, \
             patch.object(enrich, "_find_social", return_value=None) as fs, \
             patch.object(enrich, "_find_address", return_value=None) as fa, \
             patch.object(enrich.time, "sleep"):
            enrich.run_enrich(
                self.xlsx_path, "plumbers", "Austin, TX",
                api_fallback=True, google_api_key="key", google_cx="cx", google_places_key="pk",
            )

        for mock_fn in (fw, fe, fp, fa):
            self.assertTrue(mock_fn.call_args.args[-3] is True or True in mock_fn.call_args.args)
        self.assertTrue(fs.called)


class RunEnrichSaveFallbackTests(unittest.TestCase):
    def setUp(self):
        self._tmpdir = tempfile.TemporaryDirectory()

    def tearDown(self):
        self._tmpdir.cleanup()

    def test_locked_file_falls_back_to_numbered_suffix(self):
        import openpyxl
        xlsx_path = str(Path(self._tmpdir.name) / "leads.xlsx")
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(["Company Name"])
        ws.append(["Joe Plumbing"])
        wb.save(xlsx_path)

        wb_loaded = openpyxl.load_workbook(xlsx_path)
        call_count = {"n": 0}
        real_save = openpyxl.Workbook.save

        def fake_save(self_wb, path):
            call_count["n"] += 1
            if call_count["n"] == 1:
                raise PermissionError(13, "Permission denied", path)
            return real_save(self_wb, path)

        with patch("openpyxl.Workbook.save", fake_save):
            result_path = enrich._save_enriched(wb_loaded, xlsx_path)

        self.assertEqual(result_path, str(Path(self._tmpdir.name) / "leads_enriched_1.xlsx"))
        self.assertTrue(Path(result_path).exists())

    def test_cascades_past_existing_enriched_1(self):
        import openpyxl
        xlsx_path = str(Path(self._tmpdir.name) / "leads.xlsx")
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(["Company Name"])
        ws.append(["Joe Plumbing"])
        wb.save(xlsx_path)
        (Path(self._tmpdir.name) / "leads_enriched_1.xlsx").write_text("dummy")

        wb_loaded = openpyxl.load_workbook(xlsx_path)

        def fake_save(self_wb, path):
            if "leads.xlsx" in str(path):
                raise PermissionError(13, "Permission denied", str(path))
            return None  # pretend to succeed without actually writing

        with patch("openpyxl.Workbook.save", fake_save):
            result_path = enrich._save_enriched(wb_loaded, xlsx_path)

        self.assertEqual(result_path, str(Path(self._tmpdir.name) / "leads_enriched_2.xlsx"))


class CliEnrichWiringTests(unittest.TestCase):
    def test_enrich_flag_calls_run_enrich_and_skips_pipeline(self):
        from click.testing import CliRunner
        from email_harvester import cli

        with patch("email_harvester.enrich.run_enrich", return_value="my_leads.xlsx") as run_enrich_mock, \
             patch("email_harvester.proxy.init_pool"), \
             patch("email_harvester.db.init_db"), \
             patch("email_harvester.discover.run_discover") as run_discover_mock, \
             patch.dict("os.environ", {}, clear=True):
            runner = CliRunner()
            result = runner.invoke(cli.main, [
                "--enrich", "my_leads.xlsx", "--niche", "plumbers", "--location", "Austin, TX",
            ])

        self.assertEqual(result.exit_code, 0, result.output)
        run_enrich_mock.assert_called_once_with(
            input_path="my_leads.xlsx", niche="plumbers", location="Austin, TX",
            api_fallback=False, google_api_key=None, google_cx=None, google_places_key=None,
        )
        run_discover_mock.assert_not_called()
        self.assertIn("Enriched file saved to: my_leads.xlsx", result.output)

    def test_missing_company_name_column_shows_clean_error(self):
        from click.testing import CliRunner
        from email_harvester import cli

        with patch("email_harvester.enrich.run_enrich",
                   side_effect=ValueError("No 'Company Name' column found in bad.xlsx")), \
             patch("email_harvester.proxy.init_pool"), \
             patch("email_harvester.db.init_db"):
            runner = CliRunner()
            result = runner.invoke(cli.main, [
                "--enrich", "bad.xlsx", "--niche", "plumbers", "--location", "Austin, TX",
            ])

        self.assertEqual(result.exit_code, 1)
        self.assertIn("No 'Company Name' column found", result.output)


# ---------------------------------------------------------------------------
# google_api.py
# ---------------------------------------------------------------------------

class FakeHttpResponse:
    def __init__(self, json_data, status_code=200):
        self._json = json_data
        self.status_code = status_code

    def json(self):
        return self._json

    def raise_for_status(self):
        if self.status_code >= 400:
            raise httpx.HTTPStatusError("error", request=None, response=self)


class GooglePlacesSearchTests(unittest.TestCase):
    def test_yields_businesses_with_details_merged_in(self):
        textsearch_resp = FakeHttpResponse({
            "status": "OK",
            "results": [
                {"name": "Joe Plumbing", "formatted_address": "1 Main St",
                 "place_id": "abc123", "types": ["plumber", "point_of_interest"]},
            ],
        })
        details_resp = FakeHttpResponse({
            "result": {"formatted_phone_number": "555-1234", "website": "https://joeplumbing.com"},
        })

        with patch.object(google_api.httpx, "get", side_effect=[textsearch_resp, details_resp]):
            results = list(google_api.google_places_search("plumbers", "Austin, TX", "key", max_results=10))

        self.assertEqual(results, [{
            "business_name": "Joe Plumbing",
            "website_url": "https://joeplumbing.com",
            "phone": "555-1234",
            "address": "1 Main St",
            "category": "Plumber",
            "source": "google_places",
        }])

    def test_paginates_via_next_page_token(self):
        page1 = FakeHttpResponse({
            "status": "OK", "next_page_token": "TOKEN2",
            "results": [{"name": "Biz One", "place_id": "p1", "formatted_address": "1 St", "types": []}],
        })
        page2 = FakeHttpResponse({
            "status": "OK",
            "results": [{"name": "Biz Two", "place_id": "p2", "formatted_address": "2 St", "types": []}],
        })
        details = FakeHttpResponse({"result": {"formatted_phone_number": None, "website": None}})

        with patch.object(google_api.httpx, "get", side_effect=[page1, details, page2, details]), \
             patch.object(google_api.time, "sleep"):
            results = list(google_api.google_places_search("plumbers", "Austin, TX", "key", max_results=10))

        self.assertEqual([r["business_name"] for r in results], ["Biz One", "Biz Two"])

    def test_bad_key_returns_empty_without_raising(self):
        resp = FakeHttpResponse({"status": "REQUEST_DENIED", "error_message": "The provided API key is invalid."})
        with patch.object(google_api.httpx, "get", return_value=resp):
            results = list(google_api.google_places_search("plumbers", "Austin, TX", "bad_key"))
        self.assertEqual(results, [])

    def test_network_failure_returns_empty_without_raising(self):
        with patch.object(google_api.httpx, "get", side_effect=httpx.ConnectError("boom")):
            results = list(google_api.google_places_search("plumbers", "Austin, TX", "key"))
        self.assertEqual(results, [])

    def test_respects_max_results(self):
        textsearch_resp = FakeHttpResponse({
            "status": "OK",
            "results": [
                {"name": f"Biz {i}", "place_id": f"p{i}", "formatted_address": "St", "types": []}
                for i in range(5)
            ],
        })
        details = FakeHttpResponse({"result": {"formatted_phone_number": None, "website": None}})

        with patch.object(google_api.httpx, "get", side_effect=[textsearch_resp] + [details] * 2):
            results = list(google_api.google_places_search("plumbers", "Austin, TX", "key", max_results=2))

        self.assertEqual(len(results), 2)


class GoogleCustomSearchTests(unittest.TestCase):
    def test_returns_parsed_results(self):
        resp = FakeHttpResponse({
            "items": [
                {"title": "Joe Plumbing", "link": "https://joeplumbing.com", "snippet": "Plumbing in Austin"},
            ],
        })
        with patch.object(google_api.httpx, "get", return_value=resp):
            results = google_api.google_custom_search("Joe Plumbing Austin", "key", "cx")

        self.assertEqual(results, [
            {"title": "Joe Plumbing", "link": "https://joeplumbing.com", "snippet": "Plumbing in Austin"},
        ])

    def test_429_rate_limit_returns_empty(self):
        resp = FakeHttpResponse({}, status_code=429)
        with patch.object(google_api.httpx, "get", return_value=resp):
            self.assertEqual(google_api.google_custom_search("q", "key", "cx"), [])

    def test_api_error_field_returns_empty(self):
        resp = FakeHttpResponse({"error": {"message": "Invalid Value"}})
        with patch.object(google_api.httpx, "get", return_value=resp):
            self.assertEqual(google_api.google_custom_search("q", "key", "cx"), [])

    def test_network_failure_returns_empty(self):
        with patch.object(google_api.httpx, "get", side_effect=httpx.ConnectError("boom")):
            self.assertEqual(google_api.google_custom_search("q", "key", "cx"), [])

    def test_num_clamped_to_ten(self):
        resp = FakeHttpResponse({"items": []})
        with patch.object(google_api.httpx, "get", return_value=resp) as get_mock:
            google_api.google_custom_search("q", "key", "cx", num=50)
        self.assertEqual(get_mock.call_args.kwargs["params"]["num"], 10)


class FindWebsiteViaCseTests(unittest.TestCase):
    def test_skips_directory_domains_returns_first_real_site(self):
        with patch.object(google_api, "google_custom_search", return_value=[
            {"title": "Yelp", "link": "https://www.yelp.com/biz/joe-plumbing", "snippet": ""},
            {"title": "Joe Plumbing", "link": "https://joeplumbing.com", "snippet": ""},
        ]):
            result = google_api.find_website_via_cse("Joe Plumbing", "Austin, TX", "key", "cx")
        self.assertEqual(result, "https://joeplumbing.com")

    def test_returns_none_when_only_directory_results(self):
        with patch.object(google_api, "google_custom_search", return_value=[
            {"title": "FB", "link": "https://www.facebook.com/joeplumbing", "snippet": ""},
        ]):
            self.assertIsNone(google_api.find_website_via_cse("Joe Plumbing", "Austin, TX", "key", "cx"))

    def test_returns_none_when_no_results(self):
        with patch.object(google_api, "google_custom_search", return_value=[]):
            self.assertIsNone(google_api.find_website_via_cse("Joe Plumbing", "Austin, TX", "key", "cx"))


# ---------------------------------------------------------------------------
# --api-fallback wiring: discover.py / resolve.py / social.py / cli.py
# ---------------------------------------------------------------------------

class RunDiscoverApiFallbackTests(unittest.TestCase):
    def setUp(self):
        self._tmpdir = tempfile.TemporaryDirectory()
        self.db_path = str(Path(self._tmpdir.name) / "test.db")
        init_db(self.db_path)

    def tearDown(self):
        self._tmpdir.cleanup()

    def test_not_called_when_proxy_source_succeeds(self):
        def working_source(niche, location, budget):
            yield {"business_name": "Real Co", "category": "Plumbing",
                   "website_url": None, "phone": None, "address": None, "source": "bing"}

        with patch.dict(discover._SOURCE_FACTORIES, {"bing": working_source}), \
             patch.object(discover, "_REGION_SOURCES", {"us": ["bing"]}), \
             patch.object(discover, "_politeness_sleep"), \
             patch("email_harvester.google_api.google_places_search") as gp_mock:
            discover.run_discover(self.db_path, "plumbers", "Austin, TX", 30,
                                   api_fallback=True, google_places_key="fake_key")

        gp_mock.assert_not_called()

    def test_not_called_when_api_fallback_false(self):
        def failing_source(niche, location, budget):
            raise RuntimeError("403 Forbidden")

        with patch.dict(discover._SOURCE_FACTORIES, {"bing": failing_source}), \
             patch.object(discover, "_REGION_SOURCES", {"us": ["bing"]}), \
             patch.object(discover, "_politeness_sleep"), \
             patch("email_harvester.google_api.google_places_search") as gp_mock:
            discover.run_discover(self.db_path, "plumbers", "Austin, TX", 30,
                                   api_fallback=False, google_places_key="fake_key")

        gp_mock.assert_not_called()

    def test_not_called_when_no_key(self):
        def failing_source(niche, location, budget):
            raise RuntimeError("403 Forbidden")

        with patch.dict(discover._SOURCE_FACTORIES, {"bing": failing_source}), \
             patch.object(discover, "_REGION_SOURCES", {"us": ["bing"]}), \
             patch.object(discover, "_politeness_sleep"), \
             patch("email_harvester.google_api.google_places_search") as gp_mock:
            discover.run_discover(self.db_path, "plumbers", "Austin, TX", 30,
                                   api_fallback=True, google_places_key=None)

        gp_mock.assert_not_called()

    def test_activates_and_persists_when_all_sources_fail(self):
        def failing_source(niche, location, budget):
            raise RuntimeError("403 Forbidden")
            yield  # pragma: no cover — makes this a generator function

        def google_places_fixture(niche, location, api_key, max_results=100):
            yield {"business_name": "Google Found Co", "website_url": "https://gfc.com",
                   "phone": "555-1234", "address": "1 St", "category": "Plumbing",
                   "source": "google_places"}

        with patch.dict(discover._SOURCE_FACTORIES, {"bing": failing_source}), \
             patch.object(discover, "_REGION_SOURCES", {"us": ["bing"]}), \
             patch.object(discover, "_politeness_sleep"), \
             patch("email_harvester.google_api.google_places_search", side_effect=google_places_fixture):
            discover.run_discover(self.db_path, "plumbers", "Austin, TX", 30,
                                   api_fallback=True, google_places_key="fake_key")

        with get_conn(self.db_path) as conn:
            rows = conn.execute("SELECT business_name, source FROM businesses").fetchall()
        self.assertEqual([(r["business_name"], r["source"]) for r in rows],
                          [("Google Found Co", "google_places")])


class RunResolveApiFallbackTests(unittest.TestCase):
    def setUp(self):
        self._tmpdir = tempfile.TemporaryDirectory()
        self.db_path = str(Path(self._tmpdir.name) / "test.db")
        init_db(self.db_path)

    def tearDown(self):
        self._tmpdir.cleanup()

    def _make_business(self, name: str) -> int:
        with get_conn(self.db_path) as conn:
            return upsert_business(
                conn, niche="plumber", location="Austin, TX",
                business_name=name, website_url=None,
                phone="555-0000", address="1 St", category="Plumbing", source="bing",
            )

    def test_cse_not_tried_when_bing_succeeds(self):
        self._make_business("Joe Plumbing")
        with patch.object(resolve, "_search_website_via_bing", return_value="https://bing-found.com"), \
             patch("email_harvester.google_api.find_website_via_cse") as cse_mock, \
             patch.object(resolve, "time"):
            resolve.run_resolve(self.db_path, api_fallback=True,
                                 google_api_key="key", google_cx="cx")
        cse_mock.assert_not_called()

    def test_cse_tried_when_bing_fails_and_fallback_enabled(self):
        self._make_business("Joe Plumbing")
        with patch.object(resolve, "_search_website_via_bing", return_value=None), \
             patch("email_harvester.google_api.find_website_via_cse",
                   return_value="https://cse-found.com") as cse_mock, \
             patch.object(resolve, "time"):
            resolve.run_resolve(self.db_path, api_fallback=True,
                                 google_api_key="key", google_cx="cx")
        cse_mock.assert_called_once()

    def test_cse_not_tried_when_fallback_disabled(self):
        self._make_business("Joe Plumbing")
        with patch.object(resolve, "_search_website_via_bing", return_value=None), \
             patch("email_harvester.google_api.find_website_via_cse") as cse_mock, \
             patch.object(resolve, "time"):
            resolve.run_resolve(self.db_path, api_fallback=False,
                                 google_api_key="key", google_cx="cx")
        cse_mock.assert_not_called()


class BingSearchSocialApiFallbackTests(unittest.TestCase):
    def test_cse_not_tried_when_bing_succeeds(self):
        with patch.object(social, "_fetch", return_value='<li class="b_algo"><h2><a href="https://www.facebook.com/joe">Joe</a></h2></li>'), \
             patch.object(social, "_make_client") as make_client_mock, \
             patch("email_harvester.google_api.google_custom_search") as cse_mock:
            make_client_mock.return_value.__enter__.return_value = MagicMock()
            make_client_mock.return_value.__exit__.return_value = False
            social._bing_search_social("Joe Plumbing", "Austin, TX", "facebook",
                                        api_fallback=True, google_api_key="key", google_cx="cx")
        cse_mock.assert_not_called()

    def test_cse_tried_when_bing_finds_nothing_and_fallback_enabled(self):
        with patch.object(social, "_fetch", return_value="<html></html>"), \
             patch.object(social, "_make_client") as make_client_mock, \
             patch("email_harvester.google_api.google_custom_search",
                   return_value=[{"title": "Joe FB", "link": "https://www.facebook.com/joeplumbing", "snippet": ""}]) as cse_mock:
            make_client_mock.return_value.__enter__.return_value = MagicMock()
            make_client_mock.return_value.__exit__.return_value = False
            result = social._bing_search_social("Joe Plumbing", "Austin, TX", "facebook",
                                                  api_fallback=True, google_api_key="key", google_cx="cx")
        cse_mock.assert_called_once()
        self.assertEqual(result, "https://www.facebook.com/joeplumbing")

    def test_cse_not_tried_when_fallback_disabled(self):
        with patch.object(social, "_fetch", return_value="<html></html>"), \
             patch.object(social, "_make_client") as make_client_mock, \
             patch("email_harvester.google_api.google_custom_search") as cse_mock:
            make_client_mock.return_value.__enter__.return_value = MagicMock()
            make_client_mock.return_value.__exit__.return_value = False
            social._bing_search_social("Joe Plumbing", "Austin, TX", "facebook")
        cse_mock.assert_not_called()


class CliApiFallbackTests(unittest.TestCase):
    def test_missing_key_logs_warning_and_disables_fallback(self):
        from click.testing import CliRunner
        from email_harvester import cli

        # assertLogs hooks the logging module directly rather than relying
        # on CliRunner's captured stdout — logging.basicConfig() is a no-op
        # after its first call in a process, so with 150+ tests invoking
        # cli.main() in this same process, a later test's StreamHandler
        # doesn't reliably end up pointing at that test's own CliRunner
        # stdout. The functional assertion below (api_fallback actually
        # disabled) is the one that matters; this just confirms the message.
        with patch("email_harvester.discover.run_discover") as rd_mock, \
             patch("email_harvester.resolve.run_resolve"), \
             patch("email_harvester.crawl.run_crawl"), \
             patch("email_harvester.social.run_social"), \
             patch("email_harvester.verify.run_verify"), \
             patch("email_harvester.write.run_write", return_value="out.xlsx"), \
             patch("email_harvester.proxy.init_pool"), \
             patch("email_harvester.db.init_db"), \
             patch("email_harvester.db.get_conn"), \
             patch.dict("os.environ", {}, clear=True), \
             self.assertLogs("email_harvester.cli", level="WARNING") as log_capture:
            runner = CliRunner()
            result = runner.invoke(cli.main, [
                "--niche", "plumbers", "--location", "Austin, TX", "--api-fallback",
            ])

        self.assertEqual(result.exit_code, 0, result.output)
        self.assertTrue(any("GOOGLE_API_KEY not found" in msg for msg in log_capture.output))
        self.assertEqual(rd_mock.call_args.kwargs["api_fallback"], False)

    def test_keys_present_passed_through_to_all_stages(self):
        from click.testing import CliRunner
        from email_harvester import cli

        with patch("email_harvester.discover.run_discover") as rd_mock, \
             patch("email_harvester.resolve.run_resolve") as rr_mock, \
             patch("email_harvester.crawl.run_crawl"), \
             patch("email_harvester.social.run_social") as rs_mock, \
             patch("email_harvester.verify.run_verify"), \
             patch("email_harvester.write.run_write", return_value="out.xlsx"), \
             patch("email_harvester.proxy.init_pool"), \
             patch("email_harvester.db.init_db"), \
             patch("email_harvester.db.get_conn"), \
             patch.dict("os.environ", {"GOOGLE_API_KEY": "k1", "GOOGLE_CX": "cx1"}, clear=True):
            runner = CliRunner()
            result = runner.invoke(cli.main, [
                "--niche", "plumbers", "--location", "Austin, TX", "--api-fallback",
            ])

        self.assertEqual(result.exit_code, 0, result.output)
        self.assertEqual(rd_mock.call_args.kwargs["api_fallback"], True)
        self.assertEqual(rd_mock.call_args.kwargs["google_places_key"], "k1")
        self.assertEqual(rr_mock.call_args.kwargs["google_api_key"], "k1")
        self.assertEqual(rr_mock.call_args.kwargs["google_cx"], "cx1")
        self.assertEqual(rs_mock.call_args.kwargs["google_api_key"], "k1")

    def test_separate_places_key_used_when_set(self):
        from click.testing import CliRunner
        from email_harvester import cli

        with patch("email_harvester.discover.run_discover") as rd_mock, \
             patch("email_harvester.resolve.run_resolve"), \
             patch("email_harvester.crawl.run_crawl"), \
             patch("email_harvester.social.run_social"), \
             patch("email_harvester.verify.run_verify"), \
             patch("email_harvester.write.run_write", return_value="out.xlsx"), \
             patch("email_harvester.proxy.init_pool"), \
             patch("email_harvester.db.init_db"), \
             patch("email_harvester.db.get_conn"), \
             patch.dict("os.environ", {
                 "GOOGLE_API_KEY": "k1", "GOOGLE_PLACES_KEY": "places_only_key", "GOOGLE_CX": "cx1",
             }, clear=True):
            runner = CliRunner()
            result = runner.invoke(cli.main, [
                "--niche", "plumbers", "--location", "Austin, TX", "--api-fallback",
            ])

        self.assertEqual(result.exit_code, 0, result.output)
        self.assertEqual(rd_mock.call_args.kwargs["google_places_key"], "places_only_key")

    def test_no_flag_means_api_never_touched(self):
        from click.testing import CliRunner
        from email_harvester import cli

        with patch("email_harvester.discover.run_discover") as rd_mock, \
             patch("email_harvester.resolve.run_resolve"), \
             patch("email_harvester.crawl.run_crawl"), \
             patch("email_harvester.social.run_social"), \
             patch("email_harvester.verify.run_verify"), \
             patch("email_harvester.write.run_write", return_value="out.xlsx"), \
             patch("email_harvester.proxy.init_pool"), \
             patch("email_harvester.db.init_db"), \
             patch("email_harvester.db.get_conn"), \
             patch.dict("os.environ", {"GOOGLE_API_KEY": "k1", "GOOGLE_CX": "cx1"}, clear=True):
            runner = CliRunner()
            result = runner.invoke(cli.main, ["--niche", "plumbers", "--location", "Austin, TX"])

        self.assertEqual(result.exit_code, 0, result.output)
        self.assertEqual(rd_mock.call_args.kwargs["api_fallback"], False)


# ---------------------------------------------------------------------------
# browser.py
# ---------------------------------------------------------------------------

class RunInThreadTests(unittest.TestCase):
    def test_returns_fn_result(self):
        result = browser.run_in_thread(lambda x, y: x + y, 2, 3)
        self.assertEqual(result, 5)

    def test_propagates_exception_from_fn(self):
        def _boom():
            raise ValueError("boom")

        with self.assertRaises(ValueError):
            browser.run_in_thread(_boom)

    def test_works_inside_a_running_asyncio_event_loop(self):
        """Regression test for the actual bug report: Playwright's sync API
        refuses to run if the calling thread already has a running asyncio
        event loop ("It looks like you are using Playwright Sync API inside
        the asyncio loop") — confirmed live. run_in_thread must let a sync
        function run fine even when called from inside asyncio.run(...),
        since it always executes in a fresh thread with no event loop of
        its own."""
        import asyncio

        async def _call_from_event_loop():
            return browser.run_in_thread(lambda: "ran fine")

        result = asyncio.run(_call_from_event_loop())
        self.assertEqual(result, "ran fine")

    def test_times_out_on_a_hung_function(self):
        import time as _time

        def _hang():
            _time.sleep(0.3)

        with patch.object(browser, "_THREAD_TIMEOUT", 0.05):
            with self.assertRaises(Exception):
                browser.run_in_thread(_hang)


class ParseProxyForPlaywrightTests(unittest.TestCase):
    def test_full_proxy_url_converted(self):
        result = browser.parse_proxy_for_playwright("http://user1:pass1@proxy.example.com:8080")
        self.assertEqual(result, {
            "server": "http://proxy.example.com:8080",
            "username": "user1",
            "password": "pass1",
        })

    def test_percent_encoded_credentials_decoded(self):
        result = browser.parse_proxy_for_playwright("http://user1:p%40ss@proxy.example.com:8080")
        self.assertEqual(result["password"], "p@ss")

    def test_no_credentials_omits_username_password(self):
        result = browser.parse_proxy_for_playwright("http://proxy.example.com:3128")
        self.assertEqual(result, {"server": "http://proxy.example.com:3128"})

    def test_none_input_returns_none(self):
        self.assertIsNone(browser.parse_proxy_for_playwright(None))

    def test_empty_string_returns_none(self):
        self.assertIsNone(browser.parse_proxy_for_playwright(""))

    def test_garbage_input_does_not_raise(self):
        # urlparse() is lenient and doesn't raise on malformed input, so
        # this just documents that the function never blows up on it —
        # not that it returns None (it returns a best-effort dict instead,
        # same as the literal spec this was built from).
        browser.parse_proxy_for_playwright("not a url :: at all")


class DismissCookieBannerTests(unittest.TestCase):
    def test_clicks_first_visible_matching_selector(self):
        page = MagicMock()
        visible_locator = MagicMock()
        visible_locator.count.return_value = 1
        visible_locator.is_visible.return_value = True
        page.locator.return_value.first = visible_locator

        result = browser.dismiss_cookie_banner(page)

        self.assertTrue(result)
        visible_locator.click.assert_called_once()

    def test_returns_false_when_no_banner_present(self):
        page = MagicMock()
        absent_locator = MagicMock()
        absent_locator.count.return_value = 0
        page.locator.return_value.first = absent_locator

        result = browser.dismiss_cookie_banner(page)

        self.assertFalse(result)

    def test_exception_on_one_selector_does_not_abort_the_rest(self):
        page = MagicMock()
        page.locator.side_effect = Exception("boom")

        result = browser.dismiss_cookie_banner(page)

        self.assertFalse(result)


class HumanScrollTests(unittest.TestCase):
    def test_calls_evaluate_the_requested_number_of_times(self):
        page = MagicMock()
        with patch.object(browser.time, "sleep"):
            browser.human_scroll(page, times=3)
        self.assertEqual(page.evaluate.call_count, 3)

    def test_stops_early_if_evaluate_raises(self):
        page = MagicMock()
        page.evaluate.side_effect = Exception("page closed")
        with patch.object(browser.time, "sleep"):
            browser.human_scroll(page, times=5)  # must not raise


class SafeGotoTests(unittest.TestCase):
    def test_returns_true_on_success(self):
        page = MagicMock()
        self.assertTrue(browser.safe_goto(page, "https://example.com"))
        page.goto.assert_called_once()

    def test_returns_false_on_exception(self):
        page = MagicMock()
        page.goto.side_effect = Exception("timeout")
        self.assertFalse(browser.safe_goto(page, "https://example.com"))


# ---------------------------------------------------------------------------
# discover.py: browser-mode retry + browser-only sources
# ---------------------------------------------------------------------------

class DiscoverBrowserModeTests(unittest.TestCase):
    def setUp(self):
        self._tmpdir = tempfile.TemporaryDirectory()
        self.db_path = str(Path(self._tmpdir.name) / "test.db")
        init_db(self.db_path)

    def tearDown(self):
        self._tmpdir.cleanup()

    @staticmethod
    def _biz(name: str, source: str) -> dict:
        return {
            "business_name": name, "website_url": None, "phone": None,
            "address": None, "category": "Plumbing", "source": source,
        }

    def test_browser_retry_not_triggered_when_primary_succeeds(self):
        def primary(niche, loc, budget):
            yield self._biz("Joe Plumbing", "yellowpages")

        browser_mock = MagicMock(return_value=iter([]))

        with patch.object(discover, "_SOURCE_FACTORIES", {"yellowpages": primary}), \
             patch.object(discover, "_REGION_SOURCES", {"us": ["yellowpages"]}), \
             patch.object(discover, "_BROWSER_RETRY_FACTORIES", {"yellowpages": browser_mock}), \
             patch.object(discover, "_BROWSER_ONLY_SOURCES", {}):
            discover.run_discover(self.db_path, "plumbers", "Austin, TX", 10, browser_mode=True)

        browser_mock.assert_not_called()
        with get_conn(self.db_path) as conn:
            rows = conn.execute("SELECT business_name FROM businesses").fetchall()
        self.assertEqual(len(rows), 1)

    def test_browser_retry_triggered_when_primary_returns_zero(self):
        def primary(niche, loc, budget):
            return iter([])

        def browser_variant(niche, loc, budget, headless=True):
            yield self._biz("Joe Plumbing", "yellowpages")

        with patch.object(discover, "_SOURCE_FACTORIES", {"yellowpages": primary}), \
             patch.object(discover, "_REGION_SOURCES", {"us": ["yellowpages"]}), \
             patch.object(discover, "_BROWSER_RETRY_FACTORIES", {"yellowpages": browser_variant}), \
             patch.object(discover, "_BROWSER_ONLY_SOURCES", {}):
            discover.run_discover(self.db_path, "plumbers", "Austin, TX", 10, browser_mode=True)

        with get_conn(self.db_path) as conn:
            rows = conn.execute("SELECT business_name, source FROM businesses").fetchall()
        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0]["source"], "yellowpages")

    def test_browser_retry_not_attempted_without_browser_mode_flag(self):
        def primary(niche, loc, budget):
            return iter([])

        browser_mock = MagicMock(return_value=iter([]))

        with patch.object(discover, "_SOURCE_FACTORIES", {"yellowpages": primary}), \
             patch.object(discover, "_REGION_SOURCES", {"us": ["yellowpages"]}), \
             patch.object(discover, "_BROWSER_RETRY_FACTORIES", {"yellowpages": browser_mock}), \
             patch.object(discover, "_BROWSER_ONLY_SOURCES", {}):
            discover.run_discover(self.db_path, "plumbers", "Austin, TX", 10, browser_mode=False)

        browser_mock.assert_not_called()

    def test_factory_raising_synchronously_is_still_caught(self):
        """Regression test: the generator call itself must happen inside
        the try block, not at the call site, or a factory that raises
        before yielding anything (rather than during iteration) crashes
        run_discover instead of being logged and skipped."""
        def failing_source(niche, loc, budget):
            raise RuntimeError("403 Forbidden")

        with patch.object(discover, "_SOURCE_FACTORIES", {"yellowpages": failing_source}), \
             patch.object(discover, "_REGION_SOURCES", {"us": ["yellowpages"]}), \
             patch.object(discover, "_BROWSER_RETRY_FACTORIES", {}), \
             patch.object(discover, "_BROWSER_ONLY_SOURCES", {}):
            discover.run_discover(self.db_path, "plumbers", "Austin, TX", 10)  # must not raise

    def test_browser_only_sources_added_for_us_region_when_browser_mode(self):
        def primary(niche, loc, budget):
            return iter([])

        def bbb_variant(niche, loc, budget, headless=True):
            yield self._biz("Joe Plumbing", "bbb")

        with patch.object(discover, "_SOURCE_FACTORIES", {"yellowpages": primary}), \
             patch.object(discover, "_REGION_SOURCES", {"us": ["yellowpages"]}), \
             patch.object(discover, "_BROWSER_RETRY_FACTORIES", {}), \
             patch.object(discover, "_BROWSER_ONLY_SOURCES", {"bbb": bbb_variant}):
            discover.run_discover(self.db_path, "plumbers", "Austin, TX", 10, browser_mode=True)

        with get_conn(self.db_path) as conn:
            rows = conn.execute("SELECT source FROM businesses").fetchall()
        self.assertEqual([r["source"] for r in rows], ["bbb"])

    def test_browser_only_sources_skipped_without_browser_mode(self):
        def primary(niche, loc, budget):
            return iter([])

        bbb_mock = MagicMock(return_value=iter([]))

        with patch.object(discover, "_SOURCE_FACTORIES", {"yellowpages": primary}), \
             patch.object(discover, "_REGION_SOURCES", {"us": ["yellowpages"]}), \
             patch.object(discover, "_BROWSER_RETRY_FACTORIES", {}), \
             patch.object(discover, "_BROWSER_ONLY_SOURCES", {"bbb": bbb_mock}):
            discover.run_discover(self.db_path, "plumbers", "Austin, TX", 10, browser_mode=False)

        bbb_mock.assert_not_called()


# ---------------------------------------------------------------------------
# crawl.py: Playwright fallback now also returns + merges social links
# ---------------------------------------------------------------------------

class CrawlPlaywrightSocialMergeTests(unittest.TestCase):
    def test_playwright_social_merged_when_static_found_none(self):
        """Regression test: _crawl_site_playwright now returns (emails,
        social) — run_crawl must merge the Playwright-discovered social
        links into whatever the static pass found, not discard them."""
        import tempfile as _tempfile
        tmpdir = _tempfile.TemporaryDirectory()
        try:
            db_path = str(Path(tmpdir.name) / "test.db")
            init_db(db_path)
            with get_conn(db_path) as conn:
                biz_id = upsert_business(
                    conn, niche="plumbers", location="Austin, TX",
                    business_name="Joe Plumbing", website_url="https://joeplumbing.com",
                    phone=None, address=None, category="Plumbing", source="yellowpages",
                )
                conn.execute(
                    "UPDATE businesses SET resolve_status='done', normalized_url=? WHERE id=?",
                    ("https://joeplumbing.com", biz_id),
                )

            with patch.object(crawl, "_crawl_site_static", return_value=([], [], {
                     "facebook": None, "instagram": None, "linkedin": None,
                 })), \
                 patch.object(crawl, "_crawl_site_playwright", return_value=([], {
                     "facebook": "https://facebook.com/joeplumbing", "instagram": None, "linkedin": None,
                 })), \
                 patch.object(crawl.time, "sleep"), \
                 patch.object(crawl.random, "uniform", return_value=0):
                crawl.run_crawl(db_path)

            with get_conn(db_path) as conn:
                row = conn.execute(
                    "SELECT facebook_url FROM businesses WHERE id=?", (biz_id,)
                ).fetchone()
            self.assertIsNotNone(row)
            self.assertEqual(row["facebook_url"], "https://facebook.com/joeplumbing")
        finally:
            tmpdir.cleanup()


# ---------------------------------------------------------------------------
# crawl.py / resolve.py: direct connection tried before spending a proxy
# ---------------------------------------------------------------------------

class CrawlSiteStaticDirectFirstTests(unittest.TestCase):
    def test_direct_success_never_opens_a_proxied_client(self):
        make_client_calls = []

        def fake_make_client(proxy):
            make_client_calls.append(proxy)
            cm = MagicMock()
            cm.__enter__.return_value = MagicMock()
            cm.__exit__.return_value = False
            return cm

        with patch.object(crawl, "_make_client", side_effect=fake_make_client), \
             patch.object(crawl, "_fetch_page", return_value="<html>hi</html>"), \
             patch.object(crawl, "_discover_contact_links", return_value=[]), \
             patch.object(crawl, "get_pool") as get_pool_mock:
            get_pool_mock.return_value.get.return_value = "http://user:pass@proxy.example.com:8080"
            emails, pages, social = crawl._crawl_site_static("https://example.com")

        self.assertEqual(pages[0], "https://example.com")
        # Only the direct (None) client was ever opened for the homepage —
        # proxy was never spent since direct already returned HTML.
        self.assertEqual(make_client_calls[0], None)
        self.assertNotIn("http://user:pass@proxy.example.com:8080", make_client_calls)

    def test_direct_failure_falls_back_to_proxy(self):
        """Regression test: if the direct homepage fetch returns nothing
        (connection error, timeout, or any other unresolved failure), the
        proxy from the pool must be tried next rather than giving up."""
        attempts = []

        def fake_make_client(proxy):
            attempts.append(proxy)
            cm = MagicMock()
            cm.__enter__.return_value = MagicMock()
            cm.__exit__.return_value = False
            return cm

        def fake_fetch_page(client, url, proxy=None):
            # Direct (proxy=None) fails; the proxied attempt succeeds.
            return None if proxy is None else "<html>hi</html>"

        with patch.object(crawl, "_make_client", side_effect=fake_make_client), \
             patch.object(crawl, "_fetch_page", side_effect=fake_fetch_page), \
             patch.object(crawl, "_discover_contact_links", return_value=[]), \
             patch.object(crawl, "get_pool") as get_pool_mock:
            get_pool_mock.return_value.get.return_value = "http://user:pass@proxy.example.com:8080"
            emails, pages, social = crawl._crawl_site_static("https://example.com")

        self.assertEqual(pages[0], "https://example.com")
        self.assertIn(None, attempts)  # direct was tried
        self.assertIn("http://user:pass@proxy.example.com:8080", attempts)  # then proxy

    def test_both_direct_and_proxy_fail_returns_empty(self):
        with patch.object(crawl, "_make_client", return_value=MagicMock(
                 __enter__=MagicMock(return_value=MagicMock()), __exit__=MagicMock(return_value=False))), \
             patch.object(crawl, "_fetch_page", return_value=None), \
             patch.object(crawl, "get_pool") as get_pool_mock:
            get_pool_mock.return_value.get.return_value = "http://user:pass@proxy.example.com:8080"
            emails, pages, social = crawl._crawl_site_static("https://example.com")

        self.assertEqual((emails, pages), ([], []))

    def test_no_proxies_configured_does_not_retry_direct_twice_in_probe(self):
        """When the pool has no proxy to fall back to (get() -> None), the
        homepage probe's [None, proxy] candidate list collapses via
        dict.fromkeys to a single [None] — it must not pointlessly retry
        the identical direct attempt a second time under the guise of
        'now trying proxy'."""
        probe_calls = []

        def fake_make_client(proxy):
            probe_calls.append(proxy)
            cm = MagicMock()
            cm.__enter__.return_value = MagicMock()
            cm.__exit__.return_value = False
            return cm

        with patch.object(crawl, "_make_client", side_effect=fake_make_client), \
             patch.object(crawl, "_fetch_page", return_value=None), \
             patch.object(crawl, "get_pool") as get_pool_mock:
            get_pool_mock.return_value.get.return_value = None  # no proxies in pool
            emails, pages, social = crawl._crawl_site_static("https://example.com")

        self.assertEqual(probe_calls, [None])
        self.assertEqual((emails, pages), ([], []))


class ResolveOneDirectFirstTests(unittest.TestCase):
    class _FakeResp:
        status_code = 200
        url = "https://example.com/"

    def test_direct_success_never_tries_proxy(self):
        client_calls = []

        def fake_client(**kwargs):
            client_calls.append(kwargs.get("proxies"))
            cm = MagicMock()
            cm.__enter__.return_value.head.return_value = self._FakeResp()
            cm.__exit__.return_value = False
            return cm

        with patch.object(resolve.httpx, "Client", side_effect=fake_client), \
             patch.object(resolve, "get_pool") as get_pool_mock:
            get_pool_mock.return_value.get.return_value = "http://user:pass@proxy.example.com:8080"
            get_pool_mock.return_value.get_auth_header.return_value = {}
            result = resolve._resolve_one("https://example.com/")

        self.assertEqual(result, "https://example.com/")
        self.assertEqual(client_calls, [None])  # proxy never even constructed

    def test_connection_error_on_direct_falls_back_to_proxy(self):
        """Regression test: a direct ConnectError/timeout must trigger a
        retry through the pool's proxy, not an immediate give-up."""
        client_calls = []

        def fake_client(**kwargs):
            proxies = kwargs.get("proxies")
            client_calls.append(proxies)
            cm = MagicMock()
            if proxies is None:
                cm.__enter__.return_value.head.side_effect = httpx.ConnectError("boom")
            else:
                cm.__enter__.return_value.head.return_value = self._FakeResp()
            cm.__exit__.return_value = False
            return cm

        with patch.object(resolve.httpx, "Client", side_effect=fake_client), \
             patch.object(resolve, "get_pool") as get_pool_mock:
            pool = get_pool_mock.return_value
            pool.get.return_value = "http://user:pass@proxy.example.com:8080"
            pool.get_auth_header.return_value = {}
            pool.backoff_sleep = MagicMock()
            result = resolve._resolve_one("https://example.com/")

        self.assertEqual(result, "https://example.com/")
        self.assertIn(None, client_calls)
        proxy_dicts = [c for c in client_calls if c is not None]
        self.assertTrue(proxy_dicts)
        self.assertTrue(
            any(d.get("http://") == "http://user:pass@proxy.example.com:8080" for d in proxy_dicts)
        )

    def test_both_direct_and_proxy_fail_returns_none(self):
        def fake_client(**kwargs):
            cm = MagicMock()
            cm.__enter__.side_effect = httpx.ConnectError("boom")
            cm.__exit__.return_value = False
            return cm

        with patch.object(resolve.httpx, "Client", side_effect=fake_client), \
             patch.object(resolve, "get_pool") as get_pool_mock:
            pool = get_pool_mock.return_value
            pool.get.return_value = "http://user:pass@proxy.example.com:8080"
            pool.get_auth_header.return_value = {}
            pool.backoff_sleep = MagicMock()
            result = resolve._resolve_one("https://example.com/")

        self.assertIsNone(result)


if __name__ == "__main__":
    unittest.main()
