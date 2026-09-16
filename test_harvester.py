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

No real network calls are made: httpx fetching and DNS resolution are
monkeypatched out.
"""

import json
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch, MagicMock

import dns.exception

from email_harvester import discover, verify, write
from email_harvester.db import init_db, get_conn, upsert_business, upsert_email
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

    def test_falls_back_to_timestamped_path_when_output_is_locked(self):
        """Regression test for the fix: a locked output file (e.g. open in
        Excel — the file that's open when the previous run's results are
        being reviewed) must not crash the whole pipeline and discard a
        completed run's data; it must save under an alternate name."""
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


if __name__ == "__main__":
    unittest.main()
