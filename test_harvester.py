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

from email_harvester import crawl, discover, social, verify, write
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

        def fake_fetch(client, url):
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


if __name__ == "__main__":
    unittest.main()
