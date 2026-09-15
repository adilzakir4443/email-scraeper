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

No real network calls are made: httpx fetching is monkeypatched out.
"""

import json
import unittest
from unittest.mock import patch, MagicMock

from email_harvester import discover
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


if __name__ == "__main__":
    unittest.main()
