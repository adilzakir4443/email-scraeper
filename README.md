# email-harvester

A private, server-side Python 3.11 CLI tool that scrapes US business directories
for a given niche + location, crawls the resulting business websites for emails,
filters them by syntax and MX record, and writes a leads spreadsheet.

**This is a private internal tool — no web UI, no accounts, no SaaS.**

---

## Project Layout

```
email-harvester/
├── email_harvester/
│   ├── __init__.py
│   ├── cli.py        # Click entrypoint; pipeline orchestration
│   ├── db.py         # SQLite schema and query helpers
│   ├── proxy.py      # Proxy rotation from PROXY_POOL env var
│   ├── discover.py   # Stage 1: scrape Yellow Pages / Bing / Yelp
│   ├── resolve.py    # Stage 2: normalise + resolve website URLs
│   ├── crawl.py      # Stage 3: crawl sites, extract emails (Playwright fallback)
│   ├── extract.py    # Email extraction: mailto, regex, obfuscation, CF cfemail
│   ├── verify.py     # Stage 4a: syntax → disposable → role → MX verification
│   └── write.py      # Stage 4b: dedupe, suppress, write Excel
├── suppression.csv   # Global suppression list (type,value)
├── .env.example      # Template for PROXY_POOL etc.
├── requirements.txt
├── setup.py
└── README.md
```

---

## SQLite Schema

```sql
-- All pipeline state lives here. Re-running continues from the last
-- completed point rather than starting over.

CREATE TABLE businesses (
    id               INTEGER PRIMARY KEY AUTOINCREMENT,
    niche            TEXT    NOT NULL,
    location         TEXT    NOT NULL,
    business_name    TEXT    NOT NULL,
    website_url      TEXT,                -- raw URL from directory listing
    normalized_url   TEXT,                -- cleaned URL after RESOLVE stage
    phone            TEXT,
    address          TEXT,
    category         TEXT,
    source           TEXT    NOT NULL,    -- yellowpages | bing | yelp
    discovered_at    TEXT    DEFAULT (datetime('now')),
    resolve_status   TEXT    DEFAULT 'pending',  -- pending|done|no_site|failed
    crawl_status     TEXT    DEFAULT 'pending',  -- pending|done|failed|no_emails_static
    playwright_tried INTEGER DEFAULT 0,
    UNIQUE(niche, location, business_name, source)
);

CREATE TABLE emails (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    business_id     INTEGER NOT NULL REFERENCES businesses(id) ON DELETE CASCADE,
    email           TEXT    NOT NULL,
    source_url      TEXT,
    extract_method  TEXT,    -- mailto | regex | obfuscated | cfemail
    syntax_valid    INTEGER,
    is_disposable   INTEGER,
    is_role         INTEGER,
    is_freemail     INTEGER,
    mx_status       TEXT,    -- acceptable | risky | invalid | error
    tier            TEXT,    -- acceptable | risky | invalid
    verified_at     TEXT,
    UNIQUE(business_id, email)
);
```

---

## Installation

### 1. Python 3.11 + virtualenv

```bash
python3.11 -m venv .venv
source .venv/bin/activate
pip install -e .
```

### 2. Playwright Chromium (for JS-heavy site fallback)

```bash
playwright install chromium
# On a headless Linux VPS you may also need:
playwright install-deps chromium
```

### 3. Proxy pool (optional but recommended)

```bash
cp .env.example .env
# Edit .env and set PROXY_POOL=http://user:pass@host:port,...
```

---

## Usage

### Full pipeline (most common)

```bash
email-harvester --niche "plumbers" --location "Austin, TX" --max 200
```

This runs all four stages in sequence and writes:
`leads_plumbers_austin_tx_20240601.xlsx`

### All options

```
Options:
  --niche TEXT        Business type to search, e.g. "plumbers"  [required]
  --location TEXT     City + state, e.g. "Austin, TX"           [required]
  --max INTEGER       Maximum listings to discover (default: 100)
  --db TEXT           SQLite database path (default: harvester.db)
  --out TEXT          Output Excel path (auto-named if omitted)
  --suppress TEXT     CSV suppression list (columns: type, value)
  --proxy-pool TEXT   Comma-separated proxy URLs (overrides PROXY_POOL env var)
  --stage [all|discover|resolve|crawl|verify|write]
                      Run a single stage instead of the full pipeline
  -v, --verbose       Enable DEBUG logging
  -h, --help          Show this message and exit
```

### Running individual stages

Each stage is idempotent — already-completed rows are skipped:

```bash
# Re-run just the discovery after changing --max
email-harvester --niche "electricians" --location "Dallas, TX" --stage discover --max 500

# Re-run just the write stage (e.g. after updating suppression list)
email-harvester --niche "electricians" --location "Dallas, TX" --stage write --suppress suppression.csv
```

### Resume after interruption

Simply re-run the exact same command. SQLite tracks which rows have been
processed; the tool picks up where it left off.

---

## Pipeline Stages

### Stage 1 — DISCOVER

Scrapes three directories in order:

| Priority | Source       | Notes |
|----------|-------------|-------|
| Primary  | Yellow Pages | Full pagination; best coverage |
| Secondary| Bing Local   | Local pack results |
| Tertiary | Yelp         | Supplementary; links to Yelp listing pages |

All businesses (even those without a website) are inserted into the
`businesses` table. The total budget (`--max`) is split across sources.

Funnel log:
```
[DISCOVER] DONE — total=183  with_site=141  site_less=42
```

### Stage 2 — RESOLVE

For each business with a `website_url`:
- Strips tracking params (utm_*, fbclid, gclid, etc.)
- Forces https://
- Follows up to 5 redirects to find the canonical URL
- Marks `resolve_status` as `done` / `failed` / `no_site`

### Stage 3 — CRAWL + EXTRACT

For each resolved site (low concurrency, 3–8s delay between sites):

1. Fetches the homepage
2. Discovers same-domain contact/about/team links from anchor text
3. Fetches `/contact`, `/about`, `/team`, `/contact-us`, `/about-us` + discovered links
4. Extracts emails via four methods:
   - `mailto:` hrefs
   - Regex over raw HTML
   - `[at]` / `(at)` / ` at ` obfuscation de-munging
   - Cloudflare `data-cfemail` hex decoding
5. If **zero emails** found statically → retries that site **once** with
   Playwright Chromium (JavaScript rendering)

Funnel log:
```
[CRAWL] DONE — crawled=137  failed=4  emails_raw=312  playwright_fallbacks=23
```

### Stage 4a — VERIFY

For each extracted email (in order):

| Check | Fail action |
|-------|------------|
| Syntax (email-validator) | tier = invalid |
| Disposable domain blocklist | tier = invalid |
| Role account (info@, sales@, admin@, …) | tier = risky (not dropped) |
| Free-mail provider (gmail, yahoo, outlook, …) | tier = risky (not dropped) |
| MX record lookup (dnspython) | tier = invalid if no MX |

**No SMTP probing. No catch-all detection. No name→email guessing.**

### Stage 4b — WRITE

1. Loads suppression CSV (emails + domains to exclude)
2. Deduplicates: first by exact email, then one email per domain (best tier wins)
3. Writes Excel with columns:

   `Company Name | Owner Name | Phone | Category | Email | Website | Address | Comment`

   - **Owner Name** is always blank (not guessed)
   - **Comment** example: `mx_status=acceptable; role=false; freemail=false; source=yellowpages; extract=mailto; tier=acceptable`
   - Risky rows are highlighted in light yellow
   - Invalid emails are **never written** to the output file

---

## Suppression List

`suppression.csv` format:

```csv
type,value
email,noreply@somedomain.com
domain,spamtrap.io
domain,example.com
```

Pass it with `--suppress suppression.csv`. Any email or domain listed here is
removed from the Excel output silently.

---

## Proxy Configuration

Set `PROXY_POOL` as a comma-separated list of proxy endpoints:

```bash
export PROXY_POOL="http://user:pass@proxy1.host:8080,http://user:pass@proxy2.host:8080"
```

Or put it in `.env` (the tool loads `.env` automatically).

The tool selects a random proxy per request. On 403/429 responses it applies
exponential backoff (2s → 4s → 8s → 16s max 120s) before retrying.

---

## Cron / Unattended Scheduling

Add to crontab (`crontab -e`) to run weekly:

```cron
# Every Monday at 02:00 — plumbers in Austin, TX
0 2 * * 1 /path/to/.venv/bin/email-harvester \
    --niche "plumbers" \
    --location "Austin, TX" \
    --max 300 \
    --db /var/data/harvester.db \
    --suppress /var/data/suppression.csv \
    >> /var/log/email-harvester.log 2>&1
```

Multiple niches / locations can be batched in a single cron script:

```bash
#!/bin/bash
# harvest.sh
HARVESTER=/path/to/.venv/bin/email-harvester
DB=/var/data/harvester.db
SUP=/var/data/suppression.csv

declare -A jobs=(
    ["plumbers"]="Austin, TX"
    ["electricians"]="Dallas, TX"
    ["hvac contractors"]="Houston, TX"
)

for niche in "${!jobs[@]}"; do
    location="${jobs[$niche]}"
    $HARVESTER --niche "$niche" --location "$location" \
               --max 200 --db "$DB" --suppress "$SUP"
done
```

---

## Logging

Every stage logs funnel counts at INFO level. Use `-v` for DEBUG.

Redirect to a file:
```bash
email-harvester --niche "roofers" --location "Phoenix, AZ" \
    >> harvest.log 2>&1
```

---

## What This Tool Does NOT Do

- Google Maps scraping
- SMTP probing or catch-all detection
- OCR on images
- Name → email pattern guessing (e.g. `john.doe@company.com`)
- Web UI or any server-side component

---

## Dependencies

| Package | Purpose |
|---------|---------|
| `httpx` | HTTP client with proxy + redirect support |
| `playwright` | Headless Chromium fallback for JS-heavy sites |
| `beautifulsoup4` + `lxml` | HTML parsing |
| `dnspython` | MX record lookups |
| `email-validator` | RFC syntax validation |
| `openpyxl` | Excel output |
| `click` | CLI framework |
| `tenacity` | Retry logic |
| `disposable-email-domains` | Disposable domain blocklist |
| `tldextract` | Reliable domain parsing |
| `python-dotenv` | `.env` file support |
