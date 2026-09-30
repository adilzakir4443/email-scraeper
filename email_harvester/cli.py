"""
email-harvester CLI entrypoint.

Usage:
    email-harvester --niche "plumbers" --location "Austin, TX" --max 200

All four stages are run sequentially; each is idempotent and resumes from
the last completed point in SQLite.
"""

import logging
import os
import sys
from typing import Optional

import click
from dotenv import load_dotenv

load_dotenv()  # pick up .env if present (PROXY_POOL, etc.)


def _configure_logging(verbose: bool) -> None:
    level = logging.DEBUG if verbose else logging.INFO
    logging.basicConfig(
        level=level,
        format="%(asctime)s  %(levelname)-8s  %(name)s  %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
        handlers=[logging.StreamHandler(sys.stdout)],
    )
    # Quieten noisy third-party loggers
    for noisy in ("httpx", "httpcore", "urllib3", "playwright"):
        logging.getLogger(noisy).setLevel(logging.WARNING)


@click.command(context_settings={"help_option_names": ["-h", "--help"]})
@click.option("--niche",    required=True,  help='Business type to search, e.g. "plumbers"')
@click.option("--location", required=True,  help='City + state, e.g. "Austin, TX"')
@click.option("--max",      "max_results",  default=100, show_default=True,
              help="Maximum listings to discover across all sources")
@click.option("--db",       "db_path",      default="harvester.db", show_default=True,
              help="Path to SQLite database file (created if absent)")
@click.option("--out",      "out_path",     default=None,
              help="Output Excel file path (auto-named if omitted)")
@click.option("--suppress", "suppress_path", default=None,
              help="CSV suppression list (columns: type, value)")
@click.option("--proxy-pool", "proxy_pool", default=None,
              help="Comma-separated proxy URLs (overrides PROXY_POOL env var)")
@click.option("--enrich",   "enrich_path",  default=None,
              help="Path to existing Excel file to enrich with missing data")
@click.option("--api-fallback", "api_fallback", is_flag=True, default=False,
              help="Use Google API as fallback when proxies fail (needs GOOGLE_API_KEY/"
                   "GOOGLE_PLACES_KEY/GOOGLE_CX in .env — costs real API quota)")
@click.option("--stage",    "stage",        default="all",
              type=click.Choice(["all", "discover", "resolve", "crawl", "social", "verify", "write"],
                                case_sensitive=False),
              show_default=True,
              help="Run a single stage instead of the full pipeline")
@click.option("--verbose",  "-v",           is_flag=True, default=False,
              help="Enable DEBUG logging")
def main(
    niche: str,
    location: str,
    max_results: int,
    db_path: str,
    out_path: Optional[str],
    suppress_path: Optional[str],
    proxy_pool: Optional[str],
    enrich_path: Optional[str],
    api_fallback: bool,
    stage: str,
    verbose: bool,
) -> None:
    """
    email-harvester — scrape US business directories and extract contact emails.

    Runs resumable stages: DISCOVER → RESOLVE → CRAWL → SOCIAL → VERIFY+WRITE.
    All state is persisted in SQLite so the tool can be safely interrupted and
    restarted without losing progress.

    Set PROXY_POOL=http://user:pass@host:port,... in the environment (or .env)
    to enable proxy rotation.  Without proxies the tool will run unproxied and
    may be rate-limited by target sites.

    --enrich <file.xlsx> runs a completely separate mode: it fills in missing
    columns on an existing spreadsheet instead of running the normal
    DISCOVER..WRITE pipeline.

    Enrich an existing Excel sheet:
        email-harvester --enrich my_leads.xlsx --niche "plumbers" --location "Austin, TX"
    With Google API fallback:
        email-harvester --enrich my_leads.xlsx --niche "plumbers" --location "Austin, TX" --api-fallback

    --api-fallback only ever activates when the proxy-based path has already
    come back empty — never as a routine supplement to it, so API quota is
    never spent while proxies are working. Example:
        email-harvester --niche "plumbers" --location "Austin, TX" --api-fallback
    """
    _configure_logging(verbose)
    logger = logging.getLogger(__name__)

    # Initialise proxy pool
    from . import proxy as proxy_mod
    if proxy_pool:
        proxy_mod.init_pool([p.strip() for p in proxy_pool.split(",") if p.strip()])
    else:
        proxy_mod.init_pool()  # reads PROXY_POOL env var

    # Initialise DB schema
    from .db import init_db
    init_db(db_path)

    # Google API fallback keys — read from environment (.env), never hardcoded.
    google_api_key = os.getenv("GOOGLE_API_KEY")
    google_places_key = os.getenv("GOOGLE_PLACES_KEY") or google_api_key
    google_cx = os.getenv("GOOGLE_CX")

    if api_fallback and not google_api_key:
        logger.warning(
            "--api-fallback set but GOOGLE_API_KEY not found in environment — fallback disabled"
        )
        api_fallback = False

    if enrich_path:
        from .enrich import run_enrich
        try:
            output = run_enrich(
                input_path=enrich_path,
                niche=niche,
                location=location,
                api_fallback=api_fallback,
                google_api_key=google_api_key,
                google_cx=google_cx,
                google_places_key=google_places_key,
            )
        except ValueError as exc:
            # e.g. the input file has no "Company Name" column — a clear,
            # one-line message beats a raw traceback for this kind of
            # ordinary user-input mistake.
            click.echo(f"\nError: {exc}", err=True)
            raise SystemExit(1)
        click.echo(f"\nEnriched file saved to: {output}")
        return  # skip normal pipeline when enriching

    logger.info(
        "email-harvester  niche=%r  location=%r  max=%d  db=%s  stage=%s",
        niche, location, max_results, db_path, stage,
    )

    # ---------------------------------------------------------------------------
    # Run stages
    # ---------------------------------------------------------------------------

    if stage in ("all", "discover"):
        from .discover import run_discover
        run_discover(
            db_path, niche, location, max_results,
            api_fallback=api_fallback, google_places_key=google_places_key,
        )

    if stage in ("all", "resolve"):
        from .resolve import run_resolve
        run_resolve(
            db_path,
            api_fallback=api_fallback, google_api_key=google_api_key, google_cx=google_cx,
        )

    if stage in ("all", "crawl"):
        from .crawl import run_crawl
        run_crawl(db_path)

    if stage in ("all", "social"):
        from .social import run_social
        run_social(
            db_path,
            api_fallback=api_fallback, google_api_key=google_api_key, google_cx=google_cx,
        )

    if stage in ("all", "verify"):
        from .verify import run_verify
        run_verify(db_path)

    if stage in ("all", "write"):
        from .write import run_write
        output_file = run_write(
            db_path=db_path,
            niche=niche,
            location=location,
            out_path=out_path,
            suppress_path=suppress_path,
        )
        click.echo(f"\nOutput written to: {output_file}")

    if stage == "all":
        from .db import get_conn
        with get_conn(db_path) as conn:
            total_biz = conn.execute(
                "SELECT COUNT(*) FROM businesses WHERE niche=? AND location=?",
                (niche, location),
            ).fetchone()[0]
            acceptable = conn.execute(
                """SELECT COUNT(*) FROM emails e
                   JOIN businesses b ON b.id=e.business_id
                   WHERE b.niche=? AND b.location=? AND e.tier='acceptable'""",
                (niche, location),
            ).fetchone()[0]
            risky = conn.execute(
                """SELECT COUNT(*) FROM emails e
                   JOIN businesses b ON b.id=e.business_id
                   WHERE b.niche=? AND b.location=? AND e.tier='risky'""",
                (niche, location),
            ).fetchone()[0]

        click.echo(
            f"\n=== FUNNEL SUMMARY ===\n"
            f"  Discovered businesses : {total_biz}\n"
            f"  Acceptable emails     : {acceptable}\n"
            f"  Risky emails          : {risky}\n"
        )


if __name__ == "__main__":
    main()
