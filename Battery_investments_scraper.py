"""
Battery Investment Scraper
==========================
Automatically searches and extracts investment data for battery value chain projects.
Generates a structured Excel file matching the Investments_GF table format.

Usage:
    # Read project list from existing Excel (recommended)
    python battery_investment_scraper.py --from-excel Gigafactories.xlsx

    # Pass a custom project list
    python battery_investment_scraper.py --projects "Northvolt Skelleftea,ACC Douvrin,Verkor Dunkirk" --segment "gigafactory battery cell Europe"

    # Segment-only search (no project list, for a brand new topic)
    python battery_investment_scraper.py --segment "battery active materials cathode anode Europe"

    # Limit pages per project (faster / cheaper)
    python battery_investment_scraper.py --from-excel Gigafactories.xlsx --max-pages 3

Prerequisites:
    pip install anthropic playwright pandas openpyxl
    python -m playwright install chromium
    Environment variable ANTHROPIC_API_KEY must be set.
"""

import os
import re
import json
import asyncio
import logging
import argparse
from datetime import datetime
from pathlib import Path

import anthropic
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from playwright.async_api import async_playwright

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  %(message)s",
    datefmt="%H:%M:%S",
)
log = logging.getLogger(__name__)

CLAUDE_MODEL = "claude-sonnet-4-20250514"
MAX_TOKENS = 2000
PAGE_DELAY_S = 2          # seconds between page loads (avoids blocks)
MAX_PAGES_PER_PROJECT = 4 # max pages to read per project

# Domains ranked by relevance — these appear first in search results sorting
PRIORITY_DOMAINS = [
    "electrive.com",
    "electrive.net",
    "cinea.ec.europa.eu",
    "eib.org",
    "ipcei-batteries.eu",
    "ec.europa.eu",
    "benchmarkminerals.com",
    "spglobal.com",
    "reuters.com",
    "ft.com",
    "bloomberg.com",
]

# Output columns — identical to Investments_GF plus three traceability columns
OUTPUT_COLUMNS = [
    "Project",
    "Phase",
    "Unique ID",
    "Funding Amount (Original)",
    "Funding Amount (Billion €)",
    "Funding Source (Original)",
    "Funding Source (Standardized)",
    "Funding Source origin country",
    "Type of Support (Original)",
    "Type of Support (Standardized)",
    "Year / Timeline",
    "References",
    "Destination",
    "Origin region",
    # Traceability columns
    "Search Query Used",
    "Confidence",
    "Scraped At",
]

FUNDING_SOURCE_CATEGORIES = [
    "National/State Government",
    "EU Innovation Fund",
    "European Investment Bank (EIB)",
    "Commercial Banks",
    "Private Equity / VC",
    "Vehicle Manufacturers (OEMs)",
    "Horizon Europe",
    "Multilateral",
    "Other",
]

SUPPORT_TYPE_CATEGORIES = ["Grant", "Loan", "Equity", "Guarantee", "Other"]
ORIGIN_REGION_CATEGORIES = ["Europe", "North America", "Asia", "Other"]


# ---------------------------------------------------------------------------
# Read project list from Excel
# ---------------------------------------------------------------------------

def read_projects_from_excel(excel_path: str, sheet_name: str = "Investments_GF") -> list[str]:
    """Extract unique project names from the first column of a given sheet."""
    df = pd.read_excel(excel_path, sheet_name=sheet_name)
    first_col = df.columns[0]
    projects = df[first_col].dropna().unique().tolist()
    projects = [str(p).strip() for p in projects if str(p).strip()]
    log.info(f"Loaded {len(projects)} unique projects from '{excel_path}' sheet '{sheet_name}'")
    return projects


# ---------------------------------------------------------------------------
# Search query generation
# ---------------------------------------------------------------------------

def build_queries_for_project(project: str, segment: str) -> list[str]:
    """
    Generate targeted search queries for a specific project.
    Multiple angles: amount, funder type, official sources.
    """
    return [
        f'"{project}" investment funding million billion',
        f'"{project}" grant loan equity financing battery',
        f'"{project}" European Investment Bank IPCEI Innovation Fund',
        f'"{project}" battery gigafactory funding announcement',
    ]


def build_queries_for_segment(segment: str) -> list[str]:
    """Fallback: broad queries when no project list is provided."""
    return [
        f"{segment} investment funding 2022 2023 2024 2025 million billion",
        f"{segment} grant loan equity EIB IPCEI Innovation Fund",
        f"{segment} gigafactory battery plant financing Europe",
        f"site:electrive.com {segment} investment funding",
        f"site:eib.org battery {segment}",
        f"site:cinea.ec.europa.eu battery innovation fund",
    ]


# ---------------------------------------------------------------------------
# Browser helpers
# ---------------------------------------------------------------------------

async def google_search(page, query: str, max_results: int = 8) -> list[dict]:
    """Run a Google search and return a list of {url, title} dicts."""
    search_url = (
        f"https://www.google.com/search"
        f"?q={query.replace(' ', '+')}&num={max_results}&hl=en&gl=US"
    )
    results = []
    try:
        await page.goto(search_url, wait_until="domcontentloaded", timeout=15000)
        await asyncio.sleep(PAGE_DELAY_S)
        links = await page.eval_on_selector_all(
            "a[href]",
            """els => els
                .map(e => ({href: e.href, text: e.innerText}))
                .filter(e =>
                    e.href.startsWith('http') &&
                    !e.href.includes('google.com') &&
                    !e.href.includes('youtube.com') &&
                    e.href.length > 25
                )
            """,
        )
        seen = set()
        for link in links:
            url = link["href"]
            if url not in seen:
                seen.add(url)
                results.append({"url": url, "title": link.get("text", "")[:120]})
            if len(results) >= max_results:
                break
    except Exception as e:
        log.warning(f"Google search error: {e}")
    return results


async def fetch_page_text(page, url: str, max_chars: int = 8000) -> str:
    """Load a page and return its plain text (truncated to max_chars)."""
    try:
        await page.goto(url, wait_until="domcontentloaded", timeout=20000)
        await asyncio.sleep(PAGE_DELAY_S)
        text = await page.evaluate("""() => {
            ['nav','footer','script','style','header','aside'].forEach(tag => {
                document.querySelectorAll(tag).forEach(el => el.remove());
            });
            return document.body ? document.body.innerText : '';
        }""")
        return text[:max_chars]
    except Exception as e:
        log.warning(f"Could not load {url}: {e}")
        return ""


def prioritize_urls(urls: list[dict]) -> list[dict]:
    """Sort URLs so priority domains appear first."""
    def rank(item):
        for i, domain in enumerate(PRIORITY_DOMAINS):
            if domain in item["url"]:
                return i
        return len(PRIORITY_DOMAINS)
    return sorted(urls, key=rank)


# ---------------------------------------------------------------------------
# LLM extraction
# ---------------------------------------------------------------------------

def extract_investments(
    client: anthropic.Anthropic,
    text: str,
    url: str,
    segment: str,
    project_hint: str,
) -> list[dict]:
    """
    Send page text to Claude and extract structured investment records.
    Returns a list of dicts, one per investment line found.
    """
    prompt = f"""You are extracting structured investment data from a web page about battery value chain financing.

Segment: {segment}
Project context: {project_hint}
Source URL: {url}

PAGE TEXT:
{text}

---

Extract ALL distinct investment/funding events mentioned that are relevant to the project "{project_hint}".
Return an empty list [] if nothing relevant is found.

For each investment, return a JSON object with these fields:
- project: project or company name (string)
- phase: phase number if mentioned, otherwise 1 (integer)
- funding_amount_original: amount as written in the text, e.g. "€500 million", "1.2 billion USD" (string, or "Info Not Available")
- funding_amount_bn_eur: amount converted to billion euros as a float, e.g. 0.5 (float, or null if unknown)
- funding_source_original: exact name of the funder as written (string)
- funding_source_standardized: one of: {", ".join(FUNDING_SOURCE_CATEGORIES)} (string)
- funding_source_country: country of origin of the funder (string)
- support_type_original: type of instrument as written, e.g. "IPCEI grant", "EIB loan", "equity round" (string)
- support_type_standardized: one of: {", ".join(SUPPORT_TYPE_CATEGORIES)} (string)
- year_timeline: year or period, e.g. "2023", "2024-2026" (string)
- destination_country: country where the investment goes (string)
- origin_region: one of: {", ".join(ORIGIN_REGION_CATEGORIES)} (string)
- confidence: "high" if amounts and funder are explicit, "medium" if inferred, "low" if vague (string)

Respond ONLY with a valid JSON array. No markdown, no explanation.
Example: [{{"project": "Northvolt", "phase": 1, "funding_amount_original": "€1 billion", ...}}]
"""
    try:
        response = client.messages.create(
            model=CLAUDE_MODEL,
            max_tokens=MAX_TOKENS,
            messages=[{"role": "user", "content": prompt}],
        )
        raw = response.content[0].text.strip()
        raw = re.sub(r"^```json\s*", "", raw)
        raw = re.sub(r"\s*```$", "", raw)
        data = json.loads(raw)
        return data if isinstance(data, list) else []
    except json.JSONDecodeError as e:
        log.warning(f"Invalid JSON from Claude: {e}")
        return []
    except Exception as e:
        log.warning(f"Claude API error: {e}")
        return []


# ---------------------------------------------------------------------------
# Row building and deduplication
# ---------------------------------------------------------------------------

def investment_to_row(inv: dict, url: str, query: str) -> dict:
    project = inv.get("project", "Unknown")
    phase = inv.get("phase", 1)
    return {
        "Project": project,
        "Phase": phase,
        "Unique ID": f"{project}, Phase {phase}",
        "Funding Amount (Original)": inv.get("funding_amount_original", "Info Not Available"),
        "Funding Amount (Billion €)": inv.get("funding_amount_bn_eur", "-"),
        "Funding Source (Original)": inv.get("funding_source_original", "-"),
        "Funding Source (Standardized)": inv.get("funding_source_standardized", "-"),
        "Funding Source origin country": inv.get("funding_source_country", "-"),
        "Type of Support (Original)": inv.get("support_type_original", "-"),
        "Type of Support (Standardized)": inv.get("support_type_standardized", "-"),
        "Year / Timeline": inv.get("year_timeline", "-"),
        "References": url,
        "Destination": inv.get("destination_country", "-"),
        "Origin region": inv.get("origin_region", "Europe"),
        "Search Query Used": query,
        "Confidence": inv.get("confidence", "low"),
        "Scraped At": datetime.now().strftime("%Y-%m-%d %H:%M"),
    }


def deduplicate(rows: list[dict]) -> list[dict]:
    """
    Remove obvious duplicates: same project + amount + funder.
    Keep the row with the highest confidence.
    """
    rank = {"high": 3, "medium": 2, "low": 1}
    seen = {}
    for row in rows:
        key = (
            str(row.get("Project", "")).lower().strip(),
            str(row.get("Funding Amount (Original)", "")).lower().strip(),
            str(row.get("Funding Source (Original)", "")).lower().strip(),
        )
        current_rank = rank.get(row.get("Confidence", "low"), 1)
        if key not in seen or current_rank > rank.get(seen[key].get("Confidence", "low"), 1):
            seen[key] = row
    return list(seen.values())


# ---------------------------------------------------------------------------
# Excel export
# ---------------------------------------------------------------------------

def export_to_excel(rows: list[dict], output_path: str, segment: str):
    """Write results to a formatted Excel file matching the Investments_GF style."""
    df = pd.DataFrame(rows, columns=OUTPUT_COLUMNS)

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Investments", index=False)
        wb = writer.book
        ws = writer.sheets["Investments"]

        # Header styling
        header_fill = PatternFill("solid", fgColor="1F3864")
        header_font = Font(bold=True, color="FFFFFF", name="Arial", size=10)
        header_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
        thin = Side(style="thin", color="CCCCCC")
        cell_border = Border(left=thin, right=thin, top=thin, bottom=thin)

        for col_idx in range(1, len(OUTPUT_COLUMNS) + 1):
            cell = ws.cell(row=1, column=col_idx)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = header_align
            cell.border = cell_border
        ws.row_dimensions[1].height = 40

        # Column widths
        col_widths = {
            "Project": 28, "Phase": 8, "Unique ID": 32,
            "Funding Amount (Original)": 28, "Funding Amount (Billion €)": 22,
            "Funding Source (Original)": 40, "Funding Source (Standardized)": 32,
            "Funding Source origin country": 28,
            "Type of Support (Original)": 28, "Type of Support (Standardized)": 26,
            "Year / Timeline": 16, "References": 55,
            "Destination": 18, "Origin region": 18,
            "Search Query Used": 42, "Confidence": 12, "Scraped At": 18,
        }
        for col_idx, col_name in enumerate(OUTPUT_COLUMNS, 1):
            ws.column_dimensions[get_column_letter(col_idx)].width = col_widths.get(col_name, 20)

        # Row styling with confidence color coding
        confidence_fill = {
            "high":   PatternFill("solid", fgColor="E2EFDA"),  # green
            "medium": PatternFill("solid", fgColor="FFF2CC"),  # yellow
            "low":    PatternFill("solid", fgColor="FCE4D6"),  # red/orange
        }
        alt_fill = PatternFill("solid", fgColor="F7F9FF")
        conf_col = OUTPUT_COLUMNS.index("Confidence") + 1

        for row_idx in range(2, len(df) + 2):
            conf_val = ws.cell(row=row_idx, column=conf_col).value
            for col_idx in range(1, len(OUTPUT_COLUMNS) + 1):
                cell = ws.cell(row=row_idx, column=col_idx)
                cell.font = Font(name="Arial", size=10)
                cell.alignment = Alignment(vertical="center", wrap_text=True)
                cell.border = cell_border
                if col_idx == conf_col:
                    cell.fill = confidence_fill.get(conf_val, PatternFill())
                elif row_idx % 2 == 0:
                    cell.fill = alt_fill

        ws.freeze_panes = "A2"

        # Metadata sheet
        ws_meta = wb.create_sheet("Metadata")
        meta_rows = [
            ("Segment", segment),
            ("Generated at", datetime.now().strftime("%Y-%m-%d %H:%M")),
            ("Row count", len(df)),
            ("Claude model", CLAUDE_MODEL),
        ]
        for r, (k, v) in enumerate(meta_rows, 1):
            ws_meta.cell(row=r, column=1, value=k).font = Font(bold=True, name="Arial")
            ws_meta.cell(row=r, column=2, value=v).font = Font(name="Arial")
        ws_meta.column_dimensions["A"].width = 20
        ws_meta.column_dimensions["B"].width = 40

    log.info(f"Excel exported: {output_path} ({len(df)} rows)")


# ---------------------------------------------------------------------------
# Main pipeline
# ---------------------------------------------------------------------------

async def run(
    segment: str,
    project_list: list[str] | None,
    output_path: str,
    max_pages: int,
):
    api_key = os.environ.get("ANTHROPIC_API_KEY")
    if not api_key:
        raise EnvironmentError(
            "ANTHROPIC_API_KEY environment variable not set.\n"
            "Windows:    setx ANTHROPIC_API_KEY sk-ant-xxxx\n"
            "macOS/Linux: export ANTHROPIC_API_KEY=sk-ant-xxxx"
        )

    client = anthropic.Anthropic(api_key=api_key)
    all_rows = []

    # Build task list: (query, project_hint)
    tasks: list[tuple[str, str]] = []
    if project_list:
        for project in project_list:
            for query in build_queries_for_project(project, segment):
                tasks.append((query, project))
    else:
        for query in build_queries_for_segment(segment):
            tasks.append((query, segment))

    log.info(f"Segment : {segment}")
    log.info(f"Projects: {len(project_list) if project_list else 'none (segment-only mode)'}")
    log.info(f"Queries planned: {len(tasks)}")

    visited_urls: set[str] = set()

    async with async_playwright() as pw:
        browser = await pw.chromium.launch(headless=True)
        context = await browser.new_context(
            user_agent=(
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/120.0.0.0 Safari/537.36"
            )
        )
        page = await context.new_page()

        for query, project_hint in tasks:
            log.info(f"\n[{project_hint}] Query: {query[:80]}")
            results = await google_search(page, query)
            results = prioritize_urls(results)

            pages_read = 0
            for result in results:
                url = result["url"]
                if url in visited_urls or pages_read >= max_pages:
                    continue

                log.info(f"  → {url[:90]}")
                text = await fetch_page_text(page, url)
                visited_urls.add(url)
                pages_read += 1

                if len(text) < 200:
                    log.info("    (page too short, skipped)")
                    continue

                investments = extract_investments(client, text, url, segment, project_hint)

                if investments:
                    log.info(f"    {len(investments)} investment(s) extracted")
                    for inv in investments:
                        all_rows.append(investment_to_row(inv, url, query))
                else:
                    log.info("    (no investment found on this page)")

        await browser.close()

    log.info(f"\nTotal before deduplication: {len(all_rows)} rows")
    all_rows = deduplicate(all_rows)
    log.info(f"Total after deduplication:  {len(all_rows)} rows")

    if all_rows:
        export_to_excel(all_rows, output_path, segment)
    else:
        log.warning("No investments found. Try broadening the segment or project list.")

    return all_rows


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(
        description="Battery Investment Scraper — search-first, project-aware"
    )
    parser.add_argument(
        "--segment",
        default="gigafactory battery cell production Europe",
        help='Value chain segment, e.g. "gigafactory battery cell Europe"',
    )
    parser.add_argument(
        "--from-excel",
        default=None,
        metavar="EXCEL_PATH",
        help="Path to an Excel file. Reads unique project names from first column of Investments_GF sheet.",
    )
    parser.add_argument(
        "--projects",
        default=None,
        help='Comma-separated project names, e.g. "Northvolt Skelleftea,ACC Douvrin,Verkor Dunkirk"',
    )
    parser.add_argument(
        "--output",
        default=None,
        help="Output Excel path (default: auto-generated from segment + date)",
    )
    parser.add_argument(
        "--max-pages",
        type=int,
        default=MAX_PAGES_PER_PROJECT,
        help=f"Max pages to read per project query (default: {MAX_PAGES_PER_PROJECT})",
    )
    args = parser.parse_args()

    # Resolve project list
    project_list = None
    if args.from_excel:
        project_list = read_projects_from_excel(args.from_excel)
    elif args.projects:
        project_list = [p.strip() for p in args.projects.split(",") if p.strip()]

    # Resolve output path
    if args.output:
        output_path = args.output
    else:
        slug = re.sub(r"[^\w]+", "_", args.segment.lower())[:40]
        date_str = datetime.now().strftime("%Y%m%d_%H%M")
        output_path = f"investments_{slug}_{date_str}.xlsx"

    log.info("=" * 60)
    log.info("Battery Investment Scraper")
    log.info(f"Segment  : {args.segment}")
    log.info(f"Projects : {len(project_list) if project_list else 'segment-only'}")
    log.info(f"Output   : {output_path}")
    log.info("=" * 60)

    asyncio.run(
        run(
            segment=args.segment,
            project_list=project_list,
            output_path=output_path,
            max_pages=args.max_pages,
        )
    )


if __name__ == "__main__":
    main()
