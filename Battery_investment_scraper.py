"""
Battery Investment Scraper
==========================
Automatically searches and extracts investment data for battery value chain projects.
Reads project list from a Google Sheet and writes results to a new Google Sheet
in the same Drive folder.

Usage:
    # Read projects from Gigafactories Google Sheet (default)
    python battery_investment_scraper.py

    # Different segment (e.g. active materials)
    python battery_investment_scraper.py --segment "battery active materials cathode anode Europe"

    # Pass a custom project list instead
    python battery_investment_scraper.py --projects "Northvolt Skelleftea,ACC Douvrin,Verkor Dunkirk"

    # Limit pages per project (faster / cheaper for testing)
    python battery_investment_scraper.py --max-pages 2

Prerequisites:
    pip install anthropic playwright pandas gspread google-auth
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

import anthropic
import gspread
from google.oauth2.service_account import Credentials
from playwright.async_api import async_playwright

# ---------------------------------------------------------------------------
# Configuration — edit these if paths change
# ---------------------------------------------------------------------------

CREDENTIALS_FILE    = r"C:/Users/TE/Documents/battery-investments-42fad34978f6.json"
SOURCE_SHEET_ID     = "1QoVDcbHLV8ES2fcpM2nxsoOc41eMP0AA0O5OcADXzjo"
SOURCE_TAB_NAME     = "Investments_GF"
DEST_FOLDER_ID      = "1vYnUtLeHk0DSVfxdzAcKFcdWhn1yl_mF"

CLAUDE_MODEL        = "claude-sonnet-4-20250514"
MAX_TOKENS          = 2000
PAGE_DELAY_S        = 2
MAX_PAGES_PER_QUERY = 4

GOOGLE_SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

PRIORITY_DOMAINS = [
    "electrive.com", "electrive.net", "cinea.ec.europa.eu",
    "eib.org", "ipcei-batteries.eu", "ec.europa.eu",
    "benchmarkminerals.com", "spglobal.com", "reuters.com",
    "ft.com", "bloomberg.com",
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
    "Search Query Used",
    "Confidence",
    "Scraped At",
]

FUNDING_SOURCE_CATEGORIES = [
    "National/State Government", "EU Innovation Fund",
    "European Investment Bank (EIB)", "Commercial Banks",
    "Private Equity / VC", "Vehicle Manufacturers (OEMs)",
    "Horizon Europe", "Multilateral", "Other",
]
SUPPORT_TYPE_CATEGORIES  = ["Grant", "Loan", "Equity", "Guarantee", "Other"]
ORIGIN_REGION_CATEGORIES = ["Europe", "North America", "Asia", "Other"]

# Header row background / text colours (hex, no #)
HEADER_BG    = {"red": 0.122, "green": 0.220, "blue": 0.392}   # #1F3864
HEADER_FG    = {"red": 1.0,   "green": 1.0,   "blue": 1.0}     # white
CONF_COLORS  = {
    "high":   {"red": 0.886, "green": 0.937, "blue": 0.855},   # #E2EFDA green
    "medium": {"red": 1.0,   "green": 0.949, "blue": 0.800},   # #FFF2CC yellow
    "low":    {"red": 0.988, "green": 0.894, "blue": 0.839},   # #FCE4D6 orange
}

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  %(message)s",
    datefmt="%H:%M:%S",
)
log = logging.getLogger(__name__)


# ---------------------------------------------------------------------------
# Google Sheets helpers
# ---------------------------------------------------------------------------

def get_gspread_client() -> gspread.Client:
    creds = Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=GOOGLE_SCOPES)
    return gspread.authorize(creds)


def read_projects_from_gsheet(gc: gspread.Client) -> list[str]:
    """Return unique project names from the first column of the source sheet."""
    sh = gc.open_by_key(SOURCE_SHEET_ID)
    ws = sh.worksheet(SOURCE_TAB_NAME)
    # First column, skip header
    values = ws.col_values(1)[1:]
    projects = list(dict.fromkeys(v.strip() for v in values if v.strip()))
    log.info(f"Loaded {len(projects)} unique projects from Google Sheet")
    return projects


def create_result_gsheet(gc: gspread.Client, title: str) -> gspread.Spreadsheet:
    """Create a new Google Sheet in DEST_FOLDER_ID and return it."""
    sh = gc.create(title, folder_id=DEST_FOLDER_ID)
    log.info(f"Created Google Sheet '{title}' (id: {sh.id})")
    return sh


def write_results_to_gsheet(sh: gspread.Spreadsheet, rows: list[dict], segment: str):
    """
    Write all results to the Google Sheet:
      - Sheet 1 'Investments': data with header formatting and confidence colours
      - Sheet 2 'Metadata': run info
    """
    # ---- Investments tab ----
    ws = sh.sheet1
    ws.update_title("Investments")

    # Build 2D array: header + data rows
    data = [OUTPUT_COLUMNS]
    for row in rows:
        data.append([str(row.get(col, "")) for col in OUTPUT_COLUMNS])

    ws.update(data, "A1")

    n_rows = len(data)       # including header
    n_cols = len(OUTPUT_COLUMNS)

    # Batch format requests
    requests = []

    # 1. Header row: dark blue background, white bold text, centre-aligned
    requests.append({
        "repeatCell": {
            "range": {
                "sheetId": ws.id, "startRowIndex": 0, "endRowIndex": 1,
                "startColumnIndex": 0, "endColumnIndex": n_cols,
            },
            "cell": {
                "userEnteredFormat": {
                    "backgroundColor": HEADER_BG,
                    "textFormat": {"bold": True, "foregroundColor": HEADER_FG,
                                   "fontFamily": "Arial", "fontSize": 10},
                    "horizontalAlignment": "CENTER",
                    "verticalAlignment": "MIDDLE",
                    "wrapStrategy": "WRAP",
                }
            },
            "fields": "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment,wrapStrategy)",
        }
    })

    # 2. Freeze header row
    requests.append({
        "updateSheetProperties": {
            "properties": {"sheetId": ws.id, "gridProperties": {"frozenRowCount": 1}},
            "fields": "gridProperties.frozenRowCount",
        }
    })

    # 3. Alternate row fill (light blue) for even data rows
    alt_color = {"red": 0.969, "green": 0.976, "blue": 1.0}  # #F7F9FF
    for i in range(1, n_rows, 2):  # 0-indexed: rows 2,4,6... → indices 1,3,5...
        requests.append({
            "repeatCell": {
                "range": {
                    "sheetId": ws.id,
                    "startRowIndex": i, "endRowIndex": i + 1,
                    "startColumnIndex": 0, "endColumnIndex": n_cols,
                },
                "cell": {"userEnteredFormat": {"backgroundColor": alt_color}},
                "fields": "userEnteredFormat.backgroundColor",
            }
        })

    # 4. Confidence column colour coding
    conf_col_idx = OUTPUT_COLUMNS.index("Confidence")
    for row_idx, row in enumerate(rows, start=1):   # 1-indexed data rows
        conf = row.get("Confidence", "low")
        color = CONF_COLORS.get(conf, CONF_COLORS["low"])
        requests.append({
            "repeatCell": {
                "range": {
                    "sheetId": ws.id,
                    "startRowIndex": row_idx, "endRowIndex": row_idx + 1,
                    "startColumnIndex": conf_col_idx,
                    "endColumnIndex": conf_col_idx + 1,
                },
                "cell": {"userEnteredFormat": {"backgroundColor": color}},
                "fields": "userEnteredFormat.backgroundColor",
            }
        })

    # 5. Column widths (pixels)
    col_widths_px = {
        "Project": 200, "Phase": 60, "Unique ID": 240,
        "Funding Amount (Original)": 200, "Funding Amount (Billion €)": 160,
        "Funding Source (Original)": 280, "Funding Source (Standardized)": 230,
        "Funding Source origin country": 200,
        "Type of Support (Original)": 200, "Type of Support (Standardized)": 190,
        "Year / Timeline": 120, "References": 380,
        "Destination": 130, "Origin region": 130,
        "Search Query Used": 300, "Confidence": 90, "Scraped At": 140,
    }
    for col_idx, col_name in enumerate(OUTPUT_COLUMNS):
        requests.append({
            "updateDimensionProperties": {
                "range": {
                    "sheetId": ws.id, "dimension": "COLUMNS",
                    "startIndex": col_idx, "endIndex": col_idx + 1,
                },
                "properties": {"pixelSize": col_widths_px.get(col_name, 150)},
                "fields": "pixelSize",
            }
        })

    # 6. Header row height
    requests.append({
        "updateDimensionProperties": {
            "range": {"sheetId": ws.id, "dimension": "ROWS",
                      "startIndex": 0, "endIndex": 1},
            "properties": {"pixelSize": 50},
            "fields": "pixelSize",
        }
    })

    sh.batch_update({"requests": requests})

    # ---- Metadata tab ----
    ws_meta = sh.add_worksheet(title="Metadata", rows=10, cols=2)
    ws_meta.update([
        ["Field", "Value"],
        ["Segment", segment],
        ["Generated at", datetime.now().strftime("%Y-%m-%d %H:%M")],
        ["Row count", str(len(rows))],
        ["Claude model", CLAUDE_MODEL],
        ["Source sheet ID", SOURCE_SHEET_ID],
        ["Source tab", SOURCE_TAB_NAME],
    ], "A1")

    log.info(f"Results written to Google Sheet: {sh.url}")


# ---------------------------------------------------------------------------
# Search query generation
# ---------------------------------------------------------------------------

def build_queries_for_project(project: str, segment: str) -> list[str]:
    return [
        f'"{project}" investment funding million billion',
        f'"{project}" grant loan equity financing battery',
        f'"{project}" European Investment Bank IPCEI Innovation Fund',
        f'"{project}" battery gigafactory funding announcement',
    ]


def build_queries_for_segment(segment: str) -> list[str]:
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
    url = f"https://www.google.com/search?q={query.replace(' ', '+')}&num={max_results}&hl=en&gl=US"
    results = []
    try:
        await page.goto(url, wait_until="domcontentloaded", timeout=15000)
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
            u = link["href"]
            if u not in seen:
                seen.add(u)
                results.append({"url": u, "title": link.get("text", "")[:120]})
            if len(results) >= max_results:
                break
    except Exception as e:
        log.warning(f"Google search error: {e}")
    return results


async def fetch_page_text(page, url: str, max_chars: int = 8000) -> str:
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
    phase   = inv.get("phase", 1)
    return {
        "Project":                        project,
        "Phase":                          phase,
        "Unique ID":                      f"{project}, Phase {phase}",
        "Funding Amount (Original)":      inv.get("funding_amount_original", "Info Not Available"),
        "Funding Amount (Billion €)":     inv.get("funding_amount_bn_eur", "-"),
        "Funding Source (Original)":      inv.get("funding_source_original", "-"),
        "Funding Source (Standardized)":  inv.get("funding_source_standardized", "-"),
        "Funding Source origin country":  inv.get("funding_source_country", "-"),
        "Type of Support (Original)":     inv.get("support_type_original", "-"),
        "Type of Support (Standardized)": inv.get("support_type_standardized", "-"),
        "Year / Timeline":                inv.get("year_timeline", "-"),
        "References":                     url,
        "Destination":                    inv.get("destination_country", "-"),
        "Origin region":                  inv.get("origin_region", "Europe"),
        "Search Query Used":              query,
        "Confidence":                     inv.get("confidence", "low"),
        "Scraped At":                     datetime.now().strftime("%Y-%m-%d %H:%M"),
    }


def deduplicate(rows: list[dict]) -> list[dict]:
    rank = {"high": 3, "medium": 2, "low": 1}
    seen = {}
    for row in rows:
        key = (
            str(row.get("Project", "")).lower().strip(),
            str(row.get("Funding Amount (Original)", "")).lower().strip(),
            str(row.get("Funding Source (Original)", "")).lower().strip(),
        )
        score = rank.get(row.get("Confidence", "low"), 1)
        if key not in seen or score > rank.get(seen[key].get("Confidence", "low"), 1):
            seen[key] = row
    return list(seen.values())


# ---------------------------------------------------------------------------
# Main pipeline
# ---------------------------------------------------------------------------

async def run(segment: str, project_list: list[str] | None, max_pages: int):
    api_key = os.environ.get("ANTHROPIC_API_KEY")
    if not api_key:
        raise EnvironmentError(
            "ANTHROPIC_API_KEY environment variable not set.\n"
            "Windows:     setx ANTHROPIC_API_KEY sk-ant-xxxx\n"
            "macOS/Linux: export ANTHROPIC_API_KEY=sk-ant-xxxx"
        )

    client = anthropic.Anthropic(api_key=api_key)

    # Build (query, project_hint) task list
    tasks: list[tuple[str, str]] = []
    if project_list:
        for project in project_list:
            for query in build_queries_for_project(project, segment):
                tasks.append((query, project))
    else:
        for query in build_queries_for_segment(segment):
            tasks.append((query, segment))

    log.info(f"Segment : {segment}")
    log.info(f"Projects: {len(project_list) if project_list else 'none (segment-only)'}")
    log.info(f"Queries planned: {len(tasks)}")

    all_rows: list[dict] = []
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
            log.info(f"\n[{project_hint}] {query[:80]}")
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
                    log.info("    (no investment found)")

        await browser.close()

    log.info(f"\nTotal before dedup: {len(all_rows)} rows")
    all_rows = deduplicate(all_rows)
    log.info(f"Total after dedup:  {len(all_rows)} rows")

    if not all_rows:
        log.warning("No investments found. Try broadening segment or project list.")
        return

    # Write to Google Sheets
    slug      = re.sub(r"[^\w]+", "_", segment.lower())[:35]
    date_str  = datetime.now().strftime("%Y%m%d_%H%M")
    sheet_title = f"investments_{slug}_{date_str}"

    gc = get_gspread_client()
    sh = create_result_gsheet(gc, sheet_title)
    write_results_to_gsheet(sh, all_rows, segment)

    log.info(f"\nDone — {len(all_rows)} rows written.")
    log.info(f"Open: {sh.url}")


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(description="Battery Investment Scraper → Google Sheets")
    parser.add_argument(
        "--segment",
        default="gigafactory battery cell production Europe",
        help='Value chain segment, e.g. "gigafactory battery cell Europe"',
    )
    parser.add_argument(
        "--projects",
        default=None,
        help='Comma-separated project names (overrides reading from source sheet)',
    )
    parser.add_argument(
        "--no-gsheet-source",
        action="store_true",
        help="Do not read projects from source Google Sheet (use --projects or segment-only)",
    )
    parser.add_argument(
        "--max-pages",
        type=int,
        default=MAX_PAGES_PER_QUERY,
        help=f"Max pages to read per query (default: {MAX_PAGES_PER_QUERY})",
    )
    args = parser.parse_args()

    # Resolve project list
    project_list = None
    if args.projects:
        project_list = [p.strip() for p in args.projects.split(",") if p.strip()]
        log.info(f"Using {len(project_list)} projects from --projects flag")
    elif not args.no_gsheet_source:
        gc = get_gspread_client()
        project_list = read_projects_from_gsheet(gc)

    log.info("=" * 60)
    log.info("Battery Investment Scraper")
    log.info(f"Segment  : {args.segment}")
    log.info(f"Projects : {len(project_list) if project_list else 'segment-only'}")
    log.info("=" * 60)

    asyncio.run(run(
        segment=args.segment,
        project_list=project_list,
        max_pages=args.max_pages,
    ))


if __name__ == "__main__":
    main()
