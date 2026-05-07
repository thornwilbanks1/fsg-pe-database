"""
FSG PE Database — API Loop Script
===================================
Researches PE firms from a CSV list and writes structured portfolio company
data directly to your Excel database. Runs fully unattended — no "continue"
prompts, no response length limits, no session timeouts.

SETUP (one time):
    pip install anthropic httpx beautifulsoup4 lxml openpyxl rapidfuzz

USAGE:
    python3 fsg_pipeline.py

CONFIGURATION:
    Edit the CONFIG block below before running.

OUTPUT:
    - STEPHENS_FSG_DATABASE.xlsx  — updated database
    - pipeline_log.csv            — per-firm results
    - pipeline_errors.log         — any errors for review
"""

import os, csv, json, re, time, logging
from datetime import datetime
from pathlib import Path

import httpx
import anthropic
from bs4 import BeautifulSoup
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from rapidfuzz import fuzz

# ── CONFIG ─────────────────────────────────────────────────────────────────────
API_KEY      = os.getenv("ANTHROPIC_API_KEY", "")   # or paste key here
MODEL        = "claude-sonnet-4-20250514"
DB_FILE      = "2026_04_24_CLAUDE_DATABASE_PRIVATE_EQUITY_RESEARCH_v8.xlsx"
FIRMS_CSV    = "firms_to_research.csv"
OUTPUT_DB    = "STEPHENS_FSG_DATABASE.xlsx"
LOG_CSV      = "pipeline_log.csv"
ERROR_LOG    = "pipeline_errors.log"

DELAY_BETWEEN_FIRMS = 2     # seconds between firms (be polite)
MAX_HTML_CHARS      = 14000  # truncation before sending to API
FUZZY_DUPE_THRESH   = 88    # rapidfuzz score to flag near-duplicates

# ── LOGGING ────────────────────────────────────────────────────────────────────
logging.basicConfig(
    filename=ERROR_LOG,
    level=logging.WARNING,
    format="%(asctime)s [%(levelname)s] %(message)s"
)

# ── EXTRACTION PROMPT ──────────────────────────────────────────────────────────
PORTFOLIO_PROMPT = """\
You are a private equity analyst. Extract ALL portfolio companies from the HTML below.

Return ONLY a raw JSON array. No markdown, no preamble.

Each object must have exactly these keys:
- "portfolio_company": string (exact company name as shown)
- "status": "Current" or "Realized" (default "Current" if unclear)
- "sector": string — use ONLY these values:
  Business Services | Consumer | Energy / Power / Infrastructure |
  Financial Services | Healthcare & Life Sciences | Industrials |
  Real Estate | Technology | Transportation & Logistics
- "sub_sector": string (specific, e.g. "HVAC Services", "SaaS", "Staffing")
- "hq": string (City, State, Country — or "" if unknown)
- "transaction_type": one of: Buyout | Growth Capital (Minority) | Growth Equity |
  Co-Investment | Carve-Out | Carve-Out (from Corporate) | Public-to-Private | Recap
- "investing_fund": string (fund name if mentioned, else "")
- "entry_date": string in Mon-YY format (e.g. "Aug-24") — or "" if unknown
- "prior_owner": string (who sold to the PE firm — "Founder-owned", a PE firm name,
  or a corporate name) — or "" if unknown
- "sell_side_advisor": string (bank advising the SELLER) — or ""
- "entry_press_release": string (full URL if mentioned) — or ""
- "pe_firm_deal_team": string (names + titles quoted in release) — or ""
- "exit_date": "---" if Current, Mon-YY or "" if Realized
- "exit_buyer": "---" if Current, buyer name or "" if Realized
- "buyer_type": "---" if Current, one of Sponsor|Strategic|IPO|Other if Realized
- "exit_press_release": "---" if Current, URL or "" if Realized
- "website": string (bare domain, no https://, e.g. "company.com") — or ""
- "company_description": string (2-3 sentence evergreen description) — or ""
- "data_quality_notes": string (source conflicts, structural notes) — or "---"

Rules:
- Include BOTH current and realized companies
- Do NOT include the PE firm itself
- Do NOT invent data — use "" for unknown fields
- Return [] if no companies found

HTML:
{html}
"""

ENRICH_PROMPT = """\
You are a private equity research analyst. Given the PE firm name and one portfolio
company name, find and return structured data from press releases and public sources.

PE Firm: {firm}
Portfolio Company: {company}

Search your knowledge for:
1. The investment press release or announcement
2. The sell-side advisor (bank advising the SELLER)
3. Aurora/PE firm deal team members quoted
4. Prior owner (who sold the company)
5. Investing fund name
6. Entry date (Mon-YY format)
7. Company HQ (City, State, Country)
8. Company website (bare domain)

Return ONLY a JSON object with these keys (use "" for anything not found):
{{
  "sell_side_advisor": "",
  "entry_press_release": "",
  "pe_firm_deal_team": "",
  "prior_owner": "",
  "investing_fund": "",
  "entry_date": "",
  "hq": "",
  "website": "",
  "company_description": "",
  "data_quality_notes": ""
}}
"""

# ── HELPERS ────────────────────────────────────────────────────────────────────

def fetch_page(url: str) -> str | None:
    """Fetch a URL; return clean text or None."""
    headers = {
        "User-Agent": (
            "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
            "AppleWebKit/537.36 Chrome/124.0 Safari/537.36"
        ),
        "Accept": "text/html,application/xhtml+xml,*/*;q=0.8",
    }
    try:
        r = httpx.get(url, headers=headers, timeout=20, follow_redirects=True)
        if r.status_code == 200 and len(r.text) > 2000:
            return r.text
    except Exception:
        pass
    return None


def clean_html(raw: str) -> str:
    """Strip noise, return truncated clean HTML."""
    soup = BeautifulSoup(raw, "lxml")
    for tag in soup(["script","style","nav","footer","header","noscript","svg","iframe"]):
        tag.decompose()
    return soup.decode()[:MAX_HTML_CHARS]


def find_portfolio_url(base_url: str, given_suffix: str = "") -> str | None:
    """Return the portfolio page URL."""
    base = base_url.rstrip("/")
    if given_suffix:
        url = base + "/" + given_suffix.lstrip("/")
        if fetch_page(url):
            return url
    candidates = [
        "/portfolio", "/portfolio-companies", "/our-portfolio",
        "/investments", "/companies", "/our-companies",
        "/portfolio/current", "/portfolio/active",
    ]
    for suffix in candidates:
        url = base + suffix
        html = fetch_page(url)
        if html and len(html) > 3000:
            return url
    return None


def call_claude(prompt: str, client: anthropic.Anthropic) -> str:
    """Single Claude API call; return text content."""
    response = client.messages.create(
        model=MODEL,
        max_tokens=4096,
        messages=[{"role": "user", "content": prompt}]
    )
    return response.content[0].text.strip()


def parse_json(text: str) -> list | dict | None:
    """Strip markdown fences and parse JSON."""
    text = re.sub(r"^```json\s*", "", text)
    text = re.sub(r"\s*```$", "", text)
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        return None


def extract_portcos(html: str, firm_name: str, client: anthropic.Anthropic) -> list[dict]:
    """Send cleaned HTML to Claude; return portco list."""
    clean = clean_html(html)
    prompt = PORTFOLIO_PROMPT.format(html=clean)
    raw = call_claude(prompt, client)
    result = parse_json(raw)
    if not isinstance(result, list):
        logging.warning(f"{firm_name}: JSON parse failed — {raw[:200]}")
        return []
    # Tag each with PE firm
    for p in result:
        p["pe_firm"] = firm_name
    return result


def enrich_portco(firm: str, company: str, client: anthropic.Anthropic) -> dict:
    """Ask Claude to fill gaps using knowledge (no web fetch)."""
    prompt = ENRICH_PROMPT.format(firm=firm, company=company)
    raw = call_claude(prompt, client)
    result = parse_json(raw)
    if isinstance(result, dict):
        return result
    return {}


# ── VALIDATION ─────────────────────────────────────────────────────────────────

VALID_STATUS     = {"Current", "Realized"}
VALID_BUYER_TYPE = {"Sponsor", "Strategic", "IPO", "Other", "---", ""}
VALID_TX_TYPE    = {
    "Buyout", "Growth Capital (Minority)", "Growth Equity",
    "Co-Investment", "Carve-Out", "Carve-Out (from Corporate)",
    "Public-to-Private", "Recap", ""
}
DATE_RE = re.compile(r"^(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)-\d{2}$")

def validate(row: dict, existing: set, firm: str) -> tuple[bool, list[str], list[str]]:
    """
    Returns (ok_to_write, hard_failures, warnings).
    Hard failures block write; warnings write with flag.
    """
    hard, warn = [], []

    # R01 — required fields
    for f in ("pe_firm", "portfolio_company", "status"):
        if not row.get(f, "").strip():
            hard.append(f"R01: '{f}' is blank")

    # R02 — status enum
    if row.get("status") not in VALID_STATUS:
        hard.append(f"R02: status='{row.get('status')}' invalid")

    # R03 — buyer_type enum
    bt = row.get("buyer_type", "")
    if bt not in VALID_BUYER_TYPE:
        hard.append(f"R03: buyer_type='{bt}' invalid")

    # R04 — transaction_type enum
    tt = row.get("transaction_type", "")
    if tt not in VALID_TX_TYPE:
        warn.append(f"R04: transaction_type='{tt}' not in standard list")

    # R05 — date format
    for df in ("entry_date", "exit_date"):
        val = row.get(df, "")
        if val and val not in ("---", "", "Not found publicly available"):
            if not DATE_RE.match(val):
                hard.append(f"R05: {df}='{val}' not in Mon-YY format")

    # R06 — consistency
    if row.get("status") == "Current":
        for ef in ("exit_date", "exit_buyer", "exit_press_release"):
            val = row.get(ef, "")
            if val and val not in ("---", ""):
                warn.append(f"R06: status=Current but {ef}='{val}' populated")
        # Auto-fix
        row["exit_date"] = "---"
        row["exit_buyer"] = "---"
        row["buyer_type"] = "---"
        row["exit_press_release"] = "---"

    # R07 — exact duplicate
    key = (row.get("pe_firm","").lower(), row.get("portfolio_company","").lower())
    if key in existing:
        hard.append(f"R07: Duplicate — ({firm} / {row.get('portfolio_company')})")

    # W02 — website format
    site = row.get("website", "")
    if site:
        clean_site = re.sub(r"^https?://", "", site).rstrip("/")
        clean_site = re.sub(r"^www\.", "", clean_site).split("/")[0]
        if clean_site != site:
            warn.append(f"W02: website auto-corrected to '{clean_site}'")
            row["website"] = clean_site

    # W06 — fuzzy duplicate
    firm_lower = row.get("pe_firm","").lower()
    portco_lower = row.get("portfolio_company","").lower()
    for (ef, ec) in existing:
        if ef == firm_lower and ec != portco_lower:
            score = fuzz.ratio(ec, portco_lower)
            if score >= FUZZY_DUPE_THRESH:
                warn.append(f"W06: Near-duplicate '{ec}' ({score}% match)")

    ok = len(hard) == 0
    return ok, hard, warn


# ── EXCEL WRITER ───────────────────────────────────────────────────────────────

# Column order matching the database schema
SCHEMA_COLS = [
    "pe_firm", "portfolio_company", "status", "sector", "sub_sector",
    "hq", "transaction_type", "investing_fund", "entry_date",
    "prior_owner", "sell_side_advisor", "entry_press_release",
    "pe_firm_deal_team", "exit_date", "exit_buyer", "buyer_type",
    "exit_press_release", "website", "company_description",
    "business_characteristics", "end_markets", "data_quality_notes",
]

GREEN = PatternFill("solid", fgColor="E8F5E9")
BLUE  = PatternFill("solid", fgColor="D6E4F0")
WRAP  = Alignment(wrap_text=True, vertical="top")
THIN  = Side(style="thin", color="CCCCCC")
BDR   = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
DFONT = Font(name="Arial", size=9)


def load_existing(ws) -> set:
    """Build (firm.lower, company.lower) set from DB."""
    existing = set()
    for r in range(2, ws.max_row + 1):
        f = str(ws.cell(r, 1).value or "").strip().lower()
        c = str(ws.cell(r, 2).value or "").strip().lower()
        if f and c:
            existing.add((f, c))
    return existing


def write_row(ws, row: dict, existing: set):
    """Append one validated row to the Investments sheet."""
    last = ws.max_row + 1
    fill = GREEN if row.get("status") == "Current" else BLUE
    for ci, col in enumerate(SCHEMA_COLS, 1):
        val = row.get(col, "")
        if val is None:
            val = ""
        cell = ws.cell(last, ci, val)
        cell.fill = fill
        cell.alignment = WRAP
        cell.font = DFONT
    ws.row_dimensions[last].height = 55
    key = (row.get("pe_firm","").lower(), row.get("portfolio_company","").lower())
    existing.add(key)


# ── MAIN PIPELINE ──────────────────────────────────────────────────────────────

def process_firm(firm_name: str, website: str, portfolio_suffix: str,
                 ws, existing: set, client: anthropic.Anthropic) -> dict:
    """
    Full pipeline for one firm.
    Returns a result dict for the log.
    """
    result = {
        "firm": firm_name,
        "status": "fail",
        "portcos_found": 0,
        "portcos_written": 0,
        "portcos_failed": 0,
        "hard_failures": "",
        "timestamp": datetime.utcnow().isoformat(),
    }

    print(f"\n{'─'*60}")
    print(f"  {firm_name}")
    print(f"{'─'*60}")

    # Stage 1: Find and fetch portfolio page
    portfolio_url = find_portfolio_url(website, portfolio_suffix)
    if not portfolio_url:
        print(f"  ✗ Portfolio page not found — trying homepage")
        html = fetch_page(website)
        if not html:
            result["hard_failures"] = "Portfolio page and homepage both unreachable"
            print(f"  ✗ Unreachable — skipping")
            return result
        portfolio_url = website
    else:
        html = fetch_page(portfolio_url)

    if not html:
        result["hard_failures"] = "Fetch failed"
        print(f"  ✗ Fetch failed")
        return result

    print(f"  [1] Fetched: {portfolio_url} ({len(html):,} chars)")

    # Stage 2: Extract portcos via Claude API
    portcos = extract_portcos(html, firm_name, client)
    result["portcos_found"] = len(portcos)
    print(f"  [2] Extracted: {len(portcos)} companies")

    if not portcos:
        result["hard_failures"] = "Zero portcos extracted"
        return result

    # Stage 3: Enrich gaps (entry date, advisor, deal team)
    # Only enrich portcos missing key fields to save API calls
    enriched = 0
    for p in portcos:
        missing = not p.get("entry_date") or not p.get("sell_side_advisor")
        if missing:
            gap_data = enrich_portco(firm_name, p["portfolio_company"], client)
            for k, v in gap_data.items():
                if v and not p.get(k):
                    p[k] = v
            enriched += 1
            time.sleep(0.3)  # brief pause between enrichment calls

    if enriched:
        print(f"  [3] Enriched: {enriched} portcos")

    # Stage 4: Validate and write
    written = failed = 0
    failure_notes = []

    for p in portcos:
        ok, hard, warn = validate(p, existing, firm_name)

        # Append any warnings to data_quality_notes
        if warn:
            existing_notes = p.get("data_quality_notes", "---")
            if existing_notes == "---":
                existing_notes = ""
            extra = "; ".join(warn)
            p["data_quality_notes"] = f"{existing_notes} | WARN: {extra}".strip(" |")

        if ok:
            write_row(ws, p, existing)
            written += 1
        else:
            failed += 1
            failure_notes.extend(hard)
            logging.warning(f"{firm_name} / {p.get('portfolio_company')}: {hard}")

    result["portcos_written"] = written
    result["portcos_failed"] = failed
    result["hard_failures"] = "; ".join(failure_notes[:5])  # cap log length
    result["status"] = "success" if written > 0 else "partial"

    print(f"  [4] Written: {written} | Failed: {failed}")
    return result


def main():
    # ── Validate config ────────────────────────────────────────────────────────
    if not API_KEY:
        print("ERROR: Set ANTHROPIC_API_KEY environment variable")
        print("  Mac/Linux: export ANTHROPIC_API_KEY=sk-ant-...")
        print("  Windows:   $env:ANTHROPIC_API_KEY = 'sk-ant-...'")
        return

    if not Path(DB_FILE).exists():
        print(f"ERROR: Database file '{DB_FILE}' not found in current directory")
        print("Download it from the Claude chat session and place it here.")
        return

    if not Path(FIRMS_CSV).exists():
        print(f"ERROR: Firms list '{FIRMS_CSV}' not found")
        print("Create a CSV with columns: firm_name, website, portfolio_url_suffix")
        print("Example row: ABRY Partners,https://www.abry.com,/companies")
        return

    # ── Load firms ─────────────────────────────────────────────────────────────
    firms = []
    with open(FIRMS_CSV, newline="", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        for row in reader:
            firms.append({
                "firm_name": row.get("firm_name", "").strip(),
                "website":   row.get("website", "").strip().rstrip("/"),
                "suffix":    row.get("portfolio_url_suffix", "").strip(),
            })

    # ── Load database ──────────────────────────────────────────────────────────
    import shutil
    shutil.copy(DB_FILE, OUTPUT_DB)
    wb = load_workbook(OUTPUT_DB)
    ws = wb["Investments"]
    existing = load_existing(ws)

    client = anthropic.Anthropic(api_key=API_KEY)

    print(f"\n{'='*60}")
    print(f"  FSG PE DATABASE PIPELINE")
    print(f"{'='*60}")
    print(f"  Firms to process: {len(firms)}")
    print(f"  Existing rows:    {len(existing)}")
    print(f"  Model:            {MODEL}")
    print(f"  Output:           {OUTPUT_DB}")
    print(f"{'='*60}")

    # ── Process firms ──────────────────────────────────────────────────────────
    log_rows = []
    total_written = 0

    for i, firm in enumerate(firms, 1):
        print(f"\n[{i}/{len(firms)}]", end="")
        try:
            result = process_firm(
                firm["firm_name"], firm["website"], firm["suffix"],
                ws, existing, client
            )
            log_rows.append(result)
            total_written += result["portcos_written"]

            # Save after every firm — no data loss if script is interrupted
            wb.save(OUTPUT_DB)
            print(f"  → Saved. DB total: {ws.max_row - 1} rows")

        except Exception as e:
            logging.error(f"{firm['firm_name']}: Unexpected error — {e}")
            print(f"  ✗ Unexpected error: {e}")
            log_rows.append({
                "firm": firm["firm_name"],
                "status": "error",
                "portcos_written": 0,
                "hard_failures": str(e),
                "timestamp": datetime.utcnow().isoformat(),
            })

        # Polite delay between firms
        if i < len(firms):
            time.sleep(DELAY_BETWEEN_FIRMS)

    # ── Write log ──────────────────────────────────────────────────────────────
    log_fields = ["firm","status","portcos_found","portcos_written",
                  "portcos_failed","hard_failures","timestamp"]
    with open(LOG_CSV, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=log_fields, extrasaction="ignore")
        writer.writeheader()
        writer.writerows(log_rows)

    # ── Final summary ──────────────────────────────────────────────────────────
    success = sum(1 for r in log_rows if r["status"] == "success")
    partial = sum(1 for r in log_rows if r["status"] == "partial")
    failed  = sum(1 for r in log_rows if r["status"] in ("fail","error"))

    print(f"\n{'='*60}")
    print(f"  PIPELINE COMPLETE")
    print(f"{'='*60}")
    print(f"  Firms:          {len(firms)} total")
    print(f"  Success:        {success}")
    print(f"  Partial:        {partial}")
    print(f"  Failed:         {failed}")
    print(f"  Portcos added:  {total_written}")
    print(f"  DB total rows:  {ws.max_row - 1}")
    print(f"\n  Output files:")
    print(f"  → {OUTPUT_DB}")
    print(f"  → {LOG_CSV}")
    if failed:
        print(f"  → {ERROR_LOG}  ({failed} firms need review)")
    print(f"{'='*60}\n")


if __name__ == "__main__":
    main()
