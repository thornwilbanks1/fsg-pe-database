# FSG PE Database — API Loop Script

Runs the full research pipeline for every firm in your CSV, unattended.
No response length limits. No "continue" prompts. Saves after every firm.

---

## Setup (one time, ~2 minutes)

**1. Install Python packages**
```bash
pip install anthropic httpx beautifulsoup4 lxml openpyxl rapidfuzz
```

**2. Get your Anthropic API key**
Go to console.anthropic.com → API Keys → Create Key

**3. Set the key as an environment variable**
```bash
# Mac / Linux
export ANTHROPIC_API_KEY=sk-ant-api03-...

# Windows PowerShell
$env:ANTHROPIC_API_KEY = "sk-ant-api03-..."
```

**4. Place these files in one folder:**
- `fsg_pipeline.py`          — this script
- `firms_to_research.csv`    — your firm list
- `2026_04_24_CLAUDE_DATABASE_PRIVATE_EQUITY_RESEARCH_v8.xlsx` — your database

---

## firms_to_research.csv format

```
firm_name,website,portfolio_url_suffix
Rotunda Capital Partners,https://rotundacapital.com,/portfolio
ICV Partners,https://icvpartners.com,/portfolio
Pfingsten Partners,https://pfingsten.com,
```

- `portfolio_url_suffix` is optional — script auto-detects if blank
- The shipped `firms_to_research.csv` already contains all 483 firms from the v8 master "Private Equity Firms" sheet

## Resume / restart

The script saves after every firm and supports firm-level resume out of the box:
- On startup it scans `STEPHENS_FSG_DATABASE.xlsx` and skips any firm that already has ≥1 row
- If you `Ctrl-C` and re-run `python3 fsg_pipeline.py`, it picks up where it left off
- To force a re-run for one firm, delete its rows from the output workbook first
- Toggle off via `SKIP_IF_FIRM_HAS_ROWS = False` at the top of the script

## Retries

`fetch_page` and `call_claude` retry transient errors up to 3 times with
exponential backoff (4s, 8s, 16s). Tune via `MAX_RETRIES` / `RETRY_BACKOFF_BASE`.

---

## Run it

```bash
python3 fsg_pipeline.py
```

Walk away. The script:
1. Reads each firm from the CSV
2. Fetches and parses the portfolio page
3. Extracts portcos via Claude API
4. Enriches missing fields (entry date, advisor, deal team)
5. Validates every row (7 hard rules + 6 soft rules)
6. Writes valid rows to `STEPHENS_FSG_DATABASE.xlsx`
7. **Saves after every firm** — safe to interrupt and restart
8. Logs results to `pipeline_log.csv`

---

## Output files

| File | Contents |
|---|---|
| `STEPHENS_FSG_DATABASE.xlsx` | Updated database (never overwrites your original) |
| `pipeline_log.csv` | Per-firm: portcos found / written / failed |
| `pipeline_errors.log` | Hard validation failures and fetch errors |

---

## Cost estimate

~$3–6 for all 472 firms at claude-sonnet-4 pricing (~$3/MTok).
Roughly $0.006–0.012 per firm.

## Speed estimate

~2–4 minutes per firm (fetch + 2 API calls + write).
472 firms ≈ 16–32 hours unattended.
Run overnight or over a weekend.

---

## Restarting after interruption

The script saves after every firm. If interrupted, it will skip firms
whose portcos are already in the database (duplicate check on startup).
Just run it again — it picks up where it left off.

---

## CONFIG block (top of fsg_pipeline.py)

```python
API_KEY      = ""              # Or use env var (recommended)
MODEL        = "claude-sonnet-4-20250514"
DB_FILE      = "2026_04_24_CLAUDE_DATABASE_PRIVATE_EQUITY_RESEARCH_v8.xlsx"
FIRMS_CSV    = "firms_to_research.csv"
OUTPUT_DB    = "STEPHENS_FSG_DATABASE.xlsx"
DELAY_BETWEEN_FIRMS = 2        # seconds — increase if getting rate-limited
MAX_HTML_CHARS = 14000         # reduce to lower API costs
```
