# FSG PE Database — API Loop Script

Runs the full research pipeline for every firm in your CSV, unattended.
No response length limits. No "continue" prompts. Saves after every firm.

---

## Setup (one time, ~2 minutes)

**1. Install Python packages**
```bash
pip install anthropic firecrawl-py openpyxl rapidfuzz
```

**2. Get API keys**
- Anthropic: console.anthropic.com → API Keys → Create Key
- Firecrawl: https://www.firecrawl.dev/ → Dashboard → API Keys

**3. Set both keys as environment variables**
```bash
# Mac / Linux
export ANTHROPIC_API_KEY=sk-ant-api03-...
export FIRECRAWL_API_KEY=fc-...

# Windows PowerShell
$env:ANTHROPIC_API_KEY = "sk-ant-api03-..."
$env:FIRECRAWL_API_KEY = "fc-..."
```

**4. Place these files in one folder:**
- `fsg_pipeline.py`          — this script
- `firms_to_research.csv`    — your firm list
- `2026_04_24_CLAUDE_DATABASE_PRIVATE_EQUITY_RESEARCH_v3.xlsx` — your database

---

## firms_to_research.csv format

```
firm_name,website,portfolio_url_suffix
Rotunda Capital Partners,https://rotundacapital.com,/portfolio
ICV Partners,https://icvpartners.com,/portfolio
Pfingsten Partners,https://pfingsten.com,
```

- `portfolio_url_suffix` is optional — script auto-detects if blank
- To run all 472 firms: export from your database's Private Equity Firms tab

---

## Run it

```bash
python3 fsg_pipeline.py
```

Walk away. The script:
1. Reads each firm from the CSV
2. Locates the portfolio page (CSV suffix → Firecrawl `map` → fallback hints)
3. Scrapes it via Firecrawl as clean markdown (handles JS-rendered sites)
4. Extracts portcos via Claude API
5. Enriches missing fields (entry date, advisor, deal team)
6. Validates every row (7 hard rules + 6 soft rules)
7. Writes valid rows to `STEPHENS_FSG_DATABASE.xlsx`
8. **Saves after every firm** — safe to interrupt and restart
9. Logs results to `pipeline_log.csv`

Firecrawl scrape responses are cached for 7 days, so re-runs after an
interruption mostly hit cache and don't re-bill.

---

## Output files

| File | Contents |
|---|---|
| `STEPHENS_FSG_DATABASE.xlsx` | Updated database (never overwrites your original) |
| `pipeline_log.csv` | Per-firm: portcos found / written / failed |
| `pipeline_errors.log` | Hard validation failures and fetch errors |

---

## Cost estimate

- **Anthropic:** ~$3–6 for all 472 firms at claude-sonnet-4 pricing
  (~$3/MTok). Roughly $0.006–0.012 per firm.
- **Firecrawl:** ~1 scrape per firm (portfolio page) plus ~1 map call.
  At ~1 credit per scrape on the Hobby plan, 472 firms ≈ ~500–1,500 credits.
  Cached re-runs (within 7 days) don't re-bill.

## Speed estimate

~2–4 minutes per firm (Firecrawl scrape + 2 Claude API calls + write).
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
API_KEY            = ""        # Anthropic — or use env var (recommended)
FIRECRAWL_API_KEY  = ""        # Firecrawl — or use env var (recommended)
MODEL              = "claude-sonnet-4-20250514"
DB_FILE            = "2026_04_24_CLAUDE_DATABASE_PRIVATE_EQUITY_RESEARCH_v3.xlsx"
FIRMS_CSV          = "firms_to_research.csv"
OUTPUT_DB          = "STEPHENS_FSG_DATABASE.xlsx"
DELAY_BETWEEN_FIRMS = 2        # seconds — increase if getting rate-limited
MAX_MD_CHARS        = 30000    # markdown truncation; reduce to lower API costs
FIRECRAWL_CACHE_MS  = 7 * 24 * 60 * 60 * 1000  # reuse scrapes for 7 days
```
