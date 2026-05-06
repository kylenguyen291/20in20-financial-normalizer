# Vietnamese Financial Statement Normalizer

A pipeline that automatically downloads, extracts, and normalizes Vietnamese corporate financial statements (Báo cáo tài chính) from [CafeF](https://cafef.vn), then exports clean, analysis-ready Excel workbooks — including side-by-side company comparisons and auto-generated financial insights.

---

## Hosted App

The tool is deployed and accessible at:

> **[https://20in20-financial-normalizer-production-e7f7.up.railway.app](https://20in20-financial-normalizer-production-e7f7.up.railway.app)**

Enter stock tickers (e.g. `HPG`, `VIC`, `VHM`) and fiscal years (e.g. `2025, 2024, 2023`), then click **Run Pipeline**. When complete, download the Excel workbooks directly from the browser.

---

## What It Does

1. **Download** — fetches annual report PDFs for any Vietnamese-listed ticker from CafeF
2. **Extract** — pulls raw text (digital PDFs) or renders page images (scanned PDFs)
3. **Normalize** — calls the Claude API to extract Income Statement, Balance Sheet, and Cash Flow into structured JSON
4. **Export** — writes per-company Excel workbooks (6 sheets each: formatted + raw Vietnamese line items)
5. **Compare** — builds a multi-company comparison workbook with financial ratios, trends, and an auto-generated Insights sheet

---

## Output

All files are saved to the `output/` directory and downloadable from the web UI.

### Per-company workbook (`HPG_financials.xlsx`)

| Sheet | Contents |
|---|---|
| `Income Statement` | Formatted IS with canonical English field names |
| `Balance Sheet` | Formatted BS with canonical English field names |
| `Cash Flow` | Formatted CF with canonical English field names |
| `raw_Income Statement` | Raw Vietnamese line items as extracted |
| `raw_Balance Sheet` | Raw Vietnamese line items as extracted |
| `raw_Cash Flow` | Raw Vietnamese line items as extracted |

### Comparison workbook (`comparison.xlsx`)

| Sheet | Contents |
|---|---|
| `Insights` | Auto-generated English observations with ⚠ warnings and ✓ highlights |
| `Overview` | Key metrics for all companies × all years with color-scale formatting |
| `Profitability` | Gross/net/EBITDA margins, ROE, ROA |
| `Leverage` | D/E ratio, Net Debt/EBITDA, interest coverage |
| `Liquidity` | Current ratio, quick ratio |
| `Efficiency` | Days receivable, days inventory, asset turnover |
| `Cash Flow` | FCF, CFO/net income, capex intensity |
| `Key Financials` | Absolute figures in million VND |

All monetary values are in **million VND (triệu đồng)**.

---

## Local Development

### 1. Clone and install

```bash
git clone https://github.com/kylenguyen291/20in20-financial-normalizer.git
cd 20in20-financial-normalizer

python -m venv venv
source venv/bin/activate   # Windows: venv\Scripts\activate
pip install -r requirements.txt
playwright install chromium
```

### 2. Set your API key

```bash
cp .env.example .env
# Edit .env and add:  ANTHROPIC_API_KEY=sk-ant-...
```

### 3. Run the web app

```bash
python -m uvicorn app:app --port 8080
# Open http://localhost:8080
```

### 4. Or run the CLI directly

```bash
python main.py
```

CLI flags:

| Flag | Description |
|---|---|
| `--skip-download` | Use existing PDFs in `data/raw/` |
| `--skip-extract` | Use cached `.txt` files in `data/processed/` |
| `--skip-normalize` | Use cached `.json` files — re-export only |
| `--force` | Ignore all caches and re-run from scratch |

---

## Deployment (Railway)

The app runs as a single Docker container:

1. Push to GitHub
2. Create a new Railway project → **Deploy from GitHub repo**
3. Add environment variable: `ANTHROPIC_API_KEY=sk-ant-...`
4. Railway builds the `Dockerfile` and serves the app on a generated domain

---

## Project Structure

```
financial-normalizer/
├── app.py                   # FastAPI web app (serves UI + API on one port)
├── main.py                  # CLI entry point
├── config.py                # Constants (model, paths, sheet names)
├── requirements.txt
├── Dockerfile
├── railway.toml
│
├── static/
│   └── index.html           # Self-contained frontend (no build step)
│
├── src/
│   ├── downloader.py        # Downloads PDFs from CafeF via Playwright
│   ├── extractor.py         # Extracts text or images from PDFs
│   ├── normalizer.py        # Calls Claude API, parses + validates JSON
│   ├── exporter.py          # Writes per-company Excel workbooks
│   └── comparison_exporter.py  # Builds comparison workbook + insights
│
├── prompts/
│   └── extract_financials.txt  # System prompt for Claude
│
├── data/
│   ├── raw/                 # Downloaded PDFs (git-ignored)
│   └── processed/           # Cached .txt and .json files (git-ignored)
│
└── output/                  # Final Excel files (git-ignored)
```

---

## Notes

- **Scanned PDFs** are handled automatically via Claude's vision API — no configuration needed.
- **Unit normalization** converts raw VND → million VND and scales billion VND values.
- The pipeline sleeps 20 seconds between API calls to stay within Anthropic rate limits.
- Any Vietnamese-listed ticker on CafeF is supported.

---

## Running `main.py` Locally (CLI)

If you prefer to run the pipeline directly from the terminal without the web interface:

```bash
python main.py
```

You will be prompted interactively:

```
How many years? → 3
Year 1: 2023
Year 2: 2024
Year 3: 2025
Ticker (blank to finish): HPG
Ticker (blank to finish): VIC
Ticker (blank to finish):
```

The pipeline runs all five steps and saves Excel files to `output/`.

### CLI flags

| Flag | Description |
|---|---|
| `--skip-download` | Use existing PDFs in `data/raw/` — skip download |
| `--skip-extract` | Use cached `.txt` files in `data/processed/` — skip extraction |
| `--skip-normalize` | Use cached `.json` files — re-export only |
| `--force` | Ignore all caches and re-run every step from scratch |
| `--use-sonnet` | Force Claude Sonnet model |
| `--no-raw-lines` | Omit raw line items from JSON (reduces token usage) |

**Example — re-export without re-downloading or re-calling Claude:**

```bash
python main.py --skip-download --skip-extract --skip-normalize
```

**Example — force a full fresh run:**

```bash
python main.py --force --use-sonnet
```

---

## License

MIT
