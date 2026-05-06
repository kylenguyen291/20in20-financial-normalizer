# Vietnamese Financial Statement Normalizer

A Python pipeline that automatically downloads, extracts, and normalizes Vietnamese corporate financial statements (Báo cáo tài chính) from [CafeF](https://cafef.vn), then exports clean, analysis-ready Excel workbooks — including side-by-side company comparisons and auto-generated financial insights.

---

**API Key:** A demo API key is provided in [`api_key.txt`](./api_key.txt) in this repository.

---

## What It Does

1. **Download** — fetches annual report PDFs for any Vietnamese-listed ticker from CafeF
2. **Extract** — pulls raw text (digital PDFs) or renders page images (scanned PDFs)
3. **Normalize** — calls the Claude API to extract Income Statement, Balance Sheet, and Cash Flow into a structured JSON schema
4. **Export** — writes per-company Excel workbooks (9 sheets each: formatted, English-translated raw lines, and Vietnamese raw lines)
5. **Compare** — builds a multi-company comparison workbook with financial ratios, trends, and an auto-generated Insights sheet

---

## Requirements

- Python 3.11+
- An [Anthropic API key](https://console.anthropic.com/)

---

## Setup

### 1. Clone the repository

```bash
git clone https://github.com/<your-username>/financial-normalizer.git
cd financial-normalizer
```

### 2. Create and activate a virtual environment

```bash
python -m venv venv
source venv/bin/activate        # macOS / Linux
venv\Scripts\activate           # Windows
```

### 3. Install dependencies

```bash
pip install -r requirements.txt
```

### 4. Install Playwright browsers (required for PDF downloads)

```bash
playwright install chromium
```

### 5. Set up your API key

Copy the example environment file and add your key:

```bash
cp .env.example .env
```

Open `.env` and fill in:

```
ANTHROPIC_API_KEY=sk-ant-...
```

---

## Usage

Run the interactive CLI:

```bash
python main.py
```

You will be prompted to enter:
- How many years to extract (e.g. `3`)
- Each year (e.g. `2022`, `2023`, `2024`)
- Stock tickers one by one (e.g. `HPG`, `VHM`, `VIC`) — press Enter on a blank line when done

The pipeline then runs all five steps automatically and saves outputs to the `output/` folder.

### CLI flags

| Flag | Description |
|---|---|
| `--skip-download` | Use existing PDFs in `data/raw/` — skip the download step |
| `--skip-extract` | Use cached `.txt` files in `data/processed/` — skip extraction |
| `--skip-normalize` | Use cached `.json` files — re-export only |
| `--force` | Ignore all caches and re-run every step from scratch |
| `--use-sonnet` | Use Claude Sonnet instead of the default model |
| `--no-raw-lines` | Omit raw line items from JSON output (reduces API token usage) |

**Example — re-export without re-downloading or re-calling Claude:**

```bash
python main.py --skip-download --skip-extract --skip-normalize
```

**Example — force a full fresh run with Sonnet:**

```bash
python main.py --force --use-sonnet
```

---

## Output

All files are saved to the `output/` directory.

### Per-company workbook (`HPG_financials.xlsx`)

Each workbook contains 9 sheets:

| Sheet | Contents |
|---|---|
| `Income Statement` | Formatted IS with canonical English field names |
| `Balance Sheet` | Formatted BS with canonical English field names |
| `Cash Flow` | Formatted CF with canonical English field names |
| `EN Income Statement` | All raw line items with English translations |
| `EN Balance Sheet` | All raw line items with English translations |
| `EN Cash Flow` | All raw line items with English translations |
| `raw_Income Statement` | Raw Vietnamese line items as extracted |
| `raw_Balance Sheet` | Raw Vietnamese line items as extracted |
| `raw_Cash Flow` | Raw Vietnamese line items as extracted |

### Comparison workbook (`comparison.xlsx`)

| Sheet | Contents |
|---|---|
| `Insights` | Auto-generated English observations with ⚠ warnings and ✓ highlights |
| `Overview` | Key metrics for all companies × all years |
| `Profitability` | Gross/net/EBITDA margins, ROE, ROA |
| `Leverage & Solvency` | D/E ratio, Net Debt/EBITDA, interest coverage |
| `Liquidity` | Current ratio, quick ratio |
| `Efficiency` | Days receivable, days inventory, days payable |
| `Cash Flow Quality` | FCF, CFO/Net income, capex intensity |

All monetary values are in **million VND (triệu đồng)**.

---

## Project Structure

```
financial-normalizer/
├── main.py                  # CLI entry point
├── config.py                # Constants (model, paths, sheet names, colors)
├── requirements.txt
├── .env.example
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

## Supported Companies

Any Vietnamese-listed ticker on CafeF is supported. Three companies are pre-configured with display names and chart colors; all others are handled automatically:

| Ticker | Company |
|---|---|
| HPG | Hoa Phat Group |
| HSG | Hoa Sen Group |
| NKG | Nam Kim Steel |

To add a company permanently, edit the `COMPANIES` dict in `config.py`.

---

## Caching

The pipeline caches intermediate results so you only pay for what you need:

| Stage | Cache location | Skip flag |
|---|---|---|
| PDF download | `data/raw/<TICKER>_<YEAR>.pdf` | `--skip-download` |
| Text extraction | `data/processed/<TICKER>_<YEAR>.txt` | `--skip-extract` |
| Claude normalization | `data/processed/<TICKER>_<YEAR>.json` | `--skip-normalize` |

Use `--force` to bypass all caches and re-run everything.

---

## Notes

- **Scanned PDFs** (image-based) are handled automatically via Claude's vision API — no extra configuration needed.
- **Unit normalization** is applied automatically: raw VND values are converted to million VND, and billion VND values are scaled up.
- **Vietnamese section headers** are detected to skip auditor reports and accounting policy notes, sending only the financial tables to Claude.
- The pipeline sleeps 20 seconds between API calls to stay within rate limits.

---

## License

MIT
