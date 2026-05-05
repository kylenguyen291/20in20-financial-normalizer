from pathlib import Path

# ── Directories ───────────────────────────────────────────────────────────────
BASE_DIR       = Path(__file__).parent
DATA_RAW       = BASE_DIR / "data" / "raw"
DATA_PROCESSED = BASE_DIR / "data" / "processed"
OUTPUT_DIR     = BASE_DIR / "output"
PROMPTS_DIR    = BASE_DIR / "prompts"

for d in [DATA_RAW, DATA_PROCESSED, OUTPUT_DIR]:
    d.mkdir(parents=True, exist_ok=True)

# ── Known companies (name + color for Excel formatting) ───────────────────────
COMPANIES = {
    "HPG": {"name": "Hoa Phat Group",  "color": "4472C4"},   # blue
    "HSG": {"name": "Hoa Sen Group",   "color": "70AD47"},   # green
    "NKG": {"name": "Nam Kim Steel",   "color": "ED7D31"},   # orange
}

_FALLBACK_COLORS = [
    "5B9BD5", "FFC000", "A9D18E", "FF7C80",
    "9E480E", "7030A0", "00B0F0", "92D050",
]

def get_company_info(ticker: str) -> dict:
    if ticker in COMPANIES:
        return COMPANIES[ticker]
    color = _FALLBACK_COLORS[hash(ticker) % len(_FALLBACK_COLORS)]
    return {"name": ticker, "color": color}


# ── Default years ─────────────────────────────────────────────────────────────
DEFAULT_YEARS = [2022, 2023, 2024]

# ── CafeF URL ─────────────────────────────────────────────────────────────────
CAFEF_URL_TEMPLATE = "https://cafef.vn/du-lieu/upcom/{ticker}-bao-cao-tai-chinh.chn"

# ── Claude ────────────────────────────────────────────────────────────────────
CLAUDE_MODEL        = "claude-sonnet-4-6"           # default: high accuracy
CLAUDE_MODEL_SONNET = "claude-sonnet-4-6"           # alias kept for compatibility
CLAUDE_MAX_TOKENS   = 8192
CLAUDE_TEMPERATURE  = 0

# ── HTTP headers ──────────────────────────────────────────────────────────────
HTTP_HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "vi-VN,vi;q=0.9,en-US;q=0.8",
}

# ── Excel sheet names ─────────────────────────────────────────────────────────
SHEET_NAMES = {
    "income_statement": "Income Statement",
    "balance_sheet":    "Balance Sheet",
    "cash_flow":        "Cash Flow",
}

RAW_SHEET_NAMES = {
    "income_statement": "raw_Income Statement",
    "balance_sheet":    "raw_Balance Sheet",
    "cash_flow":        "raw_Cash Flow",
}

EN_RAW_SHEET_NAMES = {
    "income_statement": "EN Income Statement",
    "balance_sheet":    "EN Balance Sheet",
    "cash_flow":        "EN Cash Flow",
}

# ── Output unit ──────────────────────────────────────────────────────────────
# All monetary values are expressed in million VND (triệu đồng).
# This matches the native unit of most Vietnamese financial statements and
# keeps all numbers as whole integers — no decimal places needed.
OUTPUT_UNIT = "million VND"

# ── Excel number format ───────────────────────────────────────────────────────
# Integer format — full precision to the ones place, thousand-separator only.
NUMBER_FORMAT = '#,##0'
MISSING_COLOR = "FFFF00"   # yellow fill for N/A cells
