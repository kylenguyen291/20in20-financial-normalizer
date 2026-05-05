# Edge Cases & Fixes Log

> Running record of every non-obvious bug, data quirk, or pipeline failure encountered — and exactly how it was resolved. Update this file whenever a new edge case is handled.

---

## EC-001 · CafeF page is JavaScript-rendered

**Symptom:** `requests + BeautifulSoup` returned an HTML skeleton with 0 PDF links, even though the page clearly shows a table of reports in the browser.

**Root cause:** CafeF's financial report table is loaded via AJAX after the initial page load. The static HTML source contains only the page shell — no table rows, no links.

**Fix:** Replaced `requests + BeautifulSoup` with `Playwright` (headless Chromium). Added `wait_until="networkidle"` + `page.wait_for_timeout(2000)` to ensure the AJAX table fully renders before querying the DOM.

**Files changed:** `src/downloader.py`

---

## EC-002 · URL-based PDF filter matched 0 files

**Symptom:** Playwright found 183 PDF links on the CafeF page, but the URL keyword filter (`"hop_nhat"`, `"kiem_toan"`, `"nam_202X"`) matched none of them.

**Root cause:** CafeF's PDF filenames are inconsistent across years and companies. Some use `sau_kiem_toan` instead of `kiem_toan`, and the year encoding in the URL varies (e.g., `_2022`, `nam2022`, `202200`). URL-based filtering is not reliable.

**Fix:** Switched to Vietnamese row text filtering. For each `<a>` tag, we call `link.evaluate("el => el.closest('tr').innerText")` to read the visible row label, then apply `_is_target_row()`:
```python
def _is_target_row(row_text: str) -> tuple[bool, int | None]:
    text = row_text.lower()
    if 'hợp nhất' not in text: return False, None   # must be consolidated
    if 'đã kiểm toán' not in text: return False, None  # must be audited
    if 'quý' in text: return False, None             # exclude quarterly
    for year in TARGET_YEARS:
        if str(year) in row_text: return True, year
    return False, None
```
Vietnamese row labels are stable across years and companies — far more reliable than URL patterns.

**Files changed:** `src/downloader.py`

---

## EC-003 · All PDFs are scanned images — pdfplumber returns empty

**Symptom:** `extractor.py` logged "All N pages empty — scanned PDF detected" for every PDF. pdfplumber returned 0 text.

**Root cause:** All 9 financial statement PDFs for HPG, HSG, NKG are scanned image PDFs (not digitally typeset). pdfplumber can only extract text from digital PDFs.

**Fix:** Added a PyMuPDF (`fitz`) fallback in `_pdf_to_images()`. When pdfplumber returns empty pages, the extractor converts the first 30 pages to PNG images at 120 DPI and returns them as base64-encoded Anthropic vision content dicts. The normalizer then calls Claude vision API instead of text API.

**Files changed:** `src/extractor.py`, `src/normalizer.py`

---

## EC-004 · Mixed PDF — pdfplumber extracts cover page only

**Symptom:** HSG_2023 and NKG_2024 — pdfplumber extracted 135–146 characters (just the cover page title). The normalizer passed this garbage text to Claude's text API. Claude returned an empty string. JSON parse failed with: `"Expecting value: line 1 column 1 (char 0)"`.

**Root cause:** Some PDFs have a digitally typeset cover page followed by scanned image pages for the actual financial tables. pdfplumber picks up the cover text and reports "success", but the content is useless.

**Fix:** Added `MIN_CHARS = 1500` threshold in `extractor.py`. If pdfplumber extracts text but the total length is below 1500 characters, the extractor treats the file as scanned and falls back to PyMuPDF vision mode.

```python
MIN_CHARS = 1500
if full_text.strip() and len(full_text.strip()) >= MIN_CHARS:
    cache_path.write_text(full_text, encoding="utf-8")
    return full_text
else:
    print(f"  ⚠ Text too short ({len(full_text.strip())} chars) — mixed/partial scan")
    return _pdf_to_images(pdf_path)
```

**Side effect to watch:** If a bad `.txt` cache file (< 1500 chars) was written by a previous run, it must be manually deleted before re-running — or use `--force` flag. The extractor returns cached `.txt` without re-checking length.

**Files changed:** `src/extractor.py`

---

## EC-005 · Retry loop silently exits on JSON parse failure

**Symptom:** Normalizer logged `"✗ JSON parse error: Expecting value: line 1 column 1 (char 0)"` but did not retry — it moved on immediately as if there were no more attempts left.

**Root cause:** `_parse_response()` catches `JSONDecodeError` internally and returns `None`. The retry loop only caught `anthropic.APIError` exceptions. A `None` return from a parse failure was treated as a successful (but empty) result, causing the loop to exit.

**Fix:** Added explicit `None` check in both `_call_text` and `_call_vision`:
```python
result = _parse_response(resp.content[0].text)
if result is not None:
    return result          # genuine success
# result is None → parse failed → fall through to retry
if attempt <= retries:
    time.sleep(3)
```

**Files changed:** `src/normalizer.py`

---

## EC-006 · Claude returns raw VND instead of billion VND

**Symptom:** 8 out of 9 normalized JSON files had `"unit": "VND"` and values like `141,409,274,460,632` instead of `141,409` (billion VND). The Excel output showed wildly inconsistent columns — 2022 and 2024 in raw VND, 2023 in billion VND.

**Root cause:** Vietnamese financial PDFs declare their unit in the document header (e.g., "Đơn vị tính: đồng"). The original Claude prompt said "detect the unit and return values in that unit" — so Claude correctly read the PDF unit and returned raw VND, rather than converting to billions.

**Fix — Layer 1 (prompt):** Updated `prompts/extract_financials.txt` to explicitly require billion VND output regardless of PDF unit, with a conversion table and worked example:
```
- "tỷ đồng"    → already billion VND → use as-is
- "triệu đồng" → divide by 1,000
- "đồng" / raw → divide by 1,000,000,000
Always set "unit": "billion VND". Round to nearest integer.
```

**Fix — Layer 2 (code safety net):** Added `_normalize_units()` in `normalizer.py` that runs on every Claude response before saving. It detects raw VND (any value > 10^12) or million VND (any value > 10^9) and auto-converts, logging a warning when it fires:
```python
if probe > 1_000_000 * 1_000_000:    # raw VND
    divisor = 1_000_000_000
elif probe > 1_000_000 * 1_000:      # million VND
    divisor = 1_000
```

**One-time data fix:** Applied a conversion script to all 8 affected JSON files in `data/processed/`. Re-exported all 3 Excel workbooks.

**Files changed:** `prompts/extract_financials.txt`, `src/normalizer.py`

---

## EC-007 · API key accidentally hardcoded

**Symptom:** `normalizer.py` contained `os.environ.get("sk-ant-...")` — the actual API key was passed as the environment variable name instead of `"ANTHROPIC_API_KEY"`.

**Root cause:** Key was pasted into chat during debugging and written directly into the source file.

**Fix:** Changed to `os.environ.get("ANTHROPIC_API_KEY")` + key stored in `.env` file loaded via `python-dotenv`. Revoke the exposed key in Anthropic console immediately and rotate.

**Files changed:** `src/normalizer.py`, `.env`

---

## EC-008 · Mixed PDF with sufficient text — scanned financial table pages silently dropped

**Symptom:** A PDF containing both digital text pages (e.g., auditor's report, notes) and scanned image pages (the actual Income Statement / Balance Sheet / Cash Flow tables) passed the `MIN_CHARS = 1500` threshold because the text pages alone produced enough characters. The extractor returned text mode, and the normalizer received only the text pages. The financial tables — being on the scanned pages — were never sent to Claude, resulting in a JSON with empty or partially-filled sections.

**Root cause:** The `extract()` function checked only the *total* extracted character count. If that total exceeded 1500 chars, it always returned text mode — regardless of how many pages pdfplumber silently returned 0 characters for. Empty pages were counted but their count had no effect on the routing decision.

**Fix:** Added per-page character tracking (`page_char_counts`). After extraction, counts the number of "sparse" pages (< 50 chars per page). If sparse pages exceed **either** of two thresholds:
- Absolute: > 3 sparse pages
- Relative: > 15% of total pages are sparse

…the PDF is classified as **mixed** and routed to full vision mode regardless of total text length. This ensures Claude receives all pages (including the scanned financial table pages) as images.

```python
PAGE_MIN_CHARS        = 50    # per-page threshold: below this = effectively scanned
MIXED_ABS_THRESHOLD   = 3     # absolute: >3 sparse pages triggers vision mode
MIXED_RATIO_THRESHOLD = 0.15  # relative: >15% sparse pages triggers vision mode

is_mixed = (
    sparse_pages > MIXED_ABS_THRESHOLD
    or (total_pages > 0 and sparse_pages / total_pages > MIXED_RATIO_THRESHOLD)
)
```

**Side effect:** Mixed PDFs that previously returned partial text results will now always re-run in vision mode. If a stale `.txt` cache exists from a pre-fix run, delete it or use `--force` before re-running.

**Files changed:** `src/extractor.py`, `EDGE_CASES.md`

---

## EC-009 · DRM / eoffice watermark PDF — repetitive text passes MIN_CHARS but is junk

**Symptom:** BSR_2022 (and similar eoffice-hosted PDFs) — pdfplumber extracts ~4,773 chars across 46 pages, passes the MIN_CHARS threshold, and the file is cached as `BSR_2022.txt`. But every single page contains only the same eoffice watermark line: `"Văn bản được tải lên hệ thống eoffice.bsr.com.vn. Với số định danh: 287/CV-VPHĐQT/2023"`. Claude receives 46 copies of the same sentence, finds no financial data, and returns an empty string. JSON parse fails on all 3 attempts.

**Root cause:** The eoffice document management system stamps a DRM watermark as selectable text on every page. The actual financial content underneath is rendered as an image (invisible to pdfplumber). This PDF appears digital to MIN_CHARS but is semantically scanned.

**Fix:** Added `_is_drm_watermark(text)` in `extractor.py`. It strips page-header lines (`--- Page N ---`), then computes the ratio of unique content lines to total content lines. If fewer than 10% of lines are unique, the content is flagged as repetitive junk. This check runs in two places:

1. **On cache read** — if the cached `.txt` is DRM junk, the cache file is deleted and vision mode is used instead (no `--force` needed).
2. **On fresh extraction** — checked after pdfplumber runs, before writing the cache.

```python
def _is_drm_watermark(text: str, uniqueness_threshold: float = 0.10) -> bool:
    lines = [l.strip() for l in text.splitlines()
             if l.strip() and not l.startswith("--- Page")]
    if not lines:
        return False
    return len(set(lines)) / len(lines) < uniqueness_threshold
```

On next run for BSR_2022:
```
⚠ Cached text is DRM/watermark junk — discarding cache, switching to vision
  Converting 30/46 pages to images (DPI=120)...
```

**Files changed:** `src/extractor.py`, `EDGE_CASES.md`

---

## EC-010 · Blind text truncation cuts off financial tables (NVL_2023 empty JSON)

**Symptom:** `NVL_2023` — pdfplumber extracted ~213,000 characters (digital PDF). The normalizer truncated to the first 100,000 chars and sent that to Claude. Claude returned an empty JSON with all `null` fields. The log showed no section headers found in the extract window.

**Root cause:** A typical 200k-char Vietnamese financial statement PDF is structured as:
```
[auditor report ~20k] [accounting policies ~60k] [IS/BS/CF tables ~40k] [notes ~80k]
```
Blind truncation at 100k chars captured only the auditor boilerplate and policy notes — cutting off entirely before the actual financial tables. Claude received the wrong section of the document.

**Fix:** Added `_extract_financial_sections()` in `normalizer.py`. Instead of truncating from byte 0, it scans for 27 known Vietnamese financial statement section headers (income statement, balance sheet, cash flow, and equity statement variants) and jumps directly to the earliest match. It then takes a 60,000-char window starting from that position.

```python
SECTION_HEADERS = [
    "báo cáo kết quả hoạt động kinh doanh hợp nhất",
    "bảng cân đối kế toán hợp nhất",
    "báo cáo lưu chuyển tiền tệ hợp nhất",
    ... # 27 variants total
]
start = min(positions)   # earliest header found
return text[start : start + 60_000]
```

If no headers are found, falls back to the first 60,000 chars and logs a warning.

**Files changed:** `src/normalizer.py`

---

## EC-011 · CafeF CDN hostname is unpredictable (404 on old PDFs)

**Symptom:** PDF download returned HTTP 404 for certain tickers (e.g. NVL, VHM older years). The URL used `cafefnew.mediacdn.vn` but the file was actually hosted on `cafef1.mediacdn.vn`, or vice versa.

**Root cause:** Initial assumption was that CDN assignment followed a year-based rule (pre-2022 → `cafef1`, 2022+ → `cafefnew`). This turned out to be wrong — CafeF assigns files to either CDN host unpredictably regardless of year.

**Fix:** Removed the year-based `_fix_cdn_url()` logic entirely. Replaced with `_cdn_fallbacks()` which tries all known CDN variants on 404:

```python
_CDN_VARIANTS = ["cafefnew.mediacdn.vn", "cafef1.mediacdn.vn"]

def _cdn_fallbacks(url: str) -> list[str]:
    variants = []
    for cdn in _CDN_VARIANTS:
        candidate = re.sub(r"cafef\w+\.mediacdn\.vn", cdn, url)
        if candidate not in variants:
            variants.append(candidate)
    return variants
```

`_download_pdf()` now iterates through all variants on each 404, logging which CDN host eventually succeeds.

**Files changed:** `src/downloader.py`

---

## EC-012 · Rate limit (429) causes silent failure after inadequate retry sleep

**Symptom:** Normalizer hit Anthropic's rate limit mid-run. The 5-second sleep between retries was far shorter than the 60-second token-per-minute window. On attempt 2, the rate limit fired again immediately. All 3 attempts exhausted, file skipped.

**Root cause:** Claude API enforces a tokens-per-minute (TPM) limit. Vision calls in particular consume 20,000–60,000 tokens each. A 5-second sleep does not come close to clearing the TPM window.

**Fix (two layers):**

1. **`anthropic.RateLimitError`-specific backoff** in both `_call_text` and `_call_vision`: sleep `60 × attempt` seconds (60s, 120s, 180s) instead of the generic 5s, giving the TPM window time to reset.

```python
except anthropic.RateLimitError:
    wait = 60 * attempt
    print(f"    ✗ Rate limit (attempt {attempt}) — waiting {wait}s...")
    if attempt <= retries:
        time.sleep(wait)
```

2. **Inter-call gap** in `run()` increased from 1s to 20s between every company/year normalization job, preventing the next call from arriving before the previous minute's window has cleared.

**Files changed:** `src/normalizer.py`

---

## EC-013 · Claude appends explanation text after JSON — "Extra data" parse error

**Symptom:** `VHM_2022` — Claude returned valid JSON immediately followed by a plain-text explanation paragraph. `json.loads()` raised `JSONDecodeError: Extra data: line 1 column 1316 (char 1315)`. All 3 attempts failed with the same error.

**Root cause:** Despite the prompt instruction "return ONLY valid JSON", Claude occasionally appends a sentence like "Note: some fields could not be extracted because..." after the closing brace. `json.loads()` requires the entire string to be valid JSON and rejects any trailing content.

**Fix:** Replaced `json.loads(raw)` with `json.JSONDecoder().raw_decode(raw)` in `_parse_response()`. `raw_decode` parses the first complete JSON object it finds and stops, returning the object and the position of the trailing content. The trailing text is silently discarded.

```python
obj, _ = json.JSONDecoder().raw_decode(raw)
return obj
```

**Files changed:** `src/normalizer.py`

---

## EC-014 · Numbers rounded to integers — losing decimal precision

**Symptom:** Exported Excel values like `9923` billion VND instead of `9922.941` for a PDF value of `9.922.941.127.284 đồng`. Exact figures from the financial statement were not preserved.

**Root cause:** The original prompt instructed Claude to "round to nearest integer." The unit-safety-net in `_normalize_units()` also used `round(v / divisor)` (integer round).

**Fix:**

1. **Prompt** (`prompts/extract_financials.txt`): Changed rounding instruction to "Return exact values — do NOT round to integers. Preserve up to 3 decimal places after conversion."

2. **Normalizer** (`src/normalizer.py`): Changed `round(v / divisor)` to `round(v / divisor, 3)` in `_normalize_units()`.

3. **Config** (`config.py`): Changed Excel number format from `'#,##0'` to `'#,##0.###'` (trailing zeros suppressed, up to 3 decimal places shown).

**Files changed:** `prompts/extract_financials.txt`, `src/normalizer.py`, `config.py`

---

## EC-015 · Raw sheets show canonical English labels instead of Vietnamese line items

**Symptom:** The `raw_Balance Sheet` / `raw_Income Statement` / `raw_Cash Flow` sheets displayed the same ~30 aggregated English fields as the formatted sheets (e.g. "Cash & Cash Equivalents") instead of every individual line item from the original Vietnamese document (e.g. "311 · Phải trả người bán ngắn hạn").

**Root cause (design gap):** The raw sheet was populated by iterating over canonical JSON keys — the same fields the formatted sheet uses. It had no access to the granular line items because Claude was never asked to extract them.

**Fix (two-part):**

**Part 1 — Prompt:** Added a `raw_lines` array to each section's JSON schema. Claude now extracts every line item it sees in the document, in document order, with four fields per entry:
- `label` — exact Vietnamese text as printed
- `code` — mã số reference number (e.g. `"311"`)
- `value` — numeric value in billion VND
- `type` — `"total"` / `"subtotal"` / `"item"` (mirrors the document's bold/indent hierarchy)

**Part 2 — Exporter:** Rewrote `_write_raw_sheet()`. When `raw_lines` is present in the JSON, the raw sheet now renders three column groups: **Code** | **Line Item (Vietnamese)** | **year columns**. Row formatting follows document hierarchy — total rows get dark blue fill + thick border, subtotals get light blue, items alternate white/grey. A `_build_raw_master()` helper merges `raw_lines` across multiple years, matching rows by `code` first, then by `label`, so year columns align correctly even when one year has extra line items.

Falls back gracefully to the old canonical-field display for JSON files cached before this change (no `raw_lines` key present).

**Also:** `CLAUDE_MAX_TOKENS` raised from 4096 → 8192 to accommodate the larger response. `_normalize_units()` updated to also convert values inside `raw_lines` entries when a unit mismatch is detected.

**Note:** Existing cached `.json` files do not contain `raw_lines`. Re-run with `--skip-download --skip-extract` (or delete `data/processed/*.json`) to regenerate with the new prompt.

**Files changed:** `prompts/extract_financials.txt`, `src/exporter.py`, `src/normalizer.py`, `config.py`

---

## EC-016 · `language` parameter ghost — `TypeError: run() got an unexpected keyword argument 'language'`

**Symptom:** Step 4 (export) crashed immediately:
```
TypeError: run() got an unexpected keyword argument 'language'
```

**Root cause:** In a previous session, the Vietnamese language option was removed from `exporter.run()`. However, `main.py` still had three surviving references: the `_ask_language()` function definition, its call in `main()`, and `language=language` in the `export()` call. The export function signature no longer accepted the argument.

**Fix:** Removed all three references from `main.py`:
- Deleted `_ask_language()` function
- Removed `language = _ask_language()` from `main()`
- Removed `language=language` from the `export()` call
- Removed the `Language` line from the startup summary print block

**Files changed:** `main.py`

---

## EC-017 · Output unit changed from billion VND to million VND — full integer precision

**Symptom / Motivation:** Extracted values like `9,922.941` (billion VND) contained decimal places that were artifacts of the unit conversion, not present in the original document. The source PDF declared "triệu đồng" and showed `9,922,941` as a whole integer. The pipeline was introducing false precision by converting to billions and then rounding to 3 decimal places.

**Root cause:** The pipeline standardised on billion VND as the output unit. Most Vietnamese financial statements natively use million VND (triệu đồng), so converting *down* to billions forced fractional values. Additionally, `NUMBER_FORMAT = '#,##0.###'` allowed decimal places in Excel, making numbers inconsistent across companies that declared different source units.

**Fix (four files):**

1. **`prompts/extract_financials.txt`** — Unit instruction changed from "always output billion VND" to "always output million VND (triệu đồng)". Conversion table updated:
   - `triệu đồng` → already in million VND, use as-is (no conversion)
   - `tỷ đồng` → multiply by 1,000
   - `đồng` / raw VND → divide by 1,000,000
   - All output values must be **whole integers — no decimal places**.

2. **`config.py`** — `NUMBER_FORMAT` changed from `'#,##0.###'` to `'#,##0'` (integer, comma separator only). Added `OUTPUT_UNIT = "million VND"` as a named constant.

3. **`src/normalizer.py`** — `_normalize_units()` safety net rebuilt for million VND target:
   - `probe > 10^12` → raw VND → divide by 1,000,000, `round()` to integer
   - `probe < 10^6`  → billion VND → multiply by 1,000, `round()` to integer
   - `10^6 ≤ probe ≤ 10^12` → already million VND, no action
   - `_validate()` default unit changed from `"billion VND"` to `"million VND"`
   - All conversions use `round()` with no decimal argument → result is always a whole integer

4. **`src/exporter.py`** — Banner label updated from `"Unit: Billion VND"` to `"Unit: Million VND"` across all 6 sheet types (3 formatted + 3 raw).

**Side effect:** All existing cached `.json` files contain old billion VND values with decimals. Must re-run with `--skip-download --skip-extract` (or delete `data/processed/*.json`) to regenerate clean integer values in million VND before exporting.

**Files changed:** `prompts/extract_financials.txt`, `config.py`, `src/normalizer.py`, `src/exporter.py`

---

## EC-018 · Banking / financial institution income statement — all canonical fields N/A

**Symptom:** TCB (Techcombank) 2023 — the formatted Income Statement sheet showed N/A for every field except Profit Before Tax, Corporate Income Tax, Net Income, and the parent/minority split. The raw sheet, however, was complete with all Vietnamese line items and correct values.

**Root cause (design gap):** The canonical income statement schema was designed for industrial and commercial companies. It maps fields like `net_revenue`, `cost_of_goods_sold`, `gross_profit`, `selling_expenses`, and `operating_profit` — none of which exist in a bank's income statement.

Banks and financial institutions use a fundamentally different structure:
```
Thu nhập lãi thuần           (Net Interest Income)
Lãi thuần từ hoạt động dịch vụ  (Net Service Fee Income)
Chi phí hoạt động            (Operating Expenses)
Chi phí dự phòng rủi ro tín dụng (Credit Loss Provision)
Tổng lợi nhuận trước thuế    (Profit Before Tax)
```
Claude cannot map these to the canonical fields and returns `null` for all of them. Only bottom-line fields (profit before tax, tax, net income) exist in both banking and non-banking statements, so those three map correctly.

The raw sheet is unaffected because `raw_lines` captures every line item regardless of canonical mapping.

**Status:** Known issue — not yet fixed. The raw sheet provides full data for banks; the formatted sheet is mostly N/A.

**Planned fix:** Detect financial institution tickers (by ticker lookup table or by presence of banking-specific Vietnamese headers in the extracted text such as "thu nhập lãi" or "chi phí dự phòng rủi ro tín dụng") and route them to a separate `banking_income_statement` canonical schema with the correct fields.

**Affected tickers (examples):** TCB, VCB, BID, CTG, MBB, ACB, VPB, HDB, TPB, STB and any other listed bank, insurance company, or securities firm.

**Files to change (when implemented):** `prompts/extract_financials.txt`, `src/exporter.py`, `config.py`

---

## Template for new edge cases

```
## EC-XXX · Short title

**Symptom:** What went wrong / what the user saw.

**Root cause:** Why it happened.

**Fix:** What was changed and how.

**Files changed:** list of files
```
