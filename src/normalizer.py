"""
normalizer.py
─────────────
Calls Claude API to extract + normalize financial data.

Handles two input types from extractor.py:
  - str:        text from digital PDF  → text-based Claude call
  - list[dict]: images from scanned PDF → vision-based Claude call

Both paths return the same structured JSON schema.
"""

import json
import os
import time
from pathlib import Path

import anthropic
from dotenv import load_dotenv

from config import (DATA_PROCESSED, PROMPTS_DIR,
                    CLAUDE_MODEL, CLAUDE_MODEL_SONNET,
                    CLAUDE_MAX_TOKENS, CLAUDE_TEMPERATURE)

load_dotenv()

_SYSTEM_PROMPT_FULL  = (PROMPTS_DIR / "extract_financials.txt").read_text(encoding="utf-8")
_REQUIRED_KEYS       = {"income_statement", "balance_sheet", "cash_flow"}

# Appended to the user message when --no-raw-lines is set.
# Telling the model in the user turn overrides the system prompt's raw_lines
# instructions without us having to maintain two versions of the prompt file.
_NO_RAW_LINES_NOTE = (
    "\n\nIMPORTANT: Do NOT include raw_lines arrays in your response. "
    "Omit the raw_lines field from all three sections entirely. "
    "Return only the canonical named fields (net_revenue, gross_profit, etc.)."
)


def _parse_response(raw: str) -> dict | None:
    """Parse Claude's response, stripping markdown fences if present."""
    raw = raw.strip()
    if raw.startswith("```"):
        parts = raw.split("```")
        raw = parts[1]
        if raw.startswith("json"):
            raw = raw[4:]
        raw = raw.strip()
    try:
        # Use raw_decode so trailing text after the JSON object is ignored.
        # json.loads() raises "Extra data" when Claude appends an explanation
        # after the closing brace; raw_decode stops at the first complete object.
        obj, _ = json.JSONDecoder().raw_decode(raw)
        return obj
    except json.JSONDecodeError as e:
        print(f"    ✗ JSON parse error: {e}")
        return None


def _extract_financial_sections(text: str, window: int = 60_000) -> str:
    """
    Locate the three Vietnamese financial statement sections and return only
    that slice of the document, rather than blindly taking the first N chars.

    Strategy:
      1. Search for known Vietnamese section headers (IS / BS / CF).
      2. Use the *earliest* match as the start position — this is where the
         financial tables begin, past any cover page / auditor report / TOC.
      3. Return up to `window` chars from that start position.
      4. If no headers are found, fall back to the first `window` chars and
         log a warning so the caller knows the heuristic didn't fire.

    Why this matters:
      A 200k-char digital PDF is typically structured as:
        [auditor report ~20k] [accounting policies ~60k] [IS/BS/CF ~40k] [notes ~80k]
      Blind truncation at 100k sends Claude the audit + policies boilerplate and
      cuts off before the actual tables, causing an empty JSON response.
      Section-aware extraction jumps directly to the tables.
    """
    # All known Vietnamese header variants for the five financial statements.
    # Listed longest/most-specific first so they match before shorter substrings.
    SECTION_HEADERS = [
        # ── Income Statement ─────────────────────────────────────────────────
        "báo cáo kết quả hoạt động kinh doanh hợp nhất",
        "báo cáo kết quả hoạt động kinh doanh",
        "kết quả hoạt động kinh doanh hợp nhất",
        "kết quả hoạt động kinh doanh",
        "kết quả kinh doanh hợp nhất",
        "kết quả kinh doanh",
        "báo cáo thu nhập toàn diện hợp nhất",   # comprehensive income (IFRS)
        "báo cáo thu nhập toàn diện",
        # ── Balance Sheet ────────────────────────────────────────────────────
        "bảng cân đối kế toán hợp nhất",
        "bảng cân đối kế toán",
        "cân đối kế toán hợp nhất",
        "cân đối kế toán",
        "bảng cân đối tài sản",                  # older / less common term
        # ── Cash Flow ────────────────────────────────────────────────────────
        "báo cáo lưu chuyển tiền tệ hợp nhất",
        "báo cáo lưu chuyển tiền tệ",
        "lưu chuyển tiền tệ hợp nhất",
        "lưu chuyển tiền tệ",
        "lưu chuyển tiền hợp nhất",
        "lưu chuyển tiền",
        # ── Statement of Changes in Equity ───────────────────────────────────
        "báo cáo thay đổi vốn chủ sở hữu hợp nhất",
        "báo cáo thay đổi vốn chủ sở hữu",
        "báo cáo biến động vốn chủ sở hữu hợp nhất",
        "báo cáo biến động vốn chủ sở hữu",
        "thay đổi vốn chủ sở hữu hợp nhất",
        "thay đổi vốn chủ sở hữu",
        "biến động vốn chủ sở hữu",
    ]

    text_lower = text.lower()
    positions = []

    for header in SECTION_HEADERS:
        idx = text_lower.find(header)
        if idx != -1:
            positions.append(idx)

    if not positions:
        print(f"    ⚠ No Vietnamese section headers found — using first {window:,} chars")
        return text[:window]

    start = min(positions)
    end   = min(start + window, len(text))

    found_section = text[start : start + 60].replace("\n", " ").strip()
    print(
        f"    ✂ Section-aware extraction: start={start:,} chars in, "
        f"window={end - start:,} chars"
    )
    print(f"      First header: \"{found_section[:55]}...\"")

    return text[start:end]


def _call_text(client, text: str, ticker: str, year: int, retries: int = 4,
               model: str = CLAUDE_MODEL, include_raw_lines: bool = True) -> dict | None:
    """Claude call for digital PDFs (text input)."""
    SECTION_WINDOW = 60_000   # chars to send after the first financial section header
    FALLBACK_LIMIT = 60_000   # chars to send when no section headers are found

    if len(text) > FALLBACK_LIMIT:
        text = _extract_financial_sections(text, window=SECTION_WINDOW)
    else:
        print(f"    ✓ Text fits within limit ({len(text):,} chars) — no truncation needed")

    user_message = (
        f"Company ticker: {ticker}\n"
        f"Target year: {year}\n\n"
        f"Extract financial data for {year} only:\n\n{text}"
    )
    if not include_raw_lines:
        user_message += _NO_RAW_LINES_NOTE

    for attempt in range(1, retries + 2):
        try:
            print(f"    Claude text call [{model}] (attempt {attempt})...")
            resp = client.messages.create(
                model=model,
                max_tokens=CLAUDE_MAX_TOKENS,
                temperature=CLAUDE_TEMPERATURE,
                system=_SYSTEM_PROMPT_FULL,
                messages=[{"role": "user", "content": user_message}],
            )
            if resp.stop_reason == "max_tokens":
                if include_raw_lines:
                    print(
                        f"    ⚠ Output truncated (hit {CLAUDE_MAX_TOKENS}-token limit) — "
                        f"retrying without raw_lines..."
                    )
                    return _call_text(client, text, ticker, year, retries=retries,
                                      model=model, include_raw_lines=False)
                else:
                    print(f"    ✗ Output still truncated even without raw_lines — statement too large.")
                    return None
            result = _parse_response(resp.content[0].text)
            if result is not None:
                return result
            print(f"    ✗ Empty/invalid JSON on attempt {attempt}, retrying...")
            if attempt <= retries:
                time.sleep(3)
        except anthropic.RateLimitError:
            wait = 60 * attempt   # 60s, 120s, 180s — let the minute window reset
            print(f"    ✗ Rate limit hit (attempt {attempt}) — waiting {wait}s...")
            if attempt <= retries:
                time.sleep(wait)
        except (anthropic.APIError, Exception) as e:
            print(f"    ✗ API error (attempt {attempt}): {e}")
            if attempt <= retries:
                time.sleep(5)

    return None


def _call_vision(client, images: list[dict], ticker: str, year: int, retries: int = 4,
                 model: str = CLAUDE_MODEL, include_raw_lines: bool = True) -> dict | None:
    """Claude vision call for scanned PDFs (image input)."""

    # Claude supports up to 20 images per request — use first 20 pages.
    # At 72 DPI each page is ~660×935 px ≈ 820 tokens; 20 pages ≈ 16k input tokens.
    batch = images[:20]
    if len(images) > 20:
        print(f"    ⚠ Using first 20 of {len(images)} pages")

    vision_text = (
        f"Company ticker: {ticker}\n"
        f"Target year: {year}\n\n"
        f"These images are pages from a Vietnamese consolidated financial statement PDF. "
        f"Extract the Income Statement, Balance Sheet, and Cash Flow data for {year} only. "
        f"Return ONLY the JSON object as specified in your instructions — no markdown, no explanation."
    )
    if not include_raw_lines:
        vision_text += _NO_RAW_LINES_NOTE

    text_prompt = {"type": "text", "text": vision_text}

    for attempt in range(1, retries + 2):
        try:
            print(f"    Claude vision call [{model}] ({len(batch)} images, attempt {attempt})...")
            resp = client.messages.create(
                model=model,
                max_tokens=CLAUDE_MAX_TOKENS,
                temperature=CLAUDE_TEMPERATURE,
                system=_SYSTEM_PROMPT_FULL,
                messages=[{"role": "user", "content": batch + [text_prompt]}],
            )
            if resp.stop_reason == "max_tokens":
                if include_raw_lines:
                    print(
                        f"    ⚠ Output truncated (hit {CLAUDE_MAX_TOKENS}-token limit) — "
                        f"retrying without raw_lines..."
                    )
                    return _call_vision(client, images, ticker, year, retries=retries,
                                        model=model, include_raw_lines=False)
                else:
                    print(f"    ✗ Output still truncated even without raw_lines — statement too large.")
                    return None
            result = _parse_response(resp.content[0].text)
            if result is not None:
                return result
            print(f"    ✗ Empty/invalid JSON on attempt {attempt}, retrying...")
            if attempt <= retries:
                time.sleep(3)
        except anthropic.RateLimitError:
            wait = 60 * attempt   # 60s, 120s, 180s — let the minute window reset
            print(f"    ✗ Rate limit hit (attempt {attempt}) — waiting {wait}s...")
            if attempt <= retries:
                time.sleep(wait)
        except (anthropic.APIError, Exception) as e:
            print(f"    ✗ Vision API error (attempt {attempt}): {e}")
            if attempt <= retries:
                time.sleep(5)

    return None


def _normalize_units(data: dict) -> dict:
    """
    Safety net: detect and fix unit inconsistencies after Claude's response.

    Target unit is MILLION VND (triệu đồng). All output values are integers.

    Detection heuristic using a probe value (first positive number found):
      - raw VND:     probe > 10^12  → divide by 1,000,000
      - billion VND: probe < 10^6   → multiply by 1,000
        (no listed Vietnamese company has any metric below 1M in million VND)
      - million VND: 10^6 ≤ probe ≤ 10^12  → already correct, no action
    """
    # Use the first positive numeric value found as the probe
    probe = None
    for section in ("income_statement", "balance_sheet", "cash_flow"):
        sec = data.get(section, {})
        for k, v in sec.items():
            if k == "raw_lines":
                continue
            if isinstance(v, (int, float)) and v is not True and v is not False and v > 0:
                probe = v
                break
        if probe is not None:
            break

    if probe is None:
        return data  # no numeric data to inspect

    if probe > 1_000_000_000_000:        # raw VND (> 10^12) → divide by 10^6
        factor   = None
        divisor  = 1_000_000
        detected = "raw VND"
    elif probe < 1_000_000:              # billion VND (< 10^6) → multiply by 10^3
        factor   = 1_000
        divisor  = None
        detected = "billion VND"
    else:
        return data  # already in million VND

    print(f"    ⚠ Unit mismatch detected ({detected}) — auto-converting to million VND")

    def _conv(v):
        if not (isinstance(v, (int, float)) and v is not True and v is not False):
            return v
        result = (v * factor) if factor else (v / divisor)
        return round(result)   # integer — no decimal places

    for section in ("income_statement", "balance_sheet", "cash_flow"):
        if section not in data:
            continue
        converted = {}
        for k, v in data[section].items():
            if k == "raw_lines" and isinstance(v, list):
                converted[k] = [
                    {**line, "value": _conv(line["value"])}
                    if isinstance(line.get("value"), (int, float)) else line
                    for line in v
                ]
            else:
                converted[k] = _conv(v)
        data[section] = converted
    data["unit"] = "million VND"
    return data


def _validate(data: dict, ticker: str, year: int) -> dict:
    data.setdefault("company", ticker)
    data.setdefault("year", year)
    data.setdefault("unit", "million VND")
    for key in _REQUIRED_KEYS:
        if key not in data:
            print(f"    ⚠ Missing section '{key}'")
            data[key] = {}
    data = _normalize_units(data)
    return data


def normalize(content, ticker: str, year: int, force: bool = False,
              model: str = CLAUDE_MODEL, include_raw_lines: bool = True) -> dict | None:
    """
    Normalize content for a single company/year.

    Args:
        content:          str (text) or list[dict] (images) from extractor
        ticker:           e.g. 'HPG'
        year:             e.g. 2023
        force:            ignore cache
        model:            Claude model to use (default: Haiku)
        include_raw_lines: whether to request raw_lines in the output
    """
    cache_path = DATA_PROCESSED / f"{ticker}_{year}.json"

    if cache_path.exists() and not force:
        print(f"  → Using cached JSON: {cache_path.name}")
        return json.loads(cache_path.read_text(encoding="utf-8"))

    client = anthropic.Anthropic(api_key=os.environ.get("ANTHROPIC_API_KEY"))

    print(f"  Normalizing {ticker} {year} [{model}]...")

    if isinstance(content, str):
        data = _call_text(client, content, ticker, year,
                          model=model, include_raw_lines=include_raw_lines)
    elif isinstance(content, list):
        data = _call_vision(client, content, ticker, year,
                            model=model, include_raw_lines=include_raw_lines)
    else:
        print(f"  ✗ Unknown content type: {type(content)}")
        return None

    if data is None:
        print(f"  ✗ Normalization failed for {ticker} {year}")
        return None

    data = _validate(data, ticker, year)
    cache_path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"  ✓ Saved → {cache_path.name}")
    return data


def run(extracted: dict | None = None, force: bool = False,
        model: str = CLAUDE_MODEL, include_raw_lines: bool = True) -> list[dict]:
    """
    Normalize all extracted content.

    Args:
        extracted:        {stem: str_or_images} from extractor.run()
        force:            ignore cache
        model:            Claude model to use (default: Haiku)
        include_raw_lines: whether to request raw_lines in the output
    """
    if extracted is None:
        extracted = {}
        for txt_path in sorted(DATA_PROCESSED.glob("*.txt")):
            extracted[txt_path.stem] = txt_path.read_text(encoding="utf-8")

    if not extracted:
        print("No extracted content found. Run extractor first.")
        return []

    print(f"  Model: {model} | raw_lines: {'yes' if include_raw_lines else 'no'}")

    results = []
    for stem, content in sorted(extracted.items()):
        parts = stem.rsplit("_", 1)
        if len(parts) != 2:
            continue
        ticker, year_str = parts
        try:
            year = int(year_str)
        except ValueError:
            continue

        print(f"\n[{stem}]")
        cache_path = DATA_PROCESSED / f"{ticker}_{year}.json"
        used_cache = cache_path.exists() and not force
        data = normalize(content, ticker, year, force=force,
                         model=model, include_raw_lines=include_raw_lines)
        if data:
            results.append(data)
        # Pause between API calls to stay under the 30k tokens/min rate limit.
        # Vision calls consume ~20k–60k tokens each; without a gap the next
        # call arrives before the previous minute's window has cleared.
        # Skip the sleep when data was loaded from cache (no API call was made).
        if not used_cache:
            time.sleep(20)

    print(f"\n── Normalization complete: {len(results)} files ──")
    return results


if __name__ == "__main__":
    run()
