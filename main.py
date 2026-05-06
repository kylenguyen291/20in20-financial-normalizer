"""
main.py
───────
CLI entry point for the Vietnamese Financial Statement Normalizer.

The program interactively asks the user for:
  1. Number of years to extract
  2. Each year (one by one)
  3. Stock tickers / mã cổ phiếu (one by one, Enter to finish)

Pipeline flags (optional):
  --skip-download    Use existing PDFs in data/raw/
  --skip-extract     Use cached .txt files in data/processed/
  --skip-normalize   Use cached .json files, re-export only
  --force            Ignore all caches and re-run every step
"""

import argparse
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))

from dotenv import load_dotenv
load_dotenv()

from config import DEFAULT_YEARS, OUTPUT_DIR


# ── Interactive input helpers ─────────────────────────────────────────────────

def _ask_years() -> list[int]:
    """Ask user how many years, then collect each year."""
    print("\n── Years ────────────────────────────────────────────────")
    while True:
        raw = input("  How many years do you want to extract? ").strip()
        try:
            n = int(raw)
            if n < 1:
                print("  ✗ Please enter at least 1.")
                continue
            break
        except ValueError:
            print("  ✗ Please enter a number.")

    years = []
    for i in range(1, n + 1):
        while True:
            raw = input(f"  Year {i}: ").strip()
            try:
                year = int(raw)
                if year < 2000 or year > 2100:
                    print("  ✗ Please enter a valid year (e.g. 2024).")
                    continue
                if year in years:
                    print(f"  ✗ {year} already added.")
                    continue
                years.append(year)
                break
            except ValueError:
                print("  ✗ Please enter a valid year (e.g. 2024).")

    return sorted(years)


def _ask_tickers() -> list[str]:
    """Ask user to enter stock tickers one by one. Empty input = done."""
    print("\n── Stock tickers / Mã cổ phiếu ─────────────────────────")
    print("  Enter one ticker per line. Press Enter on an empty line when done.")
    tickers = []
    i = 1
    while True:
        raw = input(f"  Ticker {i}: ").strip().upper()
        if raw == "":
            if not tickers:
                print("  ✗ Please enter at least one ticker.")
                continue
            break
        if raw in tickers:
            print(f"  ✗ {raw} already added.")
            continue
        tickers.append(raw)
        i += 1
    return tickers




def _parse_pipeline_flags():
    """Parse only the pipeline control flags (not tickers/years)."""
    parser = argparse.ArgumentParser(
        description="Vietnamese Financial Statement Normalizer",
        add_help=True,
    )
    parser.add_argument("--skip-download",  action="store_true",
                        help="Skip PDF download — use existing files in data/raw/")
    parser.add_argument("--skip-extract",   action="store_true",
                        help="Skip text extraction — use cached .txt in data/processed/")
    parser.add_argument("--skip-normalize", action="store_true",
                        help="Skip Claude normalization — use cached .json, re-export only")
    parser.add_argument("--force",          action="store_true",
                        help="Ignore all caches and re-run every step")
    parser.add_argument("--use-sonnet",     action="store_true",
                        help="Use Claude Sonnet instead of Haiku (higher accuracy, ~4x cost)")
    parser.add_argument("--no-raw-lines",   action="store_true",
                        help="Omit raw_lines from output (smaller JSON, fewer output tokens)")
    return parser.parse_args()


# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    args = _parse_pipeline_flags()

    print("=" * 60)
    print("  Vietnamese Financial Statement Normalizer")
    print("=" * 60)

    # Interactive: collect years and tickers
    years   = _ask_years()
    tickers = _ask_tickers()
    force   = args.force

    # Model selection
    from config import CLAUDE_MODEL, CLAUDE_MODEL_SONNET
    model = CLAUDE_MODEL_SONNET if args.use_sonnet else CLAUDE_MODEL
    include_raw_lines = not args.no_raw_lines

    print("\n" + "=" * 60)
    print(f"  Tickers  : {', '.join(tickers)}")
    print(f"  Years    : {years}")
    print(f"  Output   : {OUTPUT_DIR}/")
    print(f"  Model    : {model}")
    print(f"  raw_lines: {'yes' if include_raw_lines else 'no (--no-raw-lines)'}")
    print("=" * 60)

    # ── Step 1: Download ──────────────────────────────────────────
    if not args.skip_download and not args.skip_extract:
        print("\n[Step 1/4] Downloading PDFs from CafeF...")
        from src.downloader import run as download
        pdf_paths = download(tickers=tickers, years=years)
    else:
        print("\n[Step 1/4] Skipping download")
        pdf_paths = None

    # ── Step 2: Extract ───────────────────────────────────────────
    if not args.skip_extract and not args.skip_normalize:
        print("\n[Step 2/4] Extracting text from PDFs...")
        from src.extractor import run as extract
        from config import DATA_RAW

        if pdf_paths is None:
            pdf_paths = []
            for ticker in tickers:
                for year in years:
                    p = DATA_RAW / f"{ticker}_{year}.pdf"
                    if p.exists():
                        pdf_paths.append(p)

        extracted = extract(pdf_paths=pdf_paths or None, force=force)
    else:
        print("\n[Step 2/4] Skipping extraction")
        extracted = None

    # ── Step 3: Normalize ─────────────────────────────────────────
    if not args.skip_normalize:
        print("\n[Step 3/4] Normalizing with Claude API...")
        from src.normalizer import run as normalize
        all_data = normalize(extracted=extracted, force=force,
                             model=model, include_raw_lines=include_raw_lines)
    else:
        print("\n[Step 3/4] Skipping normalization — loading cached JSON...")
        import json
        from config import DATA_PROCESSED
        all_data = []
        for ticker in tickers:
            for year in years:
                json_path = DATA_PROCESSED / f"{ticker}_{year}.json"
                if json_path.exists():
                    all_data.append(json.loads(json_path.read_text(encoding="utf-8")))
                else:
                    print(f"  ⚠ No cached JSON for {ticker}_{year} — run without --skip-normalize")

    # ── Step 4: Export ────────────────────────────────────────────
    print("\n[Step 4/4] Exporting to Excel...")
    from src.exporter import run as export

    filtered = [d for d in all_data if d.get("company") in tickers]

    if not filtered:
        print("  ✗ No data to export. Check previous steps.")
        sys.exit(1)

    saved = export(all_data=filtered, years=years)

    # ── Step 5: Comparison workbook ───────────────────────────────
    print("\n[Step 5/5] Building comparison workbook...")
    from src.comparison_exporter import build_comparison
    comparison_path = build_comparison(all_data=filtered, years=years)

    # ── Summary ───────────────────────────────────────────────────
    print("\n" + "=" * 60)
    print("  Done!")
    for path in saved:
        print(f"  → {path}")
    if comparison_path:
        print(f"  → {comparison_path}  (comparison)")
    print("=" * 60)


if __name__ == "__main__":
    main()
