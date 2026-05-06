"""
job_runner.py
─────────────
Manages pipeline jobs in background threads.
Captures stdout from the pipeline and stores log lines per job.
"""

import io
import sys
import uuid
from pathlib import Path
from typing import Optional

# ── In-memory job store ───────────────────────────────────────────────────────
# job shape: {status, logs, output_paths, error}
_jobs: dict[str, dict] = {}


class _LogCapture(io.TextIOBase):
    """
    Wraps sys.stdout to intercept print() calls from the pipeline.
    Each non-empty line is appended to the job's log list AND written
    to the real stdout so server logs stay visible.
    """

    def __init__(self, job_id: str, real_stdout):
        self._job_id = job_id
        self._real = real_stdout
        self._buf = ""

    def write(self, s: str) -> int:
        self._real.write(s)
        self._real.flush()
        self._buf += s
        # Flush complete lines to the log
        while "\n" in self._buf:
            line, self._buf = self._buf.split("\n", 1)
            stripped = line.rstrip()
            if stripped:
                _jobs[self._job_id]["logs"].append(stripped)
        return len(s)

    def flush(self):
        self._real.flush()


def create_job() -> str:
    """Create a new job entry and return its ID."""
    job_id = str(uuid.uuid4())
    _jobs[job_id] = {
        "status":       "running",
        "logs":         [],
        "output_paths": [],
        "error":        None,
    }
    return job_id


def get_job(job_id: str) -> Optional[dict]:
    return _jobs.get(job_id)


def run_pipeline(job_id: str, tickers: list[str], years: list[int],
                 use_sonnet: bool = False) -> None:
    """
    Run the full 4-step pipeline in a background thread.
    Captures all stdout into the job log list.
    """
    real_stdout = sys.stdout
    capture     = _LogCapture(job_id, real_stdout)
    sys.stdout  = capture

    try:
        # Ensure the project root is on the path so src.* imports work
        root = str(Path(__file__).parent.parent)
        if root not in sys.path:
            sys.path.insert(0, root)

        # Load .env so ANTHROPIC_API_KEY is available inside the thread
        from dotenv import load_dotenv
        load_dotenv(Path(root) / ".env", override=False)

        from config import CLAUDE_MODEL, CLAUDE_MODEL_SONNET
        model = CLAUDE_MODEL_SONNET if use_sonnet else CLAUDE_MODEL

        print(f"  Tickers : {', '.join(tickers)}")
        print(f"  Years   : {', '.join(str(y) for y in years)}")
        print(f"  Model   : {model}")

        # ── Step 1: Download ──────────────────────────────────────────────────
        print("\n[Step 1/4] Downloading PDFs from CafeF...")
        from src.downloader import run as download
        pdf_paths = download(tickers=tickers, years=years)

        if not pdf_paths:
            print("  ⚠ No PDFs downloaded — checking cache in data/raw/")
            from config import DATA_RAW
            pdf_paths = []
            for ticker in tickers:
                for year in years:
                    p = DATA_RAW / f"{ticker}_{year}.pdf"
                    if p.exists():
                        pdf_paths.append(p)
                        print(f"  → Found cached: {p.name}")
            if not pdf_paths:
                raise RuntimeError(
                    "No PDFs found — download failed and no cached files exist in data/raw/."
                )

        # ── Step 2: Extract ───────────────────────────────────────────────────
        print("\n[Step 2/4] Extracting text from PDFs...")
        from src.extractor import run as extract
        extracted = extract(pdf_paths=pdf_paths or None)

        if not extracted:
            raise RuntimeError(
                "Text extraction returned nothing — PDFs may be corrupted or image-only."
            )

        print(f"  ✓ Extracted {len(extracted)} document(s)")

        # ── Step 3: Normalize ─────────────────────────────────────────────────
        print("\n[Step 3/4] Normalizing with Claude API...")
        from src.normalizer import run as normalize
        all_data = normalize(extracted=extracted, model=model)

        if not all_data:
            raise RuntimeError(
                "Claude normalization returned no results — "
                "check ANTHROPIC_API_KEY and that PDFs contain Vietnamese financial tables."
            )

        print(f"  ✓ Normalized {len(all_data)} document(s)")

        # ── Step 4: Export ────────────────────────────────────────────────────
        print("\n[Step 4/4] Exporting to Excel...")
        from src.exporter import run as export

        # Filter out any None results but don't filter by company name —
        # normalize.run() already scopes output to the requested tickers/years
        filtered = [d for d in all_data if d is not None]

        if not filtered:
            raise RuntimeError(
                "Normalization completed but produced no valid records to export."
            )

        saved = export(all_data=filtered, years=years)

        if not saved:
            raise RuntimeError("Exporter ran but produced no output files.")

        # ── Step 5: Comparison workbook ───────────────────────────────────────
        print("\n[Step 5/5] Building comparison workbook...")
        from src.comparison_exporter import build_comparison
        comparison_path = build_comparison(all_data=filtered, years=years)
        if comparison_path:
            saved = list(saved) + [comparison_path]
            print(f"  ✓ comparison.xlsx saved")
        else:
            print("  ⚠ Comparison workbook skipped (need ≥2 companies for meaningful comparison)")

        _jobs[job_id]["output_paths"] = [str(p) for p in saved]
        _jobs[job_id]["status"]       = "done"
        print(f"\n✓ Complete — {len(saved)} workbook(s) exported.")

    except Exception as exc:
        _jobs[job_id]["status"] = "error"
        _jobs[job_id]["error"]  = str(exc)
        print(f"\n✗ Pipeline failed: {exc}")

    finally:
        sys.stdout = real_stdout
