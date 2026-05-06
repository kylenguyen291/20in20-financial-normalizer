"""
app.py  —  Single-process FastAPI app for Railway deployment.
Serves the HTML frontend + all API endpoints on one port.
No Next.js required.
"""

import asyncio
import io
import json
import sys
import threading
import uuid
from pathlib import Path
from typing import Optional

from dotenv import load_dotenv
from fastapi import FastAPI, HTTPException
from fastapi.staticfiles import StaticFiles
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse, StreamingResponse
from pydantic import BaseModel

load_dotenv()

# ── ensure project root on sys.path ──────────────────────────────────────────
ROOT = Path(__file__).parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

app = FastAPI(title="Financial Normalizer")
app.add_middleware(CORSMiddleware, allow_origins=["*"], allow_methods=["*"], allow_headers=["*"])

# Serve the static/ directory (fault-tolerant in case path differs)
_static_dir = ROOT / "static"
if _static_dir.exists():
    app.mount("/static", StaticFiles(directory=_static_dir), name="static")
    print(f"[startup] Serving static files from {_static_dir}")
else:
    print(f"[startup] WARNING: static/ directory not found at {_static_dir}")


# ── In-memory job store ───────────────────────────────────────────────────────
_jobs: dict[str, dict] = {}


def _log(job_id: str, msg: str) -> None:
    """Write a log line directly to the job store — bypasses sys.stdout entirely."""
    for line in msg.splitlines():
        stripped = line.rstrip()
        if stripped:
            _jobs[job_id]["logs"].append(stripped)


class _LogCapture(io.TextIOBase):
    """Also capture any library-level print() calls from pipeline modules."""
    def __init__(self, job_id: str, real_stdout):
        self._job_id = job_id
        self._real = real_stdout
        self._buf = ""

    def write(self, s: str) -> int:
        self._real.write(s)
        self._real.flush()
        self._buf += s
        while "\n" in self._buf:
            line, self._buf = self._buf.split("\n", 1)
            stripped = line.rstrip()
            if stripped:
                _jobs[self._job_id]["logs"].append(stripped)
        return len(s)

    def flush(self):
        self._real.flush()


def _run_pipeline(job_id: str, tickers: list[str], years: list[int]) -> None:
    real_stdout = sys.stdout
    sys.stdout = _LogCapture(job_id, real_stdout)
    try:
        from config import CLAUDE_MODEL_SONNET
        model = CLAUDE_MODEL_SONNET

        _log(job_id, f"  Tickers : {', '.join(tickers)}")
        _log(job_id, f"  Years   : {', '.join(str(y) for y in years)}")
        _log(job_id, f"  Model   : {model}")

        # Step 1
        _log(job_id, "\n[Step 1/5] Downloading PDFs from CafeF...")
        from src.downloader import run as download
        pdf_paths = download(tickers=tickers, years=years)

        if not pdf_paths:
            _log(job_id, "  ⚠ No PDFs downloaded — checking cache in data/raw/")
            from config import DATA_RAW
            pdf_paths = []
            for ticker in tickers:
                for year in years:
                    p = DATA_RAW / f"{ticker}_{year}.pdf"
                    if p.exists():
                        pdf_paths.append(p)
                        _log(job_id, f"  → Found cached: {p.name}")
            if not pdf_paths:
                raise RuntimeError("No PDFs found — download failed and no cached files exist.")

        # Step 2
        _log(job_id, "\n[Step 2/5] Extracting text from PDFs...")
        from src.extractor import run as extract
        extracted = extract(pdf_paths=pdf_paths)
        if not extracted:
            raise RuntimeError("Text extraction returned nothing.")
        _log(job_id, f"  ✓ Extracted {len(extracted)} document(s)")

        # Step 3
        _log(job_id, "\n[Step 3/5] Normalizing with Claude API...")
        from src.normalizer import run as normalize
        all_data = normalize(extracted=extracted, model=model)
        if not all_data:
            raise RuntimeError("Claude normalization returned no results — check ANTHROPIC_API_KEY.")
        _log(job_id, f"  ✓ Normalized {len(all_data)} document(s)")

        # Step 4
        _log(job_id, "\n[Step 4/5] Exporting to Excel...")
        from src.exporter import run as export
        filtered = [d for d in all_data if d is not None]
        saved = export(all_data=filtered, years=years)
        if not saved:
            raise RuntimeError("Exporter produced no output files.")

        # Step 5
        _log(job_id, "\n[Step 5/5] Building comparison workbook...")
        from src.comparison_exporter import build_comparison
        comp = build_comparison(all_data=filtered, years=years)
        if comp:
            saved = list(saved) + [comp]

        _jobs[job_id]["output_paths"] = [str(p) for p in saved]
        _jobs[job_id]["status"] = "done"
        _log(job_id, f"\n✓ Complete — {len(saved)} workbook(s) exported.")

    except Exception as exc:
        _jobs[job_id]["status"] = "error"
        _jobs[job_id]["error"] = str(exc)
        _log(job_id, f"\n✗ Pipeline failed: {exc}")
    finally:
        sys.stdout = real_stdout


# ── Request schema ────────────────────────────────────────────────────────────
class RunRequest(BaseModel):
    tickers: list[str]
    years: list[int]


# ── API endpoints ─────────────────────────────────────────────────────────────
@app.get("/api/health")
def health():
    return {"status": "ok"}


@app.post("/api/run")
def start_run(req: RunRequest):
    tickers = [t.strip().upper() for t in req.tickers if t.strip()]
    years = [y for y in req.years if 2000 <= y <= 2100]
    if not tickers:
        raise HTTPException(422, "At least one ticker required.")
    if not years:
        raise HTTPException(422, "At least one valid year required.")

    job_id = str(uuid.uuid4())
    _jobs[job_id] = {"status": "running", "logs": [f"Starting pipeline for {', '.join(tickers)} · {', '.join(str(y) for y in years)}..."], "output_paths": [], "error": None}
    threading.Thread(target=_run_pipeline, args=(job_id, tickers, years), daemon=True).start()
    return {"job_id": job_id}


@app.get("/api/status/{job_id}")
def poll_status(job_id: str, offset: int = 0):
    job = _jobs.get(job_id)
    if not job:
        raise HTTPException(404, "Job not found.")
    logs = job["logs"]
    new_lines = logs[offset:]
    return JSONResponse(
        content={
            "lines": new_lines,
            "next": offset + len(new_lines),
            "status": job["status"],
            "output_paths": job.get("output_paths", []),
            "error": job.get("error"),
        },
        headers={"Cache-Control": "no-store, no-cache, must-revalidate"},
    )


@app.get("/api/download/{job_id}")
def download_file(job_id: str, index: int = 0):
    job = _jobs.get(job_id)
    if not job:
        raise HTTPException(404, "Job not found.")
    if job["status"] != "done":
        raise HTTPException(409, "Job not complete yet.")
    paths = job.get("output_paths", [])
    if index >= len(paths):
        raise HTTPException(404, "File not found.")
    path = Path(paths[index])
    if not path.exists():
        raise HTTPException(404, "File removed.")
    return FileResponse(str(path),
                        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        filename=path.name)


# ── Frontend ──────────────────────────────────────────────────────────────────
@app.get("/", response_class=HTMLResponse)
def index():
    html = open(ROOT / "static" / "index.html", encoding="utf-8").read()
    return HTMLResponse(
        html,
        headers={"Cache-Control": "no-store, no-cache, must-revalidate"},
    )
