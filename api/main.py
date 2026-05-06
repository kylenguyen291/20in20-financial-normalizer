"""
main.py
───────
FastAPI backend for the Vietnamese Financial Statement Normalizer.

Endpoints:
  POST /api/run              → start a pipeline job
  GET  /api/logs/{job_id}    → SSE stream of log lines + completion event
  GET  /api/download/{job_id}→ stream the generated Excel file
  GET  /api/health           → health check
"""

import asyncio
import json
import threading
from pathlib import Path

from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, StreamingResponse
from pydantic import BaseModel

from api.job_runner import create_job, get_job, run_pipeline

app = FastAPI(title="Financial Normalizer API")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)


# ── Request schema ─────────────────────────────────────────────────────────────

class RunRequest(BaseModel):
    tickers:    list[str]
    years:      list[int]
    use_sonnet: bool = False


# ── Endpoints ──────────────────────────────────────────────────────────────────

@app.get("/api/health")
def health():
    return {"status": "ok"}


@app.post("/api/run")
def start_run(req: RunRequest):
    """Validate inputs, spawn a background thread, return job_id."""
    tickers = [t.strip().upper() for t in req.tickers if t.strip()]
    years   = [y for y in req.years if 2000 <= y <= 2100]

    if not tickers:
        raise HTTPException(status_code=422, detail="At least one ticker required.")
    if not years:
        raise HTTPException(status_code=422, detail="At least one valid year required.")

    job_id = create_job()
    thread = threading.Thread(
        target=run_pipeline,
        args=(job_id, tickers, years, req.use_sonnet),
        daemon=True,
    )
    thread.start()
    return {"job_id": job_id}


@app.get("/api/logs/{job_id}")
async def stream_logs(job_id: str):
    """
    Server-Sent Events stream.
    Emits:  { type: "log",      message: "..." }
    Emits:  { type: "complete", status: "done"|"error", output_paths: [...], error: "..." }
    """
    if not get_job(job_id):
        raise HTTPException(status_code=404, detail="Job not found.")

    async def event_generator():
        sent = 0
        while True:
            job  = get_job(job_id)
            logs = job["logs"]

            # Flush new log lines
            while sent < len(logs):
                payload = json.dumps({"type": "log", "message": logs[sent]})
                yield f"data: {payload}\n\n"
                sent += 1

            # Check terminal state
            if job["status"] in ("done", "error"):
                # Flush any final lines first
                while sent < len(logs):
                    payload = json.dumps({"type": "log", "message": logs[sent]})
                    yield f"data: {payload}\n\n"
                    sent += 1
                complete = json.dumps({
                    "type":         "complete",
                    "status":       job["status"],
                    "output_paths": job.get("output_paths", []),
                    "error":        job.get("error"),
                })
                yield f"data: {complete}\n\n"
                break

            await asyncio.sleep(0.25)

    return StreamingResponse(
        event_generator(),
        media_type="text/event-stream",
        headers={
            "Cache-Control": "no-cache",
            "X-Accel-Buffering": "no",   # disable nginx buffering
        },
    )


@app.get("/api/download/{job_id}")
def download_file(job_id: str, index: int = 0):
    """Download the generated Excel workbook by job ID."""
    job = get_job(job_id)
    if not job:
        raise HTTPException(status_code=404, detail="Job not found.")
    if job["status"] != "done":
        raise HTTPException(status_code=409, detail="Job not complete yet.")

    paths = job.get("output_paths", [])
    if index >= len(paths) or not paths:
        raise HTTPException(status_code=404, detail="Output file not found.")

    path = Path(paths[index])
    if not path.exists():
        raise HTTPException(status_code=404, detail="File has been removed.")

    return FileResponse(
        str(path),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename=path.name,
    )
