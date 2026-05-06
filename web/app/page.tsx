"use client";

import { useState } from "react";
import TickerInput from "./components/TickerInput";
import JobConsole from "./components/JobConsole";
import ResultCard from "./components/ResultCard";

type AppState = "idle" | "running" | "done";

function parseYears(raw: string): number[] {
  return raw
    .split(",")
    .map((s) => parseInt(s.trim(), 10))
    .filter((n) => !isNaN(n) && n >= 1990 && n <= 2100);
}

export default function Home() {
  const [tickers, setTickers]       = useState<string[]>([]);
  const [yearsRaw, setYearsRaw]     = useState<string>("");
  const [state, setState]           = useState<AppState>("idle");
  const [jobId, setJobId]           = useState<string>("");
  const [outputPaths, setOutputPaths] = useState<string[]>([]);
  const [jobError, setJobError]     = useState<string | undefined>();
  const [submitError, setSubmitError] = useState<string | undefined>();

  const years = parseYears(yearsRaw);

  async function handleRun() {
    setSubmitError(undefined);
    if (tickers.length === 0) { setSubmitError("Add at least one ticker."); return; }
    if (years.length === 0)   { setSubmitError("Enter at least one valid year (e.g. 2022, 2023)."); return; }

    try {
      const API_BASE = process.env.NEXT_PUBLIC_API_URL ?? "http://localhost:8000";
      const res = await fetch(`${API_BASE}/api/run`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ tickers, years, use_sonnet: true }),
      });
      if (!res.ok) {
        const err = await res.json().catch(() => ({}));
        throw new Error(err.detail ?? `HTTP ${res.status}`);
      }
      const { job_id } = await res.json();
      setJobId(job_id);
      setState("running");
    } catch (e: unknown) {
      const msg = e instanceof Error ? e.message : String(e);
      setSubmitError(`Failed to start job: ${msg}`);
    }
  }

  function handleComplete(_status: "done" | "error", paths: string[], error?: string) {
    setOutputPaths(paths);
    setJobError(error);
    setState("done");
  }

  function handleRunAgain() {
    setState("idle");
    setJobId("");
    setOutputPaths([]);
    setJobError(undefined);
    setSubmitError(undefined);
  }

  return (
    <div className="page-wrapper">

      {/* ── Top bar ── */}
      <nav className="topbar">
        <div className="topbar-logo">
          <span className="logo-mark">20<span>in</span>20</span>
          <span className="logo-sub">PARTNERS</span>
        </div>
        <span className="topbar-tool">Financial Normalizer</span>
      </nav>

      {/* ── Hero ── */}
      <div className="hero">
        <p className="hero-eyebrow">🇻🇳 Vietnamese Capital Markets</p>
        <h1 className="hero-title">
          Financial Statement<br />Normalizer
        </h1>
        <p className="hero-sub">
          Enter a stock ticker and fiscal years — the pipeline downloads, extracts, and normalizes
          the annual report into a structured Excel workbook automatically.
        </p>
      </div>

      {/* ── Form Card ── */}
      <div className="form-card">

        {state !== "done" && (
          <>
            {/* Tickers */}
            <div className="field">
              <label htmlFor="ticker-text-input" className="field-label">
                Stock Tickers
              </label>
              <TickerInput
                tickers={tickers}
                onChange={setTickers}
                disabled={state === "running"}
              />
              <p className="field-hint">Press Enter or comma to add · Backspace to remove last</p>
            </div>

            {/* Years */}
            <div className="field">
              <label htmlFor="years-input" className="field-label">
                Fiscal Years
              </label>
              <input
                id="years-input"
                className="text-input"
                type="text"
                placeholder="e.g. 2022, 2023, 2024"
                value={yearsRaw}
                onChange={(e) => setYearsRaw(e.target.value)}
                disabled={state === "running"}
              />
              <p className="field-hint">
                Comma-separated · any year(s) you need
                {years.length > 0 && (
                  <span className="field-hint-parsed"> → {years.join(", ")}</span>
                )}
              </p>
            </div>

            <div className="divider" />

            {submitError && (
              <p className="error-msg">{submitError}</p>
            )}

            <button
              id="run-btn"
              className="btn-cta"
              onClick={handleRun}
              disabled={state === "running"}
            >
              {state === "running" ? (
                <><span className="spinner" /> RUNNING PIPELINE…</>
              ) : (
                "▶  RUN PIPELINE"
              )}
            </button>
          </>
        )}

        {/* Live console */}
        {state === "running" && (
          <>
            <div className="divider" />
            <div className="console-header">
              <span className="pulse-dot" />
              <span>Live output</span>
              <span className="console-meta">{tickers.join(", ")} · {years.join(", ")}</span>
            </div>
            <JobConsole jobId={jobId} totalDocs={tickers.length * years.length} onComplete={handleComplete} />
          </>
        )}

        {/* Result */}
        {state === "done" && (
          <ResultCard
            jobId={jobId}
            outputPaths={outputPaths}
            tickers={tickers}
            years={years}
            error={jobError}
            onRunAgain={handleRunAgain}
          />
        )}
      </div>

      {/* ── Footer ── */}
      <footer className="page-footer">
        <p>Powered by CafeF · Claude API · Python pipeline</p>
        <a href="/api/health" className="footer-link">API health ↗</a>
      </footer>
    </div>
  );
}
