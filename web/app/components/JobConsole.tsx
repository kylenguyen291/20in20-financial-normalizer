"use client";

import { useEffect, useRef, useState } from "react";

export interface LogLine {
  ts: string;
  msg: string;
  kind: "step" | "success" | "error" | "warn" | "info";
}

interface JobConsoleProps {
  jobId: string;
  totalDocs: number; // tickers × years — used to weight step-3 progress
  onComplete: (status: "done" | "error", outputPaths: string[], error?: string) => void;
}

// ── Step definitions ──────────────────────────────────────────────────────────
const PIPELINE_STEPS = [
  { id: 1, label: "Downloading PDFs" },
  { id: 2, label: "Extracting text" },
  { id: 3, label: "Normalizing (Claude)" },
  { id: 4, label: "Exporting to Excel" },
  { id: 5, label: "Comparison workbook" },
];

// Weighted progress ranges: [start%, end%] for each step
const STEP_RANGES: Record<number, [number, number]> = {
  1: [0,   15],
  2: [15,  25],
  3: [25,  80],  // heaviest — Claude API calls
  4: [80,  90],
  5: [90, 100],
};

function classify(msg: string): LogLine["kind"] {
  if (/^\[Step/.test(msg))                                   return "step";
  if (/^✓|^✔|✓ Complete/i.test(msg))                        return "success";
  if (/^✗|^error|^failed|exception|RuntimeError/i.test(msg)) return "error";
  if (/^⚠|^warn|skipping/i.test(msg))                       return "warn";
  return "info";
}

function now(): string {
  return new Date().toLocaleTimeString("en-GB", { hour12: false });
}

function detectStep(msg: string): number | null {
  const m = msg.match(/\[Step (\d)\/5\]/);
  return m ? parseInt(m[1], 10) : null;
}

// ── Progress computation ──────────────────────────────────────────────────────
function computeProgress(lines: LogLine[], totalDocs: number): number {
  let currentStep = 0;
  let hasError    = false;
  let docsNormalized = 0;

  for (const line of lines) {
    const s = detectStep(line.msg);
    if (s && s > currentStep) currentStep = s;
    if (line.kind === "error") hasError = true;
    // Count completed normalizations — "✓ Saved → HPG_2023.json"
    if (/✓ Saved →/.test(line.msg) || /→ Using cached JSON:/.test(line.msg)) {
      docsNormalized++;
    }
  }

  if (hasError || currentStep === 0) return 0;

  const [start, end] = STEP_RANGES[currentStep] ?? [0, 0];

  // Within step 3, interpolate based on docs completed
  if (currentStep === 3 && totalDocs > 0) {
    const fraction = Math.min(docsNormalized / totalDocs, 1);
    return start + fraction * (end - start);
  }

  // For steps 1, 2, 4, 5: once we enter the step, show the start%;
  // we'll jump to end% when we see the next step start
  const nextStepSeen = lines.some(l => {
    const s = detectStep(l.msg);
    return s !== null && s > currentStep;
  });

  if (nextStepSeen) return end;

  // Within the step: animate to midpoint so it doesn't look frozen
  return start + (end - start) * 0.5;
}

// ── Progress bar component ────────────────────────────────────────────────────
function ProgressBar({ pct, hasError, done }: { pct: number; hasError: boolean; done: boolean }) {
  const color = hasError ? "#E07070" : done ? "#6BBF88" : "var(--gold-300)";
  const label = hasError ? "Failed" : done ? "Complete" : `${Math.round(pct)}%`;

  return (
    <div style={{ marginBottom: "1.25rem" }}>
      {/* Top row: label + percentage */}
      <div style={{
        display: "flex", justifyContent: "space-between", alignItems: "center",
        marginBottom: "0.45rem",
      }}>
        <span style={{
          fontSize: "0.7rem", fontWeight: 600, letterSpacing: "0.08em",
          color: "rgba(245,240,232,0.5)", textTransform: "uppercase",
        }}>
          Pipeline Progress
        </span>
        <span style={{
          fontSize: "0.85rem", fontWeight: 700, color,
          transition: "color 0.3s",
          fontVariantNumeric: "tabular-nums",
        }}>
          {label}
        </span>
      </div>

      {/* Bar track */}
      <div style={{
        width: "100%", height: "6px",
        borderRadius: "999px",
        background: "rgba(255,255,255,0.06)",
        overflow: "hidden",
      }}>
        <div style={{
          height: "100%",
          width: `${pct}%`,
          background: hasError
            ? "linear-gradient(90deg, #C0392B, #E07070)"
            : done
              ? "linear-gradient(90deg, #27AE60, #6BBF88)"
              : "linear-gradient(90deg, #8B6914, var(--gold-300))",
          borderRadius: "999px",
          transition: "width 0.6s ease, background 0.4s",
          position: "relative",
          // shimmer on active
          backgroundSize: "200% 100%",
          animation: (!done && !hasError) ? "shimmer 2s linear infinite" : "none",
        }} />
      </div>
    </div>
  );
}

// ── Step tracker component ────────────────────────────────────────────────────
function StepTracker({ lines }: { lines: LogLine[] }) {
  let currentStep = 0;
  let hasError = false;
  for (const line of lines) {
    const s = detectStep(line.msg);
    if (s && s > currentStep) currentStep = s;
    if (line.kind === "error") hasError = true;
  }

  return (
    <div style={{
      display: "flex", gap: "0",
      marginBottom: "1.25rem",
      position: "relative",
    }}>
      {PIPELINE_STEPS.map((step, i) => {
        const done    = step.id < currentStep;
        const active  = step.id === currentStep && !hasError;
        const errored = step.id === currentStep && hasError;

        const color = errored ? "#E07070"
                    : done    ? "#6BBF88"
                    : active  ? "var(--gold-300)"
                    : "rgba(245,240,232,0.2)";

        return (
          <div key={step.id} style={{ flex: 1, display: "flex", flexDirection: "column", alignItems: "center", gap: "0.4rem", position: "relative" }}>
            {i < PIPELINE_STEPS.length - 1 && (
              <div style={{
                position: "absolute", top: "12px", left: "50%",
                width: "100%", height: "2px",
                background: step.id < currentStep ? "#6BBF88" : "rgba(245,240,232,0.08)",
                transition: "background 0.4s", zIndex: 0,
              }} />
            )}
            <div style={{
              width: "24px", height: "24px", borderRadius: "50%",
              background: errored ? "rgba(224,112,112,0.15)"
                        : done    ? "rgba(107,191,136,0.15)"
                        : active  ? "rgba(201,168,76,0.15)"
                        : "rgba(255,255,255,0.05)",
              border: `2px solid ${color}`,
              display: "flex", alignItems: "center", justifyContent: "center",
              fontSize: "0.7rem", color, zIndex: 1, position: "relative",
              transition: "all 0.3s", flexShrink: 0,
            }}>
              {errored ? "✗" : done ? "✓" : active ? (
                <span style={{
                  width: "8px", height: "8px", borderRadius: "50%",
                  border: "2px solid var(--gold-300)", borderTopColor: "transparent",
                  display: "inline-block", animation: "spin 0.7s linear infinite",
                }} />
              ) : step.id}
            </div>
            <span style={{
              fontSize: "0.6rem", fontWeight: active || done ? 600 : 400,
              color, textAlign: "center", letterSpacing: "0.02em",
              lineHeight: 1.3, transition: "color 0.3s", maxWidth: "72px",
            }}>
              {step.label}
            </span>
          </div>
        );
      })}
    </div>
  );
}

// ── Main component ────────────────────────────────────────────────────────────
export default function JobConsole({ jobId, totalDocs, onComplete }: JobConsoleProps) {
  const [lines, setLines]     = useState<LogLine[]>([]);
  const [isDone, setIsDone]   = useState(false);
  const [hasError, setHasErr] = useState(false);
  const bottomRef             = useRef<HTMLDivElement>(null);

  useEffect(() => {
    if (!jobId) return;
    setLines([]);
    setIsDone(false);
    setHasErr(false);

    // Connect directly to FastAPI — the Next.js proxy buffers SSE chunks,
    // which causes logs to arrive all at once at the end instead of streaming.
    const API_BASE = process.env.NEXT_PUBLIC_API_URL ?? "http://localhost:8000";
    const es = new EventSource(`${API_BASE}/api/logs/${jobId}`);

    es.onmessage = (e) => {
      try {
        const payload = JSON.parse(e.data);

        if (payload.type === "log") {
          const msg = payload.message as string;
          setLines((prev) => [...prev, { ts: now(), msg, kind: classify(msg) }]);
        }

        if (payload.type === "complete") {
          es.close();
          setIsDone(payload.status === "done");
          setHasErr(payload.status === "error");
          onComplete(payload.status, payload.output_paths ?? [], payload.error ?? undefined);
        }
      } catch { /* ignore */ }
    };

    es.onerror = () => {
      es.close();
      setLines((prev) => [...prev, { ts: now(), msg: "Connection lost — check backend.", kind: "error" }]);
      setHasErr(true);
    };

    return () => { es.close(); };
  }, [jobId]); // eslint-disable-line react-hooks/exhaustive-deps

  useEffect(() => {
    bottomRef.current?.scrollIntoView({ behavior: "smooth" });
  }, [lines]);

  const pct = isDone ? 100 : computeProgress(lines, totalDocs);

  return (
    <div>
      {/* Progress bar */}
      <ProgressBar pct={pct} hasError={hasError} done={isDone} />

      {/* Step tracker */}
      <StepTracker lines={lines} />

      {/* Log console */}
      <div id="job-console" className="console" aria-live="polite" aria-label="Job log output">
        {lines.length === 0 && (
          <div style={{ color: "rgba(245,240,232,0.2)", fontStyle: "italic" }}>
            Starting pipeline — output will appear here momentarily…
          </div>
        )}

        {lines.map((line, i) => (
          <div key={i} className={`console-line ${line.kind}`}>
            <span className="ts">{line.ts}</span>
            <span className="msg">{line.msg}</span>
          </div>
        ))}

        <div ref={bottomRef} />
      </div>
    </div>
  );
}
