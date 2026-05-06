"use client";

interface ResultCardProps {
  jobId: string;
  outputPaths: string[];
  tickers: string[];
  years: number[];
  error?: string;
  onRunAgain: () => void;
}

export default function ResultCard({
  jobId,
  outputPaths,
  tickers,
  years,
  error,
  onRunAgain,
}: ResultCardProps) {
  const isError = !!error;

  function getFilename(path: string): string {
    return path.split("/").pop() ?? path;
  }

  return (
    <div id="result-card" style={{ display: "flex", flexDirection: "column", gap: "1rem" }}>
      {/* Status banner */}
      <div style={{
        display: "flex", alignItems: "center", gap: "0.75rem",
        padding: "1rem 1.25rem",
        background: isError ? "rgba(217,83,79,0.1)" : "rgba(74,155,111,0.1)",
        border: `1px solid ${isError ? "rgba(217,83,79,0.25)" : "rgba(74,155,111,0.25)"}`,
        borderRadius: "var(--radius-md)",
      }}>
        <span style={{ fontSize: "1.3rem" }}>{isError ? "❌" : "✅"}</span>
        <div>
          <p style={{
            fontWeight: 700, fontSize: "0.95rem",
            color: isError ? "#E07070" : "#6BBF88",
            letterSpacing: "0.02em",
          }}>
            {isError ? "Pipeline failed" : "Report ready"}
          </p>
          <p style={{ fontSize: "0.8rem", color: "rgba(245,240,232,0.45)", marginTop: "0.1rem" }}>
            {tickers.join(", ")} · {years.join(", ")}
          </p>
        </div>
      </div>

      {/* Error detail */}
      {isError && (
        <div style={{
          background: "rgba(5,15,35,0.7)", border: "1px solid rgba(217,83,79,0.2)",
          borderRadius: "var(--radius-sm)", padding: "0.75rem 1rem",
          fontFamily: "var(--font-mono)", fontSize: "0.78rem",
          color: "#E07070", wordBreak: "break-word",
        }}>
          {error}
        </div>
      )}

      {/* Download buttons */}
      {!isError && outputPaths.length > 0 && (
        <div style={{ display: "flex", flexDirection: "column", gap: "0.5rem" }}>
          {outputPaths.map((path, i) => (
            <a
              key={i}
              id={`download-btn-${i}`}
              href={`/api/download/${jobId}?index=${i}`}
              download
              className="btn-cta"
              style={{ textDecoration: "none" }}
            >
              ⬇  DOWNLOAD {getFilename(path)}
            </a>
          ))}
        </div>
      )}

      {/* Run again */}
      <button id="run-again-btn" className="btn-ghost" onClick={onRunAgain}>
        ↩ Run again
      </button>
    </div>
  );
}
