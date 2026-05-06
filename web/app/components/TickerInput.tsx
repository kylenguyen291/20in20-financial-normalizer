"use client";

import { useState, useRef, KeyboardEvent } from "react";

interface TickerInputProps {
  tickers: string[];
  onChange: (tickers: string[]) => void;
  disabled?: boolean;
}

export default function TickerInput({ tickers, onChange, disabled }: TickerInputProps) {
  const [value, setValue] = useState("");
  const inputRef = useRef<HTMLInputElement>(null);

  function addTicker(raw: string) {
    const t = raw.trim().toUpperCase().replace(/[^A-Z0-9]/g, "");
    if (!t) return;
    if (!tickers.includes(t)) {
      onChange([...tickers, t]);
    }
    setValue("");
  }

  function removeTicker(t: string) {
    onChange(tickers.filter((x) => x !== t));
  }

  function handleKey(e: KeyboardEvent<HTMLInputElement>) {
    if (e.key === "Enter" || e.key === "," || e.key === " ") {
      e.preventDefault();
      addTicker(value);
    } else if (e.key === "Backspace" && value === "" && tickers.length > 0) {
      onChange(tickers.slice(0, -1));
    }
  }

  return (
    <div
      id="ticker-input-wrapper"
      style={{
        display: "flex",
        flexWrap: "wrap",
        alignItems: "center",
        gap: "0.5rem",
        padding: "0.5rem 0.75rem",
        background: "var(--bg-elevated)",
        border: "1px solid var(--border-muted)",
        borderRadius: "var(--radius-md)",
        cursor: "text",
        transition: "border-color 0.18s, box-shadow 0.18s",
        minHeight: "46px",
      }}
      onClick={() => inputRef.current?.focus()}
      onFocus={() => {
        const el = document.getElementById("ticker-input-wrapper");
        if (el) {
          el.style.borderColor = "var(--accent-gold)";
          el.style.boxShadow = "0 0 0 3px rgba(240,180,41,0.12)";
        }
      }}
      onBlur={() => {
        const el = document.getElementById("ticker-input-wrapper");
        if (el) {
          el.style.borderColor = "var(--border-muted)";
          el.style.boxShadow = "none";
        }
      }}
    >
      {tickers.map((t) => (
        <span key={t} className="chip" id={`chip-${t}`}>
          {t}
          {!disabled && (
            <button
              onClick={(e) => { e.stopPropagation(); removeTicker(t); }}
              aria-label={`Remove ${t}`}
              title="Remove"
            >
              ×
            </button>
          )}
        </span>
      ))}

      {!disabled && (
        <input
          ref={inputRef}
          id="ticker-text-input"
          value={value}
          onChange={(e) => setValue(e.target.value.toUpperCase())}
          onKeyDown={handleKey}
          onBlur={() => { if (value) addTicker(value); }}
          placeholder={tickers.length === 0 ? "HPG, MBB, VNM…" : "Add ticker…"}
          style={{
            background: "transparent",
            border: "none",
            outline: "none",
            color: "var(--text-primary)",
            fontFamily: "var(--font-mono, monospace)",
            fontSize: "0.88rem",
            fontWeight: 600,
            letterSpacing: "0.04em",
            minWidth: "100px",
            flex: 1,
          }}
        />
      )}
    </div>
  );
}
