"""
app.py
──────
Streamlit web app for the Vietnamese Financial Statement Normalizer.
Streamlit Cloud compatible — users upload PDFs directly (no Playwright needed).
Styled to match 20in20 Partners brand.
"""

import io
import os
import sys
import tempfile
import threading
import queue
import zipfile
import time
from pathlib import Path

import subprocess

import streamlit as st
from dotenv import load_dotenv

# ── Install Playwright browser once (required on Streamlit Cloud) ──────────────
@st.cache_resource(show_spinner=False)
def _install_playwright():
    subprocess.run(["playwright", "install", "chromium"], check=False)

_install_playwright()

st.set_page_config(
    page_title="Financial Normalizer · 20in20 Partners",
    page_icon="📊",
    layout="centered",
)

load_dotenv()
BASE_DIR = Path(__file__).parent
sys.path.insert(0, str(BASE_DIR))

# ── Brand CSS ─────────────────────────────────────────────────────────────────
st.markdown("""
<style>
  html, body, [data-testid="stAppViewContainer"], [data-testid="stApp"] {
      background-color: #1a3460 !important; color: #eaf0fb !important; }
  [data-testid="stMain"] { background-color: #1a3460 !important; }
  .block-container { max-width: 740px !important; padding-top: 2rem !important; }

  .label-tag { color:#c9a042; font-size:0.72rem; font-weight:700;
      letter-spacing:0.2em; text-transform:uppercase; text-align:center; }
  .hero-title { font-family:Georgia,'Times New Roman',serif; font-size:2.1rem;
      font-weight:700; color:#fff; text-align:center; line-height:1.3; }
  .gold-divider { width:48px; height:3px; background:#c9a042;
      margin:0.8rem auto 1.2rem auto; border-radius:2px; }
  .hero-sub { color:#a0b8d8; font-size:0.95rem; text-align:center;
      max-width:520px; margin:0 auto 2rem auto; line-height:1.7; }

  .stTextInput label, .stSelectbox label {
      color:#a0b8d8 !important; font-size:0.78rem !important;
      font-weight:600 !important; letter-spacing:0.08em !important;
      text-transform:uppercase !important; }
  .stTextInput input {
      background-color:#162d57 !important;
      border:1px solid rgba(201,160,66,0.35) !important;
      border-radius:4px !important; color:#eaf0fb !important; }
  .stTextInput input:focus {
      border-color:#c9a042 !important;
      box-shadow:0 0 0 2px rgba(201,160,66,0.2) !important; }
  .stTextInput input::placeholder { color:#5a7aa0 !important; }

  /* File uploader */
  [data-testid="stFileUploader"] {
      background-color:#162d57 !important;
      border:1px dashed rgba(201,160,66,0.5) !important;
      border-radius:6px !important; }
  [data-testid="stFileUploader"] label {
      color:#a0b8d8 !important; font-size:0.78rem !important;
      font-weight:600 !important; letter-spacing:0.08em !important;
      text-transform:uppercase !important; }

  /* Primary button */
  div[data-testid="stButton"] > button {
      background-color:#c9a042 !important; color:#fff !important;
      font-weight:700 !important; font-size:0.82rem !important;
      letter-spacing:0.12em !important; text-transform:uppercase !important;
      border:none !important; border-radius:3px !important;
      padding:0.75rem 2rem !important; width:100% !important; }
  div[data-testid="stButton"] > button:hover { background-color:#b8902e !important; }

  /* Download buttons */
  [data-testid="stDownloadButton"] button {
      background-color:transparent !important; color:#eaf0fb !important;
      border:1px solid rgba(201,160,66,0.5) !important; border-radius:3px !important;
      font-size:0.78rem !important; font-weight:600 !important;
      letter-spacing:0.1em !important; text-transform:uppercase !important; }
  [data-testid="stDownloadButton"] button:hover {
      background-color:rgba(201,160,66,0.15) !important;
      border-color:#c9a042 !important; }

  [data-testid="stExpander"] {
      background-color:#162d57 !important;
      border:1px solid rgba(201,160,66,0.2) !important; border-radius:4px !important; }
  [data-testid="stExpander"] summary {
      color:#a0b8d8 !important; font-size:0.82rem !important;
      font-weight:600 !important; }
  .stCheckbox label { color:#a0b8d8 !important; font-size:0.88rem !important; }
  [data-testid="stAlert"] {
      background-color:rgba(201,160,66,0.1) !important;
      border-left:3px solid #c9a042 !important; color:#eaf0fb !important; }

  .log-box { background:#0d1f3c; color:#7fb3d3;
      font-family:'Courier New',monospace; font-size:0.76rem;
      padding:1rem 1.2rem; border-radius:4px;
      border:1px solid rgba(201,160,66,0.15);
      max-height:320px; overflow-y:auto; white-space:pre-wrap; line-height:1.6; }
  .success-card { background:rgba(46,160,67,0.1);
      border:1px solid rgba(46,160,67,0.4); border-radius:4px;
      padding:1rem 1.2rem; color:#4cbb6c; font-weight:600; font-size:0.9rem; }
  .dl-heading { color:#c9a042; font-size:0.72rem; font-weight:700;
      letter-spacing:0.18em; text-transform:uppercase; margin:1.4rem 0 0.6rem 0; }
  .hint { color:#5a7aa0; font-size:0.75rem; margin-top:0.3rem; }
  hr { border-color:rgba(201,160,66,0.2) !important; }
  .footer { color:#4a6a8a; font-size:0.75rem; text-align:center;
      padding:1.5rem 0 0.5rem 0; letter-spacing:0.04em; }
  .footer a { color:#c9a042 !important; text-decoration:none; }
</style>
""", unsafe_allow_html=True)

# ── Hero ──────────────────────────────────────────────────────────────────────
st.markdown('<p class="label-tag">20IN20 PARTNERS · RESEARCH TOOLS</p>', unsafe_allow_html=True)
st.markdown('<h1 class="hero-title">Vietnamese Financial<br>Statement Normalizer</h1>', unsafe_allow_html=True)
st.markdown('<div class="gold-divider"></div>', unsafe_allow_html=True)
st.markdown(
    '<p class="hero-sub">Upload Vietnamese annual report PDFs, enter your API key, '
    'and get clean Excel workbooks with normalized financials, ratio analysis, '
    'and AI-generated insights.</p>',
    unsafe_allow_html=True,
)
st.markdown("---")

# ── Pipeline flowchart ────────────────────────────────────────────────────────
st.markdown("""
<div style="display:flex;align-items:center;justify-content:center;
            flex-wrap:nowrap;gap:0;margin-bottom:2rem;overflow-x:auto;">
""" + "".join([
    f"""
  <div style="display:flex;flex-direction:column;align-items:center;min-width:90px;">
    <div style="width:36px;height:36px;border-radius:50%;
                background:rgba(201,160,66,0.15);border:2px solid #c9a042;
                display:flex;align-items:center;justify-content:center;
                color:#c9a042;font-weight:700;font-size:0.85rem;">{n}</div>
    <div style="color:#c9a042;font-size:0.65rem;font-weight:700;letter-spacing:0.12em;
                text-transform:uppercase;margin-top:0.4rem;text-align:center;">{label}</div>
    <div style="color:#7090b0;font-size:0.6rem;text-align:center;margin-top:0.15rem;">{sub}</div>
  </div>
  {"<div style='color:#c9a042;font-size:1.1rem;padding:0 0.3rem;margin-bottom:1.2rem;opacity:0.7;'>→</div>" if n < 5 else ""}
"""
    for n, label, sub in [
        (1, "Upload", "PDF files"),
        (2, "Extract", "Text / Images"),
        (3, "Normalize", "Claude AI"),
        (4, "Export", "Excel workbooks"),
        (5, "Compare", "Ratio analysis"),
    ]
]) + "</div>", unsafe_allow_html=True)

# ── Inputs ────────────────────────────────────────────────────────────────────
uploaded_files = st.file_uploader(
    "Upload PDF files",
    type=["pdf"],
    accept_multiple_files=True,
    help="Name your files as TICKER_YEAR.pdf (e.g. HPG_2023.pdf) for automatic detection",
)
st.markdown(
    '<p class="hint">📎 Name files as <strong>TICKER_YEAR.pdf</strong> '
    '— e.g. HPG_2023.pdf, VHM_2024.pdf</p>',
    unsafe_allow_html=True,
)

col1, col2 = st.columns([1, 1])
with col1:
    tickers_input = st.text_input(
        "Tickers (auto-detected or override)",
        placeholder="HPG, HSG, NKG",
        help="Leave blank to auto-detect from filenames",
    )
with col2:
    years_input = st.text_input(
        "Years (auto-detected or override)",
        placeholder="2022, 2023, 2024",
        help="Leave blank to auto-detect from filenames",
    )

api_key_input = st.text_input(
    "Anthropic API Key",
    type="password",
    value=os.environ.get("ANTHROPIC_API_KEY", ""),
    help="From console.anthropic.com — not stored",
)

with st.expander("Advanced options"):
    use_sonnet  = st.checkbox("Use Claude Sonnet (higher accuracy, ~4× cost)", value=True)
    no_raw      = st.checkbox("Skip raw lines (faster & cheaper)", value=False)

st.markdown("<br>", unsafe_allow_html=True)
run_button = st.button("RUN PIPELINE", use_container_width=True)


# ── Helpers ───────────────────────────────────────────────────────────────────
def _parse_list(s):
    return [x.strip().upper() for x in s.split(",") if x.strip()] if s.strip() else []

def _parse_years(s):
    years = []
    for x in s.split(","):
        x = x.strip()
        if x.isdigit():
            years.append(int(x))
    return sorted(years) if years else []

def _zip_outputs(output_dir):
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for f in output_dir.glob("*.xlsx"):
            zf.write(f, f.name)
    buf.seek(0)
    return buf.read()

class _StreamCapture(io.StringIO):
    def __init__(self, q):
        super().__init__(); self._q = q
    def write(self, s):
        if s and s != "\n": self._q.put(s)
        return len(s)
    def flush(self): pass


# ── Pipeline runner ───────────────────────────────────────────────────────────
def run_pipeline(pdf_dir, tickers, years, api_key, use_sonnet, no_raw, log_q, result_q):
    old_stdout = sys.stdout
    try:
        os.environ["ANTHROPIC_API_KEY"] = api_key
        sys.stdout = _StreamCapture(log_q)

        from config import CLAUDE_MODEL, CLAUDE_MODEL_SONNET
        model = CLAUDE_MODEL_SONNET if use_sonnet else CLAUDE_MODEL
        include_raw = not no_raw

        # Step 1: collect uploaded PDFs
        pdf_paths = list(Path(pdf_dir).glob("*.pdf"))
        log_q.put(f"[Step 1/5] Found {len(pdf_paths)} uploaded PDF(s)\n")
        for p in pdf_paths:
            log_q.put(f"  · {p.name}\n")

        # Step 2: Extract
        log_q.put("\n[Step 2/5] Extracting text from PDFs...\n")
        from src.extractor import run as extract
        extracted = extract(pdf_paths=pdf_paths, force=True)

        # Step 3: Normalize
        log_q.put(f"\n[Step 3/5] Normalizing with Claude [{model}]...\n")
        from src.normalizer import run as normalize
        all_data = normalize(extracted=extracted, force=True,
                             model=model, include_raw_lines=include_raw)

        # Step 4: Export
        log_q.put("\n[Step 4/5] Exporting to Excel...\n")
        from src.exporter import run as export
        filtered = [d for d in all_data if d.get("company") in tickers] if tickers else all_data
        if not filtered:
            result_q.put({"ok": False, "error": "No data extracted. Check your PDF filenames (must be TICKER_YEAR.pdf)."})
            return
        saved = export(all_data=filtered, years=years)

        # Step 5: Compare
        log_q.put("\n[Step 5/5] Building comparison workbook...\n")
        from src.comparison_exporter import build_comparison
        build_comparison(all_data=filtered, years=years)

        sys.stdout = old_stdout

        from config import OUTPUT_DIR
        result_q.put({"ok": True, "output_dir": OUTPUT_DIR})

    except Exception as e:
        sys.stdout = old_stdout
        import traceback
        result_q.put({"ok": False, "error": f"{e}\n\n{traceback.format_exc()}"})


# ── Run ───────────────────────────────────────────────────────────────────────
if run_button:
    if not uploaded_files:
        st.error("Please upload at least one PDF file named TICKER_YEAR.pdf (e.g. HPG_2023.pdf).")
        st.stop()
    if not api_key_input.strip():
        st.error("Please enter your Anthropic API key.")
        st.stop()

    # Auto-detect tickers and years from filenames
    detected_tickers, detected_years = [], []
    for f in uploaded_files:
        stem = Path(f.name).stem  # e.g. "HPG_2023"
        parts = stem.rsplit("_", 1)
        if len(parts) == 2 and parts[1].isdigit():
            detected_tickers.append(parts[0].upper())
            detected_years.append(int(parts[1]))

    tickers = _parse_list(tickers_input) or sorted(set(detected_tickers))
    years   = _parse_years(years_input)  or sorted(set(detected_years))

    if not tickers or not years:
        st.error("Could not detect tickers/years from filenames. Name files as TICKER_YEAR.pdf or fill in the fields manually.")
        st.stop()

    st.markdown(
        f'<div style="color:#a0b8d8;font-size:0.82rem;margin-bottom:1rem;">'
        f'Processing <strong style="color:#eaf0fb">{", ".join(tickers)}</strong> '
        f'· Years: <strong style="color:#eaf0fb">{", ".join(map(str, years))}</strong>'
        f'</div>',
        unsafe_allow_html=True,
    )

    # Save uploaded PDFs to a temp directory
    tmp_dir = tempfile.mkdtemp()
    for f in uploaded_files:
        dest = Path(tmp_dir) / f.name
        dest.write_bytes(f.read())

    log_placeholder = st.empty()
    log_q    = queue.Queue()
    result_q = queue.Queue()
    log_lines = []

    thread = threading.Thread(
        target=run_pipeline,
        args=(tmp_dir, tickers, years, api_key_input, use_sonnet, no_raw, log_q, result_q),
        daemon=True,
    )
    thread.start()

    with st.spinner("Pipeline running — typically 10–20 min for 3 companies × 3 years..."):
        while thread.is_alive() or not log_q.empty():
            try:
                while True: log_lines.append(log_q.get_nowait())
            except queue.Empty: pass
            log_placeholder.markdown(
                f'<div class="log-box">{"".join(log_lines)}</div>',
                unsafe_allow_html=True,
            )
            time.sleep(0.3)

    while not log_q.empty():
        log_lines.append(log_q.get())
    log_placeholder.markdown(
        f'<div class="log-box">{"".join(log_lines)}</div>',
        unsafe_allow_html=True,
    )

    result = result_q.get()

    if not result["ok"]:
        st.error(f"Pipeline failed: {result['error']}")
    else:
        st.markdown(
            '<div class="success-card">✓ &nbsp; Pipeline complete — your files are ready below</div>',
            unsafe_allow_html=True,
        )
        from config import OUTPUT_DIR
        output_dir = Path(str(OUTPUT_DIR))
        xlsx_files = sorted(output_dir.glob("*.xlsx"))

        st.markdown('<p class="dl-heading">Download Files</p>', unsafe_allow_html=True)
        for f in xlsx_files:
            label = "COMPARISON WORKBOOK  ↓" if "comparison" in f.name else f"{f.stem.upper()}  ↓"
            st.download_button(
                label=label, data=f.read_bytes(), file_name=f.name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key=f.name, use_container_width=True,
            )
        if xlsx_files:
            st.markdown("<br>", unsafe_allow_html=True)
            st.download_button(
                label="DOWNLOAD ALL AS ZIP  ↓", data=_zip_outputs(output_dir),
                file_name="financial_reports.zip", mime="application/zip",
                key="zip_all", use_container_width=True,
            )

# ── Footer ────────────────────────────────────────────────────────────────────
st.markdown("---")
st.markdown(
    '<p class="footer">'
    'Normalized using Claude AI · Values in million VND<br>'
    '© 2026 <a href="https://20in20partners.com" target="_blank">20in20 Partners</a> · '
    'Ho Chi Minh City · Los Angeles'
    '</p>',
    unsafe_allow_html=True,
)
