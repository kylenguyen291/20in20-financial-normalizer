"""
app.py
──────
Streamlit web interface for the Vietnamese Financial Statement Normalizer.
Styled to match 20in20 Partners brand: deep navy, gold accents, serif headings.
"""

import io
import os
import sys
import threading
import queue
import zipfile
import time
from pathlib import Path

import streamlit as st
from dotenv import load_dotenv

# ── Page config ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Financial Normalizer · 20in20 Partners",
    page_icon="📊",
    layout="centered",
)

load_dotenv()

BASE_DIR = Path(__file__).parent
sys.path.insert(0, str(BASE_DIR))

# ── 20in20 Brand Styles ───────────────────────────────────────────────────────
# Navy: #1e3a6e  |  Dark navy: #162d57  |  Gold: #c9a042  |  Light text: #d4dff0
st.markdown("""
<style>
  /* ── Global background & font ── */
  html, body, [data-testid="stAppViewContainer"], [data-testid="stApp"] {
      background-color: #1a3460 !important;
      color: #eaf0fb !important;
  }
  [data-testid="stMain"] { background-color: #1a3460 !important; }
  [data-testid="stSidebar"] { background-color: #162d57 !important; }

  /* ── Main content width ── */
  .block-container { max-width: 740px !important; padding-top: 2rem !important; }

  /* ── Label tag above title ── */
  .label-tag {
      color: #c9a042;
      font-size: 0.72rem;
      font-weight: 700;
      letter-spacing: 0.2em;
      text-transform: uppercase;
      text-align: center;
      margin-bottom: 0.4rem;
  }

  /* ── Hero title ── */
  .hero-title {
      font-family: Georgia, 'Times New Roman', serif;
      font-size: 2.1rem;
      font-weight: 700;
      color: #ffffff;
      text-align: center;
      line-height: 1.3;
      margin-bottom: 0.5rem;
  }

  /* ── Gold divider ── */
  .gold-divider {
      width: 48px; height: 3px;
      background: #c9a042;
      margin: 0.8rem auto 1.2rem auto;
      border-radius: 2px;
  }

  /* ── Subtitle ── */
  .hero-sub {
      color: #a0b8d8;
      font-size: 0.95rem;
      text-align: center;
      max-width: 520px;
      margin: 0 auto 2rem auto;
      line-height: 1.7;
  }

  /* ── Section labels (above inputs) ── */
  .stTextInput label, .stSelectbox label {
      color: #a0b8d8 !important;
      font-size: 0.78rem !important;
      font-weight: 600 !important;
      letter-spacing: 0.08em !important;
      text-transform: uppercase !important;
  }

  /* ── Input fields ── */
  .stTextInput input {
      background-color: #162d57 !important;
      border: 1px solid rgba(201,160,66,0.35) !important;
      border-radius: 4px !important;
      color: #eaf0fb !important;
      font-size: 0.95rem !important;
  }
  .stTextInput input:focus {
      border-color: #c9a042 !important;
      box-shadow: 0 0 0 2px rgba(201,160,66,0.2) !important;
  }
  .stTextInput input::placeholder { color: #5a7aa0 !important; }

  /* ── Primary CTA button (gold fill) ── */
  div[data-testid="stButton"]:first-of-type > button {
      background-color: #c9a042 !important;
      color: #ffffff !important;
      font-weight: 700 !important;
      font-size: 0.82rem !important;
      letter-spacing: 0.12em !important;
      text-transform: uppercase !important;
      border: none !important;
      border-radius: 3px !important;
      padding: 0.75rem 2rem !important;
      width: 100% !important;
      transition: background 0.2s !important;
  }
  div[data-testid="stButton"]:first-of-type > button:hover {
      background-color: #b8902e !important;
  }

  /* ── Download buttons (navy border style) ── */
  [data-testid="stDownloadButton"] button {
      background-color: transparent !important;
      color: #eaf0fb !important;
      border: 1px solid rgba(201,160,66,0.5) !important;
      border-radius: 3px !important;
      font-size: 0.78rem !important;
      font-weight: 600 !important;
      letter-spacing: 0.1em !important;
      text-transform: uppercase !important;
      padding: 0.55rem 1.2rem !important;
      transition: all 0.2s !important;
  }
  [data-testid="stDownloadButton"] button:hover {
      background-color: rgba(201,160,66,0.15) !important;
      border-color: #c9a042 !important;
  }

  /* ── Expander ── */
  [data-testid="stExpander"] {
      background-color: #162d57 !important;
      border: 1px solid rgba(201,160,66,0.2) !important;
      border-radius: 4px !important;
  }
  [data-testid="stExpander"] summary {
      color: #a0b8d8 !important;
      font-size: 0.82rem !important;
      font-weight: 600 !important;
      letter-spacing: 0.05em !important;
  }

  /* ── Checkbox ── */
  .stCheckbox label { color: #a0b8d8 !important; font-size: 0.88rem !important; }
  .stCheckbox [data-testid="stCheckbox"] { accent-color: #c9a042; }

  /* ── Alert / info boxes ── */
  [data-testid="stAlert"] {
      background-color: rgba(201,160,66,0.1) !important;
      border-left: 3px solid #c9a042 !important;
      border-radius: 3px !important;
      color: #eaf0fb !important;
  }

  /* ── Log box ── */
  .log-box {
      background: #0d1f3c;
      color: #7fb3d3;
      font-family: 'Courier New', monospace;
      font-size: 0.76rem;
      padding: 1rem 1.2rem;
      border-radius: 4px;
      border: 1px solid rgba(201,160,66,0.15);
      max-height: 320px;
      overflow-y: auto;
      white-space: pre-wrap;
      line-height: 1.6;
  }

  /* ── Step progress pills ── */
  .step-pill {
      display: inline-block;
      background: rgba(201,160,66,0.15);
      border: 1px solid rgba(201,160,66,0.4);
      color: #c9a042;
      font-size: 0.72rem;
      font-weight: 700;
      letter-spacing: 0.1em;
      text-transform: uppercase;
      padding: 0.25rem 0.75rem;
      border-radius: 20px;
      margin: 0.15rem 0.2rem;
  }
  .step-pill.done { background: rgba(46,160,67,0.15); border-color: #2ea043;
                    color: #4cbb6c; }

  /* ── Section divider ── */
  hr { border-color: rgba(201,160,66,0.2) !important; }

  /* ── Footer ── */
  .footer {
      color: #4a6a8a;
      font-size: 0.75rem;
      text-align: center;
      padding: 1.5rem 0 0.5rem 0;
      letter-spacing: 0.04em;
  }
  .footer a { color: #c9a042 !important; text-decoration: none; }

  /* ── Success card ── */
  .success-card {
      background: rgba(46,160,67,0.1);
      border: 1px solid rgba(46,160,67,0.4);
      border-radius: 4px;
      padding: 1rem 1.2rem;
      color: #4cbb6c;
      font-weight: 600;
      font-size: 0.9rem;
      margin: 1rem 0;
  }

  /* ── Download section heading ── */
  .dl-heading {
      color: #c9a042;
      font-size: 0.72rem;
      font-weight: 700;
      letter-spacing: 0.18em;
      text-transform: uppercase;
      margin: 1.4rem 0 0.6rem 0;
  }
</style>
""", unsafe_allow_html=True)


# ── Hero Header ───────────────────────────────────────────────────────────────
st.markdown('<p class="label-tag">20IN20 PARTNERS · RESEARCH TOOLS</p>', unsafe_allow_html=True)
st.markdown('<h1 class="hero-title">Vietnamese Financial<br>Statement Normalizer</h1>', unsafe_allow_html=True)
st.markdown('<div class="gold-divider"></div>', unsafe_allow_html=True)
st.markdown(
    '<p class="hero-sub">Enter Vietnamese listed company tickers and target years. '
    'The tool downloads their annual reports from CafeF, extracts financial tables using AI, '
    'and delivers clean Excel workbooks ready for analysis.</p>',
    unsafe_allow_html=True,
)

st.markdown("---")

# ── Pipeline flowchart ────────────────────────────────────────────────────────
st.markdown("""
<div style="display:flex; align-items:center; justify-content:center;
            flex-wrap:nowrap; gap:0; margin-bottom:2rem; overflow-x:auto;">

  <div style="display:flex; flex-direction:column; align-items:center; min-width:90px;">
    <div style="width:36px; height:36px; border-radius:50%;
                background:rgba(201,160,66,0.15); border:2px solid #c9a042;
                display:flex; align-items:center; justify-content:center;
                color:#c9a042; font-weight:700; font-size:0.85rem;">1</div>
    <div style="color:#c9a042; font-size:0.65rem; font-weight:700;
                letter-spacing:0.12em; text-transform:uppercase;
                margin-top:0.4rem; text-align:center;">Download</div>
    <div style="color:#7090b0; font-size:0.6rem; text-align:center;
                margin-top:0.15rem;">CafeF PDFs</div>
  </div>

  <div style="color:#c9a042; font-size:1.1rem; padding:0 0.3rem;
              margin-bottom:1.2rem; opacity:0.7;">→</div>

  <div style="display:flex; flex-direction:column; align-items:center; min-width:90px;">
    <div style="width:36px; height:36px; border-radius:50%;
                background:rgba(201,160,66,0.15); border:2px solid #c9a042;
                display:flex; align-items:center; justify-content:center;
                color:#c9a042; font-weight:700; font-size:0.85rem;">2</div>
    <div style="color:#c9a042; font-size:0.65rem; font-weight:700;
                letter-spacing:0.12em; text-transform:uppercase;
                margin-top:0.4rem; text-align:center;">Extract</div>
    <div style="color:#7090b0; font-size:0.6rem; text-align:center;
                margin-top:0.15rem;">Text / Images</div>
  </div>

  <div style="color:#c9a042; font-size:1.1rem; padding:0 0.3rem;
              margin-bottom:1.2rem; opacity:0.7;">→</div>

  <div style="display:flex; flex-direction:column; align-items:center; min-width:90px;">
    <div style="width:36px; height:36px; border-radius:50%;
                background:rgba(201,160,66,0.15); border:2px solid #c9a042;
                display:flex; align-items:center; justify-content:center;
                color:#c9a042; font-weight:700; font-size:0.85rem;">3</div>
    <div style="color:#c9a042; font-size:0.65rem; font-weight:700;
                letter-spacing:0.12em; text-transform:uppercase;
                margin-top:0.4rem; text-align:center;">Normalize</div>
    <div style="color:#7090b0; font-size:0.6rem; text-align:center;
                margin-top:0.15rem;">Claude AI</div>
  </div>

  <div style="color:#c9a042; font-size:1.1rem; padding:0 0.3rem;
              margin-bottom:1.2rem; opacity:0.7;">→</div>

  <div style="display:flex; flex-direction:column; align-items:center; min-width:90px;">
    <div style="width:36px; height:36px; border-radius:50%;
                background:rgba(201,160,66,0.15); border:2px solid #c9a042;
                display:flex; align-items:center; justify-content:center;
                color:#c9a042; font-weight:700; font-size:0.85rem;">4</div>
    <div style="color:#c9a042; font-size:0.65rem; font-weight:700;
                letter-spacing:0.12em; text-transform:uppercase;
                margin-top:0.4rem; text-align:center;">Export</div>
    <div style="color:#7090b0; font-size:0.6rem; text-align:center;
                margin-top:0.15rem;">Excel workbooks</div>
  </div>

  <div style="color:#c9a042; font-size:1.1rem; padding:0 0.3rem;
              margin-bottom:1.2rem; opacity:0.7;">→</div>

  <div style="display:flex; flex-direction:column; align-items:center; min-width:90px;">
    <div style="width:36px; height:36px; border-radius:50%;
                background:rgba(201,160,66,0.15); border:2px solid #c9a042;
                display:flex; align-items:center; justify-content:center;
                color:#c9a042; font-weight:700; font-size:0.85rem;">5</div>
    <div style="color:#c9a042; font-size:0.65rem; font-weight:700;
                letter-spacing:0.12em; text-transform:uppercase;
                margin-top:0.4rem; text-align:center;">Compare</div>
    <div style="color:#7090b0; font-size:0.6rem; text-align:center;
                margin-top:0.15rem;">Ratio analysis</div>
  </div>

</div>
""", unsafe_allow_html=True)

# ── Inputs ────────────────────────────────────────────────────────────────────
col1, col2 = st.columns([1, 1])
with col1:
    tickers_input = st.text_input(
        "Stock tickers",
        placeholder="HPG, HSG, NKG",
        help="Vietnamese listed company tickers from HOSE / HNX / UPCoM",
    )
with col2:
    years_input = st.text_input(
        "Years",
        placeholder="2022, 2023, 2024",
        value="2022, 2023, 2024",
    )

api_key_input = st.text_input(
    "Anthropic API Key",
    type="password",
    value=os.environ.get("ANTHROPIC_API_KEY", ""),
    help="From console.anthropic.com — never stored or logged",
)

with st.expander("Advanced options"):
    use_sonnet  = st.checkbox("Use Claude Sonnet (higher accuracy, ~4× cost)", value=True)
    no_raw      = st.checkbox("Skip raw lines output (faster & cheaper)", value=False)
    force_rerun = st.checkbox("Force re-run — ignore all cached results", value=False)

st.markdown("<br>", unsafe_allow_html=True)
run_button = st.button("RUN PIPELINE", use_container_width=True)


# ── Helpers ───────────────────────────────────────────────────────────────────
def _parse_list(s):
    return [x.strip().upper() for x in s.split(",") if x.strip()]

def _parse_years(s):
    years = []
    for x in s.split(","):
        x = x.strip()
        if x.isdigit():
            years.append(int(x))
    return sorted(years)

def _zip_outputs(output_dir):
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for f in output_dir.glob("*.xlsx"):
            zf.write(f, f.name)
    buf.seek(0)
    return buf.read()

class _StreamCapture(io.StringIO):
    def __init__(self, q):
        super().__init__()
        self._q = q
    def write(self, s):
        if s and s != "\n":
            self._q.put(s)
        return len(s)
    def flush(self):
        pass


# ── Pipeline runner (background thread) ──────────────────────────────────────
def run_pipeline(tickers, years, api_key, use_sonnet, no_raw, force, log_q, result_q):
    try:
        os.environ["ANTHROPIC_API_KEY"] = api_key
        old_stdout = sys.stdout
        sys.stdout = _StreamCapture(log_q)

        from config import CLAUDE_MODEL, CLAUDE_MODEL_SONNET, OUTPUT_DIR
        model = CLAUDE_MODEL_SONNET if use_sonnet else CLAUDE_MODEL
        include_raw = not no_raw

        log_q.put("[Step 1/5] Downloading PDFs from CafeF...\n")
        from src.downloader import run as download
        pdf_paths = download(tickers=tickers, years=years)

        log_q.put("\n[Step 2/5] Extracting text from PDFs...\n")
        from src.extractor import run as extract
        extracted = extract(pdf_paths=pdf_paths or None, force=force)

        log_q.put(f"\n[Step 3/5] Normalizing with Claude [{model}]...\n")
        from src.normalizer import run as normalize
        all_data = normalize(extracted=extracted, force=force,
                             model=model, include_raw_lines=include_raw)

        log_q.put("\n[Step 4/5] Exporting to Excel...\n")
        from src.exporter import run as export
        filtered = [d for d in all_data if d.get("company") in tickers]
        if not filtered:
            result_q.put({"ok": False, "error": "No data extracted. Check tickers and years."})
            return
        saved = export(all_data=filtered, years=years)

        log_q.put("\n[Step 5/5] Building comparison workbook...\n")
        from src.comparison_exporter import build_comparison
        build_comparison(all_data=filtered, years=years)

        sys.stdout = old_stdout
        result_q.put({"ok": True, "output_dir": OUTPUT_DIR})

    except Exception as e:
        try: sys.stdout = old_stdout
        except: pass
        result_q.put({"ok": False, "error": str(e)})


# ── Run on button click ───────────────────────────────────────────────────────
if run_button:
    tickers = _parse_list(tickers_input)
    years   = _parse_years(years_input)

    if not tickers:
        st.error("Please enter at least one ticker (e.g. HPG, VHM).")
        st.stop()
    if not years:
        st.error("Please enter at least one valid year (e.g. 2022, 2023, 2024).")
        st.stop()
    if not api_key_input.strip():
        st.error("Please enter your Anthropic API key.")
        st.stop()

    st.markdown(
        f'<div style="color:#a0b8d8;font-size:0.82rem;margin-bottom:1rem;">'
        f'Running for <strong style="color:#eaf0fb">{", ".join(tickers)}</strong> '
        f'· Years: <strong style="color:#eaf0fb">{", ".join(map(str, years))}</strong>'
        f'</div>',
        unsafe_allow_html=True,
    )

    log_placeholder = st.empty()
    log_q    = queue.Queue()
    result_q = queue.Queue()
    log_lines = []

    thread = threading.Thread(
        target=run_pipeline,
        args=(tickers, years, api_key_input, use_sonnet, no_raw, force_rerun, log_q, result_q),
        daemon=True,
    )
    thread.start()

    with st.spinner("Pipeline running — typically 10–20 min for 3 companies × 3 years..."):
        while thread.is_alive() or not log_q.empty():
            try:
                while True:
                    log_lines.append(log_q.get_nowait())
            except queue.Empty:
                pass
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
                label=label,
                data=f.read_bytes(),
                file_name=f.name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key=f.name,
                use_container_width=True,
            )

        if xlsx_files:
            st.markdown("<br>", unsafe_allow_html=True)
            st.download_button(
                label="DOWNLOAD ALL AS ZIP  ↓",
                data=_zip_outputs(output_dir),
                file_name="financial_reports.zip",
                mime="application/zip",
                key="zip_all",
                use_container_width=True,
            )

# ── Footer ────────────────────────────────────────────────────────────────────
st.markdown("---")
st.markdown(
    '<p class="footer">'
    'Data sourced from <a href="https://cafef.vn" target="_blank">CafeF</a> · '
    'Normalized using Claude AI · Values in million VND<br>'
    '© 2026 <a href="https://20in20partners.com" target="_blank">20in20 Partners</a> · '
    'Ho Chi Minh City · Los Angeles'
    '</p>',
    unsafe_allow_html=True,
)
