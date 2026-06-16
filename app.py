"""
app.py  –  Mutual Fund Brokerage Structure Filler
Run:  streamlit run app.py
"""
import streamlit as st
import zipfile, io, traceback
from parsers import (
    detect_source_amc, detect_target_amc, get_parser,
    fill_workbook, fill_workbook_vijay, detect_format, read_excel_wb,
)

st.set_page_config(
    page_title="BrokerageAI · MF Filler",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── Global CSS ─────────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&family=Space+Grotesk:wght@400;500;600;700;800&display=swap');

/* ── Variables ── */
:root {
  --bg:           #080d1a;
  --bg2:          #0d1428;
  --g0:           rgba(255,255,255,0.04);
  --g1:           rgba(255,255,255,0.07);
  --gb:           rgba(255,255,255,0.10);
  --gb2:          rgba(255,255,255,0.18);
  --blue:         #3b82f6;
  --blue-l:       #60a5fa;
  --purple:       #8b5cf6;
  --purple-l:     #a78bfa;
  --cyan:         #06b6d4;
  --cyan-l:       #22d3ee;
  --green:        #10b981;
  --green-l:      #34d399;
  --amber:        #f59e0b;
  --amber-l:      #fbbf24;
  --red:          #ef4444;
  --red-l:        #f87171;
  --t1:           #f1f5f9;
  --t2:           #94a3b8;
  --t3:           #64748b;
}

/* ── Base ── */
*, *::before, *::after { box-sizing: border-box; }
html, body { background: var(--bg) !important; }
.stApp    { background: var(--bg) !important; font-family: 'Inter', sans-serif !important; }

/* Animated deep-space background */
.stApp::before {
  content: '';
  position: fixed; inset: 0; pointer-events: none; z-index: 0;
  background:
    radial-gradient(ellipse 70% 55% at 15%  8%, rgba(59,130,246,.14) 0%, transparent 65%),
    radial-gradient(ellipse 55% 45% at 85% 85%, rgba(139,92,246,.11) 0%, transparent 65%),
    radial-gradient(ellipse 45% 35% at 55% 35%, rgba(6,182,212,.08)  0%, transparent 55%),
    radial-gradient(ellipse 40% 40% at 80% 15%, rgba(16,185,129,.06) 0%, transparent 50%);
  animation: bgPulse 18s ease-in-out infinite alternate;
}
@keyframes bgPulse {
  0%   { opacity: 1; }
  50%  { opacity: .65; }
  100% { opacity: 1; }
}

/* ── Layout ── */
.block-container {
  max-width: 1380px !important;
  padding: 1.5rem 2.5rem 3rem !important;
  position: relative; z-index: 1;
}

/* ── Typography ── */
h1, h2, h3, h4 { font-family: 'Space Grotesk', sans-serif !important; }

/* ── Sidebar ── */
[data-testid="stSidebar"] > div:first-child {
  background: rgba(8,13,26,.97) !important;
  backdrop-filter: blur(30px) !important;
  border-right: 1px solid var(--gb) !important;
}
[data-testid="stSidebar"] * { color: var(--t1) !important; }
[data-testid="stSidebar"] .stCaption p { color: var(--t3) !important; }

/* Sidebar section headers */
[data-testid="stSidebar"] .stMarkdown strong { color: var(--t1) !important; }

/* ── Hero section ── */
.hero {
  text-align: center;
  padding: 2rem 1rem 2.5rem;
  animation: fadeDown .7s ease both;
}
.hero-badge {
  display: inline-flex; align-items: center; gap: .45rem;
  background: rgba(59,130,246,.12);
  border: 1px solid rgba(59,130,246,.28);
  border-radius: 100px;
  padding: .3rem .95rem;
  font-size: .72rem; font-weight: 700; letter-spacing: .1em;
  text-transform: uppercase; color: var(--blue-l);
  margin-bottom: 1.1rem;
}
.hero-title {
  font-family: 'Space Grotesk', sans-serif !important;
  font-size: clamp(2rem, 4.5vw, 3.4rem);
  font-weight: 800; line-height: 1.12; margin: 0 0 .9rem;
  background: linear-gradient(135deg, #f1f5f9 0%, #60a5fa 35%, #a78bfa 65%, #22d3ee 100%);
  -webkit-background-clip: text; -webkit-text-fill-color: transparent;
  background-clip: text;
  background-size: 200% 200%;
  animation: gradShift 6s ease infinite;
}
@keyframes gradShift {
  0%   { background-position: 0% 50%; }
  50%  { background-position: 100% 50%; }
  100% { background-position: 0% 50%; }
}
.hero-sub {
  font-size: 1rem; color: var(--t2); line-height: 1.65;
  max-width: 560px; margin: 0 auto;
}

/* ── Divider ── */
.hline {
  height: 1px;
  background: linear-gradient(90deg, transparent, var(--gb2) 30%, var(--gb2) 70%, transparent);
  margin: 1.8rem 0;
  border: none;
}

/* ── Section labels ── */
.sec-label {
  display: flex; align-items: center; gap: .6rem;
  font-family: 'Space Grotesk', sans-serif;
  font-size: 1rem; font-weight: 700; color: var(--t1);
  margin-bottom: .5rem;
}
.sec-label .num {
  width: 26px; height: 26px; border-radius: 7px;
  background: linear-gradient(135deg, var(--blue), var(--purple));
  display: inline-flex; align-items: center; justify-content: center;
  font-size: .75rem; font-weight: 800; color: #fff; flex-shrink: 0;
}
.sec-cap {
  font-size: .8rem; color: var(--t3); margin-bottom: .65rem; line-height: 1.5;
}

/* ── Upload zone ── */
[data-testid="stFileUploaderDropzone"] {
  background: var(--g1) !important;
  border: 2px dashed var(--gb) !important;
  border-radius: 18px !important;
  transition: all .25s ease !important;
}
[data-testid="stFileUploaderDropzone"]:hover {
  border-color: var(--blue) !important;
  background: rgba(59,130,246,.08) !important;
}
[data-testid="stFileUploaderDropzone"] > div { padding: 1.6rem !important; }
[data-testid="stFileUploaderDropzone"] span { color: var(--t2) !important; }
[data-testid="stFileUploaderDropzone"] small { color: var(--t3) !important; }
[data-testid="stFileUploaderDropzone"] button {
  background: rgba(59,130,246,.15) !important;
  border: 1px solid rgba(59,130,246,.35) !important;
  color: var(--blue-l) !important;
  border-radius: 9px !important;
  font-weight: 600 !important;
  box-shadow: none !important;
  animation: none !important;
  padding: .45rem 1.1rem !important;
}
[data-testid="stFileUploaderDropzone"] button:hover {
  background: rgba(59,130,246,.25) !important;
  transform: none !important;
}

/* Uploaded file chips */
[data-testid="stFileUploader"] [data-testid="stMarkdownContainer"] p,
[data-testid="uploadedFileName"] {
  background: rgba(59,130,246,.10) !important;
  border: 1px solid rgba(59,130,246,.2) !important;
  border-radius: 8px !important;
  padding: .25rem .6rem !important;
  font-size: .8rem !important;
  color: var(--blue-l) !important;
}

/* ── Process button ── */
.stButton > button {
  background: linear-gradient(135deg, #2563eb 0%, #7c3aed 60%, #0891b2 100%) !important;
  background-size: 200% 200% !important;
  border: none !important;
  border-radius: 14px !important;
  color: #fff !important;
  font-family: 'Space Grotesk', sans-serif !important;
  font-weight: 700 !important;
  font-size: 1.05rem !important;
  letter-spacing: .03em !important;
  padding: .85rem 2.4rem !important;
  transition: transform .25s ease, box-shadow .25s ease !important;
  box-shadow: 0 4px 24px rgba(59,130,246,.45), 0 2px 8px rgba(0,0,0,.4) !important;
  animation: btnGlow 4s ease-in-out infinite, gradShift 5s ease infinite !important;
}
@keyframes btnGlow {
  0%, 100% { box-shadow: 0 4px 24px rgba(59,130,246,.45), 0 2px 8px rgba(0,0,0,.4); }
  50%       { box-shadow: 0 6px 40px rgba(124,58,237,.7),  0 2px 8px rgba(0,0,0,.4); }
}
.stButton > button:hover {
  transform: translateY(-3px) scale(1.01) !important;
  box-shadow: 0 10px 50px rgba(59,130,246,.7), 0 4px 12px rgba(0,0,0,.5) !important;
  animation: none !important;
}
.stButton > button:active { transform: translateY(0) !important; }

/* ── Progress bar ── */
[data-testid="stProgress"] > div {
  background: rgba(255,255,255,.07) !important;
  border-radius: 100px !important;
  height: 5px !important;
  overflow: hidden;
}
[data-testid="stProgress"] > div > div {
  background: linear-gradient(90deg, var(--blue), var(--purple), var(--cyan), var(--blue)) !important;
  background-size: 300% 100% !important;
  border-radius: 100px !important;
  animation: shimmerBar 2s linear infinite !important;
}
@keyframes shimmerBar {
  0%   { background-position: 100% 0; }
  100% { background-position: -200% 0; }
}
[data-testid="stProgress"] ~ div p { color: var(--t2) !important; font-size: .85rem !important; }

/* ── Text input (password) ── */
.stTextInput input {
  background: var(--g1) !important;
  border: 1px solid var(--gb) !important;
  border-radius: 11px !important;
  color: var(--t1) !important;
  font-size: .9rem !important;
  transition: all .2s !important;
}
.stTextInput input:focus {
  border-color: var(--blue) !important;
  box-shadow: 0 0 0 3px rgba(59,130,246,.2) !important;
  outline: none !important;
}
.stTextInput label p { color: var(--t2) !important; font-size: .85rem !important; }

/* ── Radio ── */
[data-testid="stRadio"] > div { gap: .35rem !important; }
[data-testid="stRadio"] label {
  background: var(--g0) !important;
  border: 1px solid var(--gb) !important;
  border-radius: 10px !important;
  padding: .5rem .9rem !important;
  transition: all .2s !important;
  cursor: pointer !important;
}
[data-testid="stRadio"] label:hover {
  border-color: var(--blue) !important;
  background: rgba(59,130,246,.09) !important;
}
[data-testid="stRadio"] label[data-baseweb="radio"] span { color: var(--t2) !important; }

/* ── Alert boxes ── */
[data-testid="stAlert"] {
  border-radius: 12px !important;
  border-left-width: 3px !important;
}
.stSuccess { background: rgba(16,185,129,.1) !important; border-color: var(--green) !important; }
.stInfo    { background: rgba(59,130,246,.1)  !important; border-color: var(--blue) !important; }
.stWarning { background: rgba(245,158,11,.1)  !important; border-color: var(--amber) !important; }
.stError   { background: rgba(239,68,68,.1)   !important; border-color: var(--red) !important; }

/* ── Expander ── */
.stExpander {
  background: var(--g0) !important;
  border: 1px solid var(--gb) !important;
  border-radius: 14px !important;
  overflow: hidden;
}
[data-testid="stExpander"] summary {
  padding: .8rem 1.1rem !important;
  font-weight: 600 !important;
  color: var(--t2) !important;
}
[data-testid="stExpander"] summary:hover { color: var(--t1) !important; }

/* ── Download button ── */
.stDownloadButton button {
  background: linear-gradient(135deg, #059669 0%, #0891b2 100%) !important;
  background-size: 200% 200% !important;
  border: none !important;
  border-radius: 14px !important;
  color: #fff !important;
  font-family: 'Space Grotesk', sans-serif !important;
  font-weight: 700 !important;
  font-size: 1.05rem !important;
  letter-spacing: .03em !important;
  padding: .85rem 2.4rem !important;
  transition: transform .25s ease, box-shadow .25s ease !important;
  box-shadow: 0 4px 24px rgba(5,150,105,.45) !important;
}
.stDownloadButton button:hover {
  transform: translateY(-3px) !important;
  box-shadow: 0 10px 45px rgba(5,150,105,.65) !important;
}

/* ── Stats row ── */
.stats-row {
  display: flex; gap: 1rem; margin: 1.4rem 0;
  animation: fadeUp .5s ease both;
}
.stat-box {
  flex: 1; text-align: center;
  background: var(--g0);
  border: 1px solid var(--gb);
  border-radius: 16px;
  padding: 1.2rem 1rem;
  position: relative; overflow: hidden;
}
.stat-box::before {
  content: '';
  position: absolute; inset: 0;
  background: linear-gradient(135deg, rgba(255,255,255,.025) 0%, transparent 100%);
  pointer-events: none;
}
.stat-num {
  font-family: 'Space Grotesk', sans-serif;
  font-size: 2.1rem; font-weight: 800;
  background: linear-gradient(135deg, var(--blue-l), var(--purple-l));
  -webkit-background-clip: text; -webkit-text-fill-color: transparent;
  background-clip: text; line-height: 1.1;
}
.stat-lbl {
  font-size: .73rem; color: var(--t3);
  text-transform: uppercase; letter-spacing: .09em; margin-top: .3rem;
}

/* ── Result card ── */
.res-wrap { display: flex; flex-direction: column; gap: .45rem; margin: .5rem 0; }
.res-card {
  display: flex; align-items: center; gap: .9rem;
  background: var(--g0);
  border: 1px solid var(--gb);
  border-radius: 13px;
  padding: .8rem 1.1rem;
  animation: fadeUp .4s ease both;
  transition: all .2s ease;
  overflow: hidden;
  position: relative;
}
.res-card::before {
  content: '';
  position: absolute; left: 0; top: 0; bottom: 0; width: 3px;
  border-radius: 13px 0 0 13px;
}
.res-card.ok::before  { background: var(--green); }
.res-card.err::before { background: var(--red); }
.res-card.warn::before{ background: var(--amber); }
.res-card:hover {
  border-color: var(--gb2);
  transform: translateX(4px);
  background: var(--g1);
}
.res-ico  { font-size: 1.2rem; flex-shrink: 0; }
.res-name {
  flex: 1; font-size: .88rem; font-weight: 500; color: var(--t1);
  white-space: nowrap; overflow: hidden; text-overflow: ellipsis; min-width: 0;
}
.badge {
  flex-shrink: 0;
  font-size: .7rem; font-weight: 700; letter-spacing: .05em;
  padding: .2rem .6rem; border-radius: 6px; white-space: nowrap;
}
.b-blue   { background: rgba(59,130,246,.15); color: var(--blue-l); border: 1px solid rgba(59,130,246,.25); }
.b-purple { background: rgba(139,92,246,.15); color: var(--purple-l); border: 1px solid rgba(139,92,246,.25); }
.b-green  { background: rgba(16,185,129,.15); color: var(--green-l); border: 1px solid rgba(16,185,129,.25); }
.b-amber  { background: rgba(245,158,11,.15); color: var(--amber-l); border: 1px solid rgba(245,158,11,.25); }
.b-red    { background: rgba(239,68,68,.15);  color: var(--red-l);   border: 1px solid rgba(239,68,68,.25); }
.res-filled { font-size: .82rem; font-weight: 700; color: var(--green-l); flex-shrink: 0; min-width: 90px; text-align: right; }
.res-notes  { font-size: .77rem; color: var(--t3); flex-shrink: 0; max-width: 270px;
              white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }

/* ── Source log chips ── */
.src-log { display: flex; flex-direction: column; gap: .35rem; }
.src-row {
  display: flex; align-items: center; gap: .65rem;
  font-size: .82rem; color: var(--t2); padding: .3rem 0;
  border-bottom: 1px solid rgba(255,255,255,.04);
}
.src-row:last-child { border-bottom: none; }
.src-ico { font-size: 1rem; flex-shrink: 0; }
.src-fn  { flex: 1; color: var(--t1); font-weight: 500; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.src-msg { color: var(--t3); font-size: .78rem; }

/* ── AMC chip list ── */
.amc-chips { display: flex; flex-wrap: wrap; gap: .35rem; margin-top: .5rem; }
.amc-chip {
  font-size: .7rem; padding: .18rem .5rem; border-radius: 5px;
  background: var(--g1); border: 1px solid var(--gb); color: var(--t3);
}

/* ── Sidebar logo ── */
.sb-logo {
  display: flex; align-items: center; gap: .6rem;
  padding: .3rem 0 1rem;
  border-bottom: 1px solid var(--gb);
  margin-bottom: 1.2rem;
}
.sb-logo-icon { font-size: 1.7rem; }
.sb-logo-name {
  font-family: 'Space Grotesk', sans-serif;
  font-size: 1.05rem; font-weight: 700;
  background: linear-gradient(135deg, var(--blue-l), var(--purple-l));
  -webkit-background-clip: text; -webkit-text-fill-color: transparent;
  background-clip: text;
}
.sb-logo-ver {
  font-size: .65rem; color: var(--t3);
  background: var(--g1); border: 1px solid var(--gb);
  border-radius: 4px; padding: .1rem .35rem;
}
.sb-section-head {
  font-size: .72rem; font-weight: 700; letter-spacing: .1em;
  text-transform: uppercase; color: var(--t3);
  margin: 1rem 0 .5rem;
}

/* ── Animations ── */
@keyframes fadeDown { from { opacity:0; transform: translateY(-24px); } to { opacity:1; transform: translateY(0); } }
@keyframes fadeUp   { from { opacity:0; transform: translateY(14px);  } to { opacity:1; transform: translateY(0); } }

/* ── Scrollbar ── */
::-webkit-scrollbar { width: 5px; height: 5px; }
::-webkit-scrollbar-track { background: transparent; }
::-webkit-scrollbar-thumb { background: rgba(255,255,255,.12); border-radius: 100px; }
::-webkit-scrollbar-thumb:hover { background: rgba(255,255,255,.22); }

/* ── Hide Streamlit chrome ── */
#MainMenu, footer, header           { visibility: hidden !important; }
[data-testid="stToolbar"]           { display: none !important; }
[data-testid="stDecoration"]        { display: none !important; }
.viewerBadge_container__r5tak       { display: none !important; }
</style>
""", unsafe_allow_html=True)

# ── Sidebar ────────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("""
    <div class="sb-logo">
      <span class="sb-logo-icon">📊</span>
      <span class="sb-logo-name">BrokerageAI</span>
      <span class="sb-logo-ver">v2.0</span>
    </div>
    """, unsafe_allow_html=True)

    st.markdown('<div class="sb-section-head">🔑 Security</div>', unsafe_allow_html=True)
    pdf_password = st.text_input(
        "PDF Password",
        type="password",
        placeholder="AADCH3479E",
        help="Password for encrypted PDFs (AXIS, BOI). Leave blank if none.",
        label_visibility="collapsed",
    )
    st.caption("For encrypted PDFs (AXIS, BOI). Leave blank otherwise.")

    st.markdown('<div class="sb-section-head">🎯 Target Format</div>', unsafe_allow_html=True)
    vendor_fmt = st.radio(
        "Target file vendor format",
        ["Auto-detect", "Redos (column-based)", "Vijayinfotech (row-based)"],
        help=(
            "Redos: one row per scheme, trail values in columns.\n"
            "Vijayinfotech: one row per trail-year, fills T15 & B15.\n"
            "Auto-detect reads the file header automatically."
        ),
        label_visibility="collapsed",
    )

    st.markdown('<div class="sb-section-head">💡 How to use</div>', unsafe_allow_html=True)
    st.markdown("""
    <p style="font-size:.8rem;line-height:1.7;color:#94a3b8">
    1. Upload AMC source files (PDF/Excel)<br>
    2. Upload blank Excel templates<br>
    3. Click <strong style="color:#f1f5f9">Process &amp; Fill</strong><br>
    4. Download the filled ZIP
    </p>
    """, unsafe_allow_html=True)

    st.markdown('<div class="sb-section-head">🏦 Supported AMCs</div>', unsafe_allow_html=True)
    st.markdown("""
    <div class="amc-chips">
      <span class="amc-chip">ABSL</span><span class="amc-chip">AXIS</span>
      <span class="amc-chip">Bandhan</span><span class="amc-chip">BOI</span>
      <span class="amc-chip">Canara</span><span class="amc-chip">DSP</span>
      <span class="amc-chip">Franklin</span><span class="amc-chip">HDFC</span>
      <span class="amc-chip">HSBC</span><span class="amc-chip">Invesco</span>
      <span class="amc-chip">ICICI Pru</span><span class="amc-chip">LIC</span>
      <span class="amc-chip">Mahindra</span><span class="amc-chip">Mirae</span>
      <span class="amc-chip">Motilal</span><span class="amc-chip">Nippon</span>
      <span class="amc-chip">PGIM</span><span class="amc-chip">SBI</span>
      <span class="amc-chip">Sundaram</span><span class="amc-chip">TATA</span>
      <span class="amc-chip">Trust</span>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("""
    <p style="font-size:.72rem;color:#475569;margin-top:1.2rem;line-height:1.6">
    All values filled are <strong style="color:#64748b">EX-GST</strong> trail commissions.
    </p>
    """, unsafe_allow_html=True)

# ── Hero ───────────────────────────────────────────────────────────────────────
st.markdown("""
<div class="hero">
  <div class="hero-badge">⚡ &nbsp;AI-Powered &nbsp;·&nbsp; Automated &nbsp;·&nbsp; Accurate</div>
  <h1 class="hero-title">Mutual Fund<br>Brokerage Filler</h1>
  <p class="hero-sub">
    Upload AMC brokerage PDFs and blank Excel templates — get filled commission
    structures for <strong>Redos</strong> and <strong>Vijayinfotech</strong> formats instantly.
  </p>
</div>
<div class="hline"></div>
""", unsafe_allow_html=True)

# ── Upload columns ─────────────────────────────────────────────────────────────
col_src, col_tgt = st.columns(2, gap="large")

with col_src:
    st.markdown("""
    <div class="sec-label"><span class="num">1</span> Source Files</div>
    <div class="sec-cap">AMC brokerage PDFs or Excel files containing commission rates</div>
    """, unsafe_allow_html=True)
    src_uploads = st.file_uploader(
        "src", accept_multiple_files=True, type=["pdf", "xlsx"],
        key="src", label_visibility="collapsed",
    )

with col_tgt:
    st.markdown("""
    <div class="sec-label"><span class="num">2</span> Target Templates</div>
    <div class="sec-cap">Blank Excel files that need to be filled with trail values</div>
    """, unsafe_allow_html=True)
    tgt_uploads = st.file_uploader(
        "tgt", accept_multiple_files=True, type=["xlsx"],
        key="tgt", label_visibility="collapsed",
    )

st.markdown('<div class="hline"></div>', unsafe_allow_html=True)

# ── Guards ─────────────────────────────────────────────────────────────────────
if not src_uploads:
    st.info("⬆️  Upload at least one source file (AMC PDF or Excel) to begin.")
    st.stop()
if not tgt_uploads:
    st.info("⬆️  Upload at least one blank Excel template to begin.")
    st.stop()

# ── Process button ─────────────────────────────────────────────────────────────
run = st.button(
    "🚀  Process & Fill",
    type="primary",
    use_container_width=True,
)

if not run:
    st.stop()

# ── Step 1: Parse all source files ────────────────────────────────────────────
progress = st.progress(0, text="Initialising…")
source_data: dict[str, dict] = {}
source_log  = []

for i, f in enumerate(src_uploads):
    progress.progress(
        (i + 1) / (len(src_uploads) + len(tgt_uploads)),
        text=f"📄 Parsing source: {f.name}",
    )
    raw = f.read()
    ext = f.name.lower().rsplit('.', 1)[-1]

    preview = ''
    if ext == 'pdf':
        try:
            from parsers import read_pdf as _rpdf
            preview = _rpdf(raw, pdf_password)[:800]
        except Exception:
            pass

    amc = detect_source_amc(f.name, preview)
    if amc is None:
        source_log.append(('⚠️', f.name, '?', 'AMC not recognised – skipped'))
        continue

    parser_fn, needs_bytes = get_parser(amc)
    if parser_fn is None:
        source_log.append(('⚠️', f.name, amc, 'No parser found – skipped'))
        continue

    try:
        if not needs_bytes:
            data = parser_fn()
        elif amc == 'hdfc':
            existing = source_data.get('hdfc_raw', [])
            existing.append(raw)
            source_data['hdfc_raw'] = existing
            data = parser_fn(existing, pdf_password)
        else:
            data = parser_fn(raw, pdf_password) if ext == 'pdf' else parser_fn(raw)
        merged = source_data.get(amc, {})
        merged.update(data)
        source_data[amc] = merged
        source_log.append(('✅', f.name, amc.upper(), f'{len(data)} schemes parsed'))
    except Exception as e:
        source_log.append(('❌', f.name, amc.upper(), f'Error: {e}'))

# ABSL is always available (hardcoded rates)
if 'absl' not in source_data:
    from parsers import parse_absl
    source_data['absl'] = parse_absl()

# ── Step 2: Fill target files ──────────────────────────────────────────────────
filled_files: list[tuple[str, bytes]] = []
target_log   = []

# Merge all AMC dicts for vijayinfotech (single cross-AMC file)
merged_source: dict = {}
for _amc, _d in source_data.items():
    if isinstance(_d, list): continue
    merged_source.update(_d)

for i, f in enumerate(tgt_uploads):
    raw = f.read()
    try:
        wb = read_excel_wb(raw)
    except Exception as e:
        target_log.append(('❌', f.name, '–', '–', f'Cannot open: {e}'))
        continue

    if vendor_fmt == "Vijayinfotech (row-based)":
        fmt = 'vijay'
    elif vendor_fmt == "Redos (column-based)":
        fmt = 'redos'
    else:
        fmt = detect_format(wb)

    progress.progress(
        (len(src_uploads) + i + 1) / (len(src_uploads) + len(tgt_uploads)),
        text=(
            f"⏳ Filling Vijayinfotech (large file — please wait…): {f.name}"
            if fmt == 'vijay' else f"✍️ Filling: {f.name}"
        ),
    )

    if fmt == 'vijay':
        if not merged_source:
            target_log.append(('⚠️', f.name, 'ALL', '–', 'No source files uploaded'))
            continue
        try:
            filled, not_found = fill_workbook_vijay(wb, merged_source)
            buf = io.BytesIO(); wb.save(buf)
            filled_files.append((f.name, buf.getvalue()))
            total = len(filled) + len(not_found)
            target_log.append((
                '✅', f.name, 'Vijayinfotech',
                f'{len(filled)}/{total}',
                f'{len(not_found)} not matched' if not_found else 'All matched',
            ))
        except Exception:
            target_log.append(('❌', f.name, 'Vijayinfotech', '–', traceback.format_exc(limit=2)))
    else:
        amc = detect_target_amc(wb)
        if amc is None:
            target_log.append(('⚠️', f.name, '–', '–', 'AMC not recognised – skipped'))
            continue
        data = source_data.get(amc)
        if data is None:
            target_log.append(('⚠️', f.name, amc.upper(), '–',
                                'No source uploaded for this AMC'))
            continue
        try:
            filled, not_found = fill_workbook(wb, data)
            buf = io.BytesIO(); wb.save(buf)
            filled_files.append((f.name, buf.getvalue()))
            nf_str = ', '.join(not_found[:5]) + ('…' if len(not_found) > 5 else '')
            target_log.append((
                '✅', f.name, amc.upper(),
                f'{len(filled)}/{len(filled)+len(not_found)}',
                f'{len(not_found)} blank: {nf_str}' if not_found else 'All matched',
            ))
        except Exception:
            target_log.append(('❌', f.name, amc.upper(), '–', traceback.format_exc(limit=2)))

progress.empty()

# ── Results ────────────────────────────────────────────────────────────────────
# Stats
total_filled  = sum(int(r[3].split('/')[0]) for r in target_log if r[3] not in ('–', ''))
ok_count      = sum(1 for r in target_log if r[0] == '✅')
src_ok_count  = sum(1 for r in source_log  if r[0] == '✅')

st.markdown(f"""
<div class="stats-row">
  <div class="stat-box">
    <div class="stat-num">{len(filled_files)}</div>
    <div class="stat-lbl">Files Filled</div>
  </div>
  <div class="stat-box">
    <div class="stat-num">{total_filled:,}</div>
    <div class="stat-lbl">Schemes Matched</div>
  </div>
  <div class="stat-box">
    <div class="stat-num">{src_ok_count}</div>
    <div class="stat-lbl">Sources Parsed</div>
  </div>
  <div class="stat-box">
    <div class="stat-num">{ok_count}/{len(target_log)}</div>
    <div class="stat-lbl">Templates Filled</div>
  </div>
</div>
""", unsafe_allow_html=True)

# Target results cards
_AMC_BADGE = {
    'Vijayinfotech': 'b-purple',
    '–': 'b-amber',
    '?': 'b-amber',
}
_FALLBACK_BADGE_COLORS = [
    'b-blue','b-green','b-blue','b-blue','b-green','b-blue',
]

def _badge_cls(amc, icon):
    if amc in _AMC_BADGE: return _AMC_BADGE[amc]
    if icon == '❌': return 'b-red'
    if icon == '⚠️': return 'b-amber'
    return 'b-blue'

def _card_cls(icon):
    return 'ok' if icon == '✅' else ('err' if icon == '❌' else 'warn')

st.markdown('<div class="hline"></div>', unsafe_allow_html=True)
st.markdown("""
<div style="font-family:'Space Grotesk',sans-serif;font-size:1rem;font-weight:700;
            color:#f1f5f9;margin-bottom:.8rem;display:flex;align-items:center;gap:.5rem">
  📋 &nbsp;Fill Results
</div>
""", unsafe_allow_html=True)

cards_html = '<div class="res-wrap">'
for delay, (icon, fname, amc, filled_str, notes) in enumerate(target_log):
    bc = _badge_cls(amc, icon)
    cc = _card_cls(icon)
    cards_html += f"""
<div class="res-card {cc}" style="animation-delay:{delay*0.06:.2f}s">
  <span class="res-ico">{icon}</span>
  <span class="res-name" title="{fname}">{fname}</span>
  <span class="badge {bc}">{amc}</span>
  <span class="res-filled">{filled_str}</span>
  <span class="res-notes" title="{notes}">{notes}</span>
</div>"""
cards_html += '</div>'
st.markdown(cards_html, unsafe_allow_html=True)

# Source parse expander
if source_log:
    with st.expander("📂  Source files parsed", expanded=False):
        log_html = '<div class="src-log">'
        for icon, fname, amc, msg in source_log:
            log_html += f"""
<div class="src-row">
  <span class="src-ico">{icon}</span>
  <span class="src-fn" title="{fname}">{fname}</span>
  <span class="badge b-blue">{amc}</span>
  <span class="src-msg">{msg}</span>
</div>"""
        log_html += '</div>'
        st.markdown(log_html, unsafe_allow_html=True)

# ── Download ───────────────────────────────────────────────────────────────────
if filled_files:
    st.markdown('<div class="hline"></div>', unsafe_allow_html=True)
    zip_buf = io.BytesIO()
    with zipfile.ZipFile(zip_buf, 'w', zipfile.ZIP_DEFLATED) as zf:
        for fname, data in filled_files:
            zf.writestr(fname, data)
    zip_buf.seek(0)

    st.download_button(
        label=f"📥  Download {len(filled_files)} filled Excel file{'s' if len(filled_files) != 1 else ''}",
        data=zip_buf,
        file_name="brokerage_filled.zip",
        mime="application/zip",
        type="primary",
        use_container_width=True,
    )
elif target_log:
    st.warning("⚠️  No files were filled successfully. Check the results above.")
