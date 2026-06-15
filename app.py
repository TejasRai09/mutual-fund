"""
app.py  –  Mutual Fund Brokerage Structure Filler
Run:  streamlit run app.py
"""
import streamlit as st
import openpyxl, zipfile, io, traceback
from parsers import (
    detect_source_amc, detect_target_amc, get_parser,
    fill_workbook, fill_workbook_vijay, detect_format, read_excel_wb,
)

st.set_page_config(
    page_title="Brokerage Filler",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── Sidebar ────────────────────────────────────────────────────────────────────
with st.sidebar:
    st.title("⚙️ Settings")
    pdf_password = st.text_input(
        "PDF Password",
        type="password",
        help="Password for encrypted PDFs (e.g. AXIS, BOI). Leave blank if none.",
    )
    st.divider()
    st.markdown("""
**How to use**
1. Upload brokerage PDFs / Excel from AMCs
2. Upload blank Excel templates to fill
3. Click **Process**
4. Download the filled zip

**Supported AMCs**
ABSL · AXIS · Bandhan · BOI · Canara Robeco
DSP · Franklin · HDFC · HSBC · Invesco
ICICI Pru · LIC · Mahindra · Mirae · Motilal
Nippon · PGIM · SBI · Sundaram · TATA · Trust
""")
    st.divider()
    vendor_fmt = st.radio(
        "Target file vendor format",
        ["Auto-detect", "Redos (column-based)", "Vijayinfotech (row-based)"],
        help=(
            "Redos: one row per scheme, trail values in separate columns.\n"
            "Vijayinfotech: one row per (scheme × trail year), fills T15 & B15 columns.\n"
            "Auto-detect identifies the format from the file header."
        ),
    )
    st.divider()
    st.caption("Values filled are EX-GST trail commissions.")

# ── Main ───────────────────────────────────────────────────────────────────────
st.title("📊 Mutual Fund Brokerage Filler")
st.markdown("Upload AMC source files and blank Excel templates. Get filled Excel files instantly.")

col_src, col_tgt = st.columns(2, gap="large")

with col_src:
    st.subheader("① Source Files")
    st.caption("PDFs or Excel files from AMCs containing commission rates")
    src_uploads = st.file_uploader(
        "Drop source files here",
        accept_multiple_files=True,
        type=["pdf", "xlsx"],
        key="src",
        label_visibility="collapsed",
    )

with col_tgt:
    st.subheader("② Target Excel Templates")
    st.caption("Blank Excel files that need to be filled")
    tgt_uploads = st.file_uploader(
        "Drop target files here",
        accept_multiple_files=True,
        type=["xlsx"],
        key="tgt",
        label_visibility="collapsed",
    )

st.divider()

# ── Process ───────────────────────────────────────────────────────────────────
if not src_uploads:
    st.info("Upload at least one source file to begin.")
    st.stop()
if not tgt_uploads:
    st.info("Upload at least one target Excel file to begin.")
    st.stop()

if st.button("🚀 Process & Fill", type="primary", use_container_width=True):

    progress = st.progress(0, text="Starting…")

    # ── Step 1: Parse all source files ────────────────────────────────────────
    source_data: dict[str, dict] = {}   # {amc_id: {scheme: (t1,t2,t3,t4)}}
    source_log = []

    for i, f in enumerate(src_uploads):
        progress.progress((i + 1) / (len(src_uploads) + len(tgt_uploads)),
                          text=f"Parsing source: {f.name}")
        raw = f.read()
        ext = f.name.lower().rsplit('.', 1)[-1]

        # Get a text preview for AMC detection (PDF only)
        preview = ''
        if ext == 'pdf':
            try:
                from parsers import read_pdf as _rpdf
                preview = _rpdf(raw, pdf_password)[:800]
            except Exception:
                preview = ''

        amc = detect_source_amc(f.name, preview)

        if amc is None:
            source_log.append(('⚠️', f.name, 'unknown', 'AMC not recognised – skipped'))
            continue

        parser_fn, needs_bytes = get_parser(amc)
        if parser_fn is None:
            source_log.append(('⚠️', f.name, amc, 'No parser found – skipped'))
            continue

        try:
            if not needs_bytes:
                data = parser_fn()
            elif amc == 'hdfc':
                # HDFC may have multiple files (main + NFO)
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

    # ABSL is always available (hardcoded)
    if 'absl' not in source_data:
        from parsers import parse_absl
        source_data['absl'] = parse_absl()

    # ── Step 2: Fill target files ──────────────────────────────────────────────
    filled_files: list[tuple[str, bytes]] = []   # [(filename, xlsx_bytes)]
    target_log = []

    # Pre-merge all source dicts for vijayinfotech (covers all AMCs in one file)
    merged_source: dict = {}
    for _amc, _d in source_data.items():
        if isinstance(_d, list): continue   # skip hdfc_raw accumulator
        merged_source.update(_d)

    for i, f in enumerate(tgt_uploads):
        raw = f.read()
        try:
            wb = read_excel_wb(raw)
        except Exception as e:
            target_log.append(('❌', f.name, '–', '–', f'Cannot open: {e}'))
            continue

        # Determine vendor format
        if vendor_fmt == "Vijayinfotech (row-based)":
            fmt = 'vijay'
        elif vendor_fmt == "Redos (column-based)":
            fmt = 'redos'
        else:
            fmt = detect_format(wb)

        progress.progress(
            (len(src_uploads) + i + 1) / (len(src_uploads) + len(tgt_uploads)),
            text=(
                f"Filling Vijayinfotech (large file — please wait…): {f.name}"
                if fmt == 'vijay' else f"Filling: {f.name}"
            ),
        )

        if fmt == 'vijay':
            if not merged_source:
                target_log.append(('⚠️', f.name, 'ALL', '–', 'No source files uploaded'))
                continue
            try:
                filled, not_found = fill_workbook_vijay(wb, merged_source)
                buf = io.BytesIO()
                wb.save(buf)
                filled_files.append((f.name, buf.getvalue()))
                total = len(filled) + len(not_found)
                target_log.append((
                    '✅', f.name, 'Vijayinfotech',
                    f'{len(filled)}/{total} schemes',
                    f'{len(not_found)} not matched' if not_found else 'All schemes matched',
                ))
            except Exception as e:
                target_log.append(('❌', f.name, 'Vijayinfotech', '–', traceback.format_exc(limit=2)))

        else:
            # Redos: per-AMC matching
            amc = detect_target_amc(wb)
            if amc is None:
                target_log.append(('⚠️', f.name, '–', '–', 'AMC not recognised – skipped'))
                continue

            data = source_data.get(amc)
            if data is None:
                target_log.append(('⚠️', f.name, amc.upper(), '–',
                                    'No source file uploaded for this AMC'))
                continue

            try:
                filled, not_found = fill_workbook(wb, data)
                buf = io.BytesIO()
                wb.save(buf)
                filled_files.append((f.name, buf.getvalue()))
                nf_str = ', '.join(not_found[:5]) + ('…' if len(not_found) > 5 else '')
                target_log.append((
                    '✅', f.name, amc.upper(),
                    f'{len(filled)}/{len(filled)+len(not_found)}',
                    f'{len(not_found)} blank: {nf_str}' if not_found else 'All matched',
                ))
            except Exception as e:
                target_log.append(('❌', f.name, amc.upper(), '–', traceback.format_exc(limit=2)))

    progress.empty()

    # ── Results ────────────────────────────────────────────────────────────────
    st.success(f"Done! {len(filled_files)} file(s) filled.")

    # Source parse summary
    with st.expander("📂 Source files parsed", expanded=False):
        for icon, fname, amc, msg in source_log:
            st.markdown(f"{icon} **{fname}** `{amc}` — {msg}")

    # Target fill summary
    st.subheader("Results")
    cols = st.columns([0.05, 0.30, 0.12, 0.13, 0.40])
    cols[0].markdown("**#**"); cols[1].markdown("**File**")
    cols[2].markdown("**Format/AMC**"); cols[3].markdown("**Filled**"); cols[4].markdown("**Notes**")
    st.markdown("<hr style='margin:4px 0'>", unsafe_allow_html=True)

    for idx, (icon, fname, amc, filled_str, notes) in enumerate(target_log, 1):
        c = st.columns([0.05, 0.30, 0.12, 0.13, 0.40])
        c[0].write(f"{icon}")
        c[1].write(fname)
        c[2].write(amc)
        c[3].write(filled_str)
        c[4].write(notes)

    # ── Download ───────────────────────────────────────────────────────────────
    if filled_files:
        zip_buf = io.BytesIO()
        with zipfile.ZipFile(zip_buf, 'w', zipfile.ZIP_DEFLATED) as zf:
            for fname, data in filled_files:
                zf.writestr(fname, data)
        zip_buf.seek(0)

        st.divider()
        st.download_button(
            label=f"📥 Download {len(filled_files)} filled Excel file(s)",
            data=zip_buf,
            file_name="brokerage_filled.zip",
            mime="application/zip",
            type="primary",
            use_container_width=True,
        )
    else:
        st.warning("No files were filled. Check the logs above.")
