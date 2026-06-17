"""
server.py  –  BrokerageAI FastAPI backend
Run:  python server.py   (or via run_app.bat)
"""
import io, uuid, zipfile, traceback, os
from fastapi import FastAPI, File, UploadFile, Form, HTTPException
from fastapi.responses import HTMLResponse, StreamingResponse, JSONResponse, FileResponse
from fastapi.staticfiles import StaticFiles
import uvicorn

from parsers import (
    detect_source_amc, detect_target_amc, get_parser,
    fill_workbook, fill_workbook_vijay, detect_format,
    read_excel_wb, parse_absl, clear_match_cache,
)
try:
    from parsers import read_pdf as _read_pdf
except ImportError:
    _read_pdf = None

app = FastAPI(title="BrokerageAI")
_downloads: dict[str, bytes] = {}          # token → zip bytes (cleared after download)

# Determine static dir (React dist/ next to server.py)
_BASE = os.path.dirname(os.path.abspath(__file__))
_DIST = os.path.join(_BASE, "dist")
_DIST_EXISTS = os.path.isdir(_DIST)

if _DIST_EXISTS:
    app.mount("/assets", StaticFiles(directory=os.path.join(_DIST, "assets")), name="assets")


@app.get("/", response_class=HTMLResponse)
async def serve_ui():
    if _DIST_EXISTS:
        return FileResponse(os.path.join(_DIST, "index.html"))
    with open(os.path.join(_BASE, "index.html"), encoding="utf-8") as f:
        return f.read()


@app.post("/api/process")
async def process(
    source_files: list[UploadFile] = File(default=[]),
    target_files: list[UploadFile]  = File(default=[]),
    pdf_password: str = Form(default=""),
    vendor_fmt:   str = Form(default="auto"),
):
    clear_match_cache()   # fresh norms for each request
    # ── Step 1: parse source files ────────────────────────────────────────
    source_data: dict[str, dict] = {}
    source_log:  list[dict]      = []

    for f in source_files:
        raw = await f.read()
        name = f.filename or ""
        ext  = name.lower().rsplit(".", 1)[-1] if "." in name else "pdf"

        preview = ""
        if ext == "pdf" and _read_pdf:
            try:
                preview = _read_pdf(raw, pdf_password)[:800]
            except Exception:
                pass

        amc = detect_source_amc(name, preview)
        if amc is None:
            source_log.append({"icon": "warn", "file": name, "amc": "?",
                                "msg": "AMC not recognised – skipped"})
            continue

        parser_fn, needs_bytes = get_parser(amc)
        if parser_fn is None:
            source_log.append({"icon": "warn", "file": name, "amc": amc,
                                "msg": "No parser available"})
            continue

        try:
            if not needs_bytes:
                data = parser_fn()
            elif amc == "hdfc":
                existing = source_data.get("hdfc_raw", [])
                existing.append(raw)
                source_data["hdfc_raw"] = existing
                data = parser_fn(existing, pdf_password)
            else:
                data = parser_fn(raw, pdf_password) if ext == "pdf" else parser_fn(raw)

            merged = source_data.get(amc, {})
            merged.update(data)
            source_data[amc] = merged
            source_log.append({"icon": "ok", "file": name, "amc": amc.upper(),
                                "msg": f"{len(data)} schemes parsed"})
        except Exception as e:
            source_log.append({"icon": "err", "file": name, "amc": amc.upper(),
                                "msg": f"Error: {e}"})

    # ABSL rates are always available (hardcoded)
    if "absl" not in source_data:
        source_data["absl"] = parse_absl()

    # Flat merge for vijayinfotech (all AMCs in one file)
    merged_source: dict = {}
    for _amc, _d in source_data.items():
        if isinstance(_d, list):
            continue
        merged_source.update(_d)

    # ── Step 2: fill target templates ─────────────────────────────────────
    filled_files: list[tuple[str, bytes]] = []
    target_log:   list[dict]              = []

    for f in target_files:
        raw  = await f.read()
        name = f.filename or "unknown.xlsx"

        try:
            wb = read_excel_wb(raw)
        except Exception as e:
            target_log.append({"icon": "err", "file": name, "amc": "–",
                                "filled": "–", "notes": f"Cannot open: {e}",
                                "status": "error"})
            continue

        if vendor_fmt == "vijay":
            fmt = "vijay"
        elif vendor_fmt == "redos":
            fmt = "redos"
        else:
            fmt = detect_format(wb)

        if fmt == "vijay":
            if not merged_source:
                target_log.append({"icon": "warn", "file": name, "amc": "ALL",
                                    "filled": "–", "notes": "No source files uploaded",
                                    "status": "warn"})
                continue
            try:
                filled, not_found = fill_workbook_vijay(wb, merged_source)
                buf = io.BytesIO(); wb.save(buf)
                filled_files.append((name, buf.getvalue()))
                total = len(filled) + len(not_found)
                # Show a sample of unmatched names so naming issues can be diagnosed
                sample = not_found[:8]
                notes = (f"{len(not_found)} not matched — sample: {' | '.join(sample)}"
                         if not_found else "All matched")
                target_log.append({
                    "icon": "ok", "file": name, "amc": "Vijayinfotech",
                    "filled": f"{len(filled)}/{total}",
                    "notes": notes,
                    "status": "ok",
                    "unmatched_sample": not_found[:20],
                })
            except Exception:
                target_log.append({"icon": "err", "file": name, "amc": "Vijayinfotech",
                                    "filled": "–", "notes": "Processing error",
                                    "status": "error"})
        else:
            amc = detect_target_amc(wb)
            if amc is None:
                target_log.append({"icon": "warn", "file": name, "amc": "–",
                                    "filled": "–", "notes": "AMC not recognised",
                                    "status": "warn"})
                continue
            data = source_data.get(amc)
            if data is None:
                target_log.append({"icon": "warn", "file": name, "amc": amc.upper(),
                                    "filled": "–", "notes": "No source uploaded for this AMC",
                                    "status": "warn"})
                continue
            try:
                filled, not_found = fill_workbook(wb, data)
                buf = io.BytesIO(); wb.save(buf)
                filled_files.append((name, buf.getvalue()))
                nf  = not_found[:3]
                nfs = ", ".join(nf) + ("…" if len(not_found) > 3 else "")
                target_log.append({
                    "icon": "ok", "file": name, "amc": amc.upper(),
                    "filled": f"{len(filled)}/{len(filled)+len(not_found)}",
                    "notes": f"{len(not_found)} blank: {nfs}" if not_found else "All matched",
                    "status": "ok",
                })
            except Exception:
                target_log.append({"icon": "err", "file": name, "amc": amc.upper(),
                                    "filled": "–", "notes": "Processing error",
                                    "status": "error"})

    # ── Build download ZIP ─────────────────────────────────────────────────
    zip_buf = io.BytesIO()
    with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for fname, data in filled_files:
            zf.writestr(fname, data)

    token = str(uuid.uuid4())
    _downloads[token] = zip_buf.getvalue()

    total_matched = sum(
        int(r["filled"].split("/")[0])
        for r in target_log
        if "/" in r.get("filled", "")
    )

    return JSONResponse({
        "token":         token if filled_files else None,
        "filled_count":  len(filled_files),
        "total_matched": total_matched,
        "source_log":    source_log,
        "target_log":    target_log,
    })


@app.get("/api/download/{token}")
async def download(token: str):
    data = _downloads.pop(token, None)
    if data is None:
        raise HTTPException(404, "Download link expired or not found")
    return StreamingResponse(
        io.BytesIO(data),
        media_type="application/zip",
        headers={"Content-Disposition": "attachment; filename=brokerage_filled.zip"},
    )


if __name__ == "__main__":
    uvicorn.run(app, host="0.0.0.0", port=8502, reload=False)
