"""
parsers.py  –  AMC-specific brokerage parsers and matching utilities.
All parse_* functions accept raw bytes and return {scheme: (T1, T2, T3, T4+)}.
"""
import PyPDF2, openpyxl, re
from io import BytesIO

# ── PDF / Excel readers ───────────────────────────────────────────────────────

def read_pdf(data: bytes, pwd: str = '') -> str:
    r = PyPDF2.PdfReader(BytesIO(data))
    if r.is_encrypted:
        if not r.decrypt(''):       # Sundaram / open PDFs
            r.decrypt(pwd)
    return '\n'.join(p.extract_text() or '' for p in r.pages)

def read_excel_wb(data: bytes):
    """Load an xlsx workbook, stripping broken dataValidations if needed."""
    try:
        return openpyxl.load_workbook(BytesIO(data))
    except TypeError:
        # Some xlsx files (e.g. vijayinfotech) have malformed dataValidation
        # sqref values that crash openpyxl. Strip those elements and retry.
        import zipfile as _zf
        zin = _zf.ZipFile(BytesIO(data))
        buf = BytesIO()
        with _zf.ZipFile(buf, 'w', _zf.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                content = zin.read(item.filename)
                if (item.filename.startswith('xl/worksheets/')
                        and item.filename.endswith('.xml')):
                    # Strip dataValidations (may use x: namespace prefix)
                    content = re.sub(
                        rb'<(?:x:)?dataValidations\b[^>]*>.*?</(?:x:)?dataValidations>',
                        b'', content, flags=re.DOTALL,
                    )
                zout.writestr(item, content)
        buf.seek(0)
        return openpyxl.load_workbook(buf)

# ── Number helpers ─────────────────────────────────────────────────────────────

def get_floats(line):
    return [float(x) for x in re.findall(r'\d+\.\d+', line.replace('%', ''))]

def get_all_nums(line):
    return [float(x) for x in re.findall(r'\d+(?:\.\d+)?', line.replace('%', ''))]

# ── Name normalisation ─────────────────────────────────────────────────────────

def norm(s: str) -> str:
    s = str(s).lower().strip()
    # AMC name aliases
    s = s.replace('aditya birla sun life', 'absl')
    s = s.replace('icici prudential', 'icici pru')
    s = s.replace('reliance', 'nippon india')   # Reliance→Nippon (legacy vijay names)
    # Punctuation / separators
    s = re.sub(r'\s*&\s*', ' and ', s)
    s = re.sub(r"'", ' ', s)
    s = re.sub(r'-', ' ', s)
    s = re.sub(r'[^a-z0-9 ]', ' ', s)
    for old, new in [
        ('flexicap','flexi cap'),('midcap','mid cap'),('multicap','multi cap'),
        ('smallcap','small cap'),('largecap','large cap'),
        ('nifty100','nifty 100'),('nifty50','nifty 50'),('nifty500','nifty 500'),
        ('healthcare','health care'),
        ('fof','fund of fund'),
        ('savings fund','saving fund'),
        ('business groups','business group'),
        ('capital markets','capital market'),
        ('liquid plan','liquid fund'),
        (' cb fund',' corporate bond fund'),
        ('omni fund of fund','active fund of fund'),
    ]:
        s = s.replace(old, new)
    s = re.sub(r'\s+', ' ', s).strip()
    # Strip common plan/option suffixes that vijayinfotech appends
    s = re.sub(r'\b(regular|direct)\s+plan\b\s*', '', s).strip()
    s = re.sub(r'\b(regular|direct)\s+(growth|idcw|dividend)\b\s*', '', s).strip()
    s = re.sub(r'\bgrowth\s+option\b\s*', '', s).strip()
    s = re.sub(r'\s+', ' ', s).strip()
    s = re.sub(r'\b([a-z])(\s[a-z])+\b', lambda m: m.group(0).replace(' ', ''), s)
    return re.sub(r'\s+', ' ', s).strip()

# ── Fuzzy scheme matching ──────────────────────────────────────────────────────

def _strip_word(text, word):
    return re.sub(r'\s+', ' ', re.sub(rf'\b{word}\b\s*', '', text)).strip()

# Per-data-dict index cache (avoids re-computing norm() on source keys)
_match_idx_cache: dict = {}

def clear_match_cache():
    """Call this when source data changes between requests."""
    _match_idx_cache.clear()

def _get_match_idx(data: dict) -> dict:
    """Return (or build) a pre-normalized index for `data`."""
    did = id(data)
    if did not in _match_idx_cache:
        pairs_norm = [(k, norm(k), v) for k, v in data.items()]
        _match_idx_cache[did] = {
            'data':    data,
            'lower':   {k.lower(): v for k, v in data.items()},
            'norm':    {kn: v for _, kn, v in pairs_norm},
            'pairs':   [(kn, v) for _, kn, v in pairs_norm],
            'strip': {
                sw: [(_strip_word(kn, sw), v) for _, kn, v in pairs_norm]
                for sw in ('india', 'mf', 'elss')
            },
            'words':   [(frozenset(kn.split()), v) for _, kn, v in pairs_norm],
        }
    return _match_idx_cache[did]

def best_match(name, data):
    # O(1) exact lookups first
    if name in data: return data[name]
    nl = name.lower()
    idx = _get_match_idx(data)
    if nl in idx['lower']: return idx['lower'][nl]
    nn = norm(name)
    if nn in idx['norm']: return idx['norm'][nn]

    # O(n) partial match using pre-computed norms
    for kn, v in idx['pairs']:
        if nn in kn or kn in nn: return v

    # Strip words and try again
    for sw in ('india', 'mf', 'elss'):
        nn_s = _strip_word(nn, sw)
        for kn_s, v in idx['strip'][sw]:
            if nn_s == kn_s or nn_s in kn_s or kn_s in nn_s: return v
            if sw == 'elss':
                nw, kw = set(nn_s.split()), set(kn_s.split())
                if len(nw) >= 3 and len(kw) >= 3:
                    if nw.issubset(kw) and len(kw) - len(nw) <= 2: return v
                    if kw.issubset(nw) and len(nw) - len(kw) <= 2: return v

    # Word-subset match
    nw = frozenset(nn.split())
    for kw, v in idx['words']:
        if len(nw) >= 3 and len(kw) >= 3:
            if nw <= kw and len(kw) - len(nw) <= 2: return v
            if kw <= nw and len(nw) - len(kw) <= 2: return v

    return None

# ── Excel filler ──────────────────────────────────────────────────────────────

def fill_workbook(wb, data: dict):
    """Fill trail columns in-place. Returns (filled_names, not_found_names)."""
    ws = wb.active
    hdr = {ws.cell(1, c).value: c for c in range(1, ws.max_column + 1)}
    sc = hdr.get('SchemeName', 4)
    trail_order = [
        'Trail1stYear','T1','Trail2ndYear','T2','Trail3rdYear','T3',
        'Trail3rdYearOnwards','Trail4thYearOnwards','T4+',
        'Trail1stYearOnwards','Trail2ndYearOnwards',
    ]
    trail_cols = [(n, hdr[n]) for n in trail_order if n in hdr]
    filled, not_found = [], []
    for row in range(2, ws.max_row + 1):
        scheme = ws.cell(row, sc).value
        if not scheme: continue
        match = best_match(str(scheme), data)
        if match is not None:
            vals = match if isinstance(match, tuple) else (match,) * 4
            for i, (_, col_idx) in enumerate(trail_cols):
                if i < len(vals) and vals[i] is not None:
                    ws.cell(row, col_idx).value = round(float(vals[i]), 4)
            filled.append(str(scheme))
        else:
            not_found.append(str(scheme))
    return filled, not_found

# ── Vijayinfotech row-based format ────────────────────────────────────────────

_VIJAY_TRAIL_IDX = {
    'FIRST YEAR TRAIL':    0,
    'SECOND YEAR TRAIL':   1,
    'THIRD YEAR TRAIL':    2,
    'FOURTH YEAR TRAIL':   3,
    'LONGTERM YEAR TRAIL': 3,
}

def detect_format(wb) -> str:
    """Return 'vijay' if vijayinfotech row-based format, else 'redos'."""
    ws = wb.active
    headers = set()
    for c in range(1, min((ws.max_column or 0) + 1, 25)):
        v = ws.cell(1, c).value
        if v:
            headers.add(str(v))
    return 'vijay' if ('BrokerageName' in headers or 'T15' in headers) else 'redos'

def fill_workbook_vijay(wb, data: dict):
    """Fill vijayinfotech row-based xlsx in-place.
    One row per (scheme × trail-year); columns T15 and B15 are filled.
    Returns (filled_scheme_list, blank_scheme_list).
    """
    ws = wb.active
    hdr = {}
    for c in range(1, (ws.max_column or 0) + 1):
        v = ws.cell(1, c).value
        if v:
            hdr[str(v)] = c

    sc = hdr.get('Schemename', 6)       # scheme name
    jc = hdr.get('BrokerageName', 10)   # e.g. "FIRST YEAR TRAIL"
    nc = hdr.get('T15', 14)             # column to fill
    oc = hdr.get('B15', 15)             # column to fill

    # Pass 1 (read-only, fast): collect rows needing updates & build match cache
    cache = {}
    to_write = []   # (row_num, fill_value, scheme_name)

    for row_num, row in enumerate(
        ws.iter_rows(min_row=2, values_only=True), start=2
    ):
        scheme = row[sc - 1]
        bname  = row[jc - 1]
        if not scheme or not bname:
            continue

        bname_u = str(bname).upper()
        trail_idx = _VIJAY_TRAIL_IDX.get(bname_u)
        if trail_idx is None:
            continue    # CLAWBACK, UPFRONT, flat rows, etc.

        s = str(scheme).strip()
        if s not in cache:
            cache[s] = best_match(s, data)

        match = cache[s]
        if match is None:
            continue

        vals = match if isinstance(match, tuple) else (match,) * 4
        if trail_idx < len(vals) and vals[trail_idx] is not None:
            to_write.append((row_num, round(float(vals[trail_idx]), 4), s))

    # Pass 2 (write-only, small): update T15 and B15 for matched rows
    filled_set: set = set()
    for row_num, val, s in to_write:
        ws.cell(row_num, nc).value = val
        ws.cell(row_num, oc).value = val
        filled_set.add(s)

    blank_set = {s for s, m in cache.items() if m is None} - filled_set
    return list(filled_set), sorted(blank_set)

# ── AMC detection ──────────────────────────────────────────────────────────────

_SOURCE_HINTS = {
    'absl':     ['aditya birla sun life','aditya birla commission'],
    'axis':     ['axis brokerage','axis mutual fund commission','axis mf commission'],
    'bandhan':  ['bandhan brokerage','bandhan mutual fund commission'],
    'boi':      ['bank of india commission','boi commission'],
    'canara':   ['canara robeco'],
    'dsp':      ['dsp commission','dsp brokerage'],
    'ft':       ['franklin templeton','franklin india brokerage','ft brokerage'],
    'hdfc':     ['hdfc brokerage','hdfc mutual fund brokerage','hdfc mf brokerage'],
    'hsbc':     ['hsbc brokerage','hsbc mutual fund commission'],
    'invesco':  ['invesco brokerage'],
    'lic':      ['lic mf commission','lic mutual fund commission'],
    'mahindra': ['mahindra manulife commission','mahindra commission'],
    'mirae':    ['mirae asset brokerage','mirae brokerage'],
    'motilal':  ['motilal oswal brokerage'],
    'nippon':   ['nippon brokerage','nippon india commission'],
    'pgim':     ['pgim brokerage','pgim india commission'],
    'sbi':      ['sbi mf brokerage','sbi mutual fund brokerage'],
    'sundaram': ['sundaram brokerage'],
    'tata':     ['tata brokerage','tata mutual fund commission'],
    'trust':    ['trust brokerage','trustmf brokerage'],
    'old_bridge':['old bridge arbitrage','old bridge flexi','old bridge focused',
                  'old bridge fund commission','old bridge brokerage','old bridge mutual fund commission'],
    'icici':    ['icici prudential mutual fund commission','icici prudential brokerage',
                 'icici pru commission','icici prudential commission structure'],
    # New AMCs
    'bajaj':    ['bajaj finserv brokerage structure','bajaj finserv asset management brokerage'],
    'baroda':   ['baroda bnp paribas brokerage','baroda bnp brokerage','baroda bnp commission structure'],
    'edelweiss':['edelweiss brokerage structure','edelwiess brokerage structure'],
    'iti':      ['iti mf combined ongoing brokerage','iti brokerage structure','iti commission structure',
                 'iti ongoing brokerage'],
    'jm':       ['jm mf brokerage structure','jm financial brokerage structure','jm brokerage structure'],
    'quant':    ['quant mf brokerage','quant mutual fund brokerage','quant money managers'],
    'uti':      ['uti brokerage structure','uti commission structure','uti asset management company limited','uti quant fund'],
    'whiteoak': ['whiteoak brokerage structure','whiteoak capital brokerage','whiteoak commission structure'],
    '360one':   ['360 one brokerage structure','360 one asset management brokerage',
                 '360 one commission structure','360one brokerage'],
    'kotak':    ['kotak mahindra brokerage','kotak brokerage structure','kotak mf brokerage'],
    'groww':    ['groww commission structure','groww brokerage structure'],
}

_TARGET_PREFIXES = {
    'absl':     ['absl ','aditya birla'],
    'axis':     ['axis '],
    'bandhan':  ['bandhan '],
    'boi':      ['bank of india'],
    'canara':   ['canara robeco'],
    'dsp':      ['dsp '],
    'ft':       ['franklin '],
    'hdfc':     ['hdfc '],
    'hsbc':     ['hsbc '],
    'invesco':  ['invesco '],
    'lic':      ['lic mf','lic mutual'],
    'mahindra': ['mahindra manulife'],
    'mirae':    ['mirae asset'],
    'motilal':  ['motilal oswal'],
    'nippon':   ['nippon india'],
    'pgim':     ['pgim india'],
    'sbi':      ['sbi '],
    'sundaram': ['sundaram '],
    'tata':     ['tata '],
    'trust':    ['trustmf'],
    'icici':    ['icici pru','icici prudential'],
    'bajaj':    ['bajaj finserv'],
    'baroda':   ['baroda bnp paribas'],
    'edelweiss':['edelweiss '],
    'iti':      ['iti '],
    'jm':       ['jm '],
    'quant':    ['quant '],
    'uti':      ['uti '],
    'whiteoak': ['whiteoak'],
    '360one':   ['360 one'],
    'kotak':    ['kotak mahindra','kotak '],
    'groww':    ['groww '],
    'old_bridge':['old bridge'],
}

def detect_source_amc(filename: str, preview: str = '') -> str | None:
    fn_norm = filename.replace('_', ' ')
    text = (fn_norm + ' ' + preview[:800]).lower()
    fn   = fn_norm.lower()

    # Skip Fixed Deposit / FD-only files early
    if ('fixed deposit' in fn or (' fd ' in fn and 'brokerage' not in fn)
            or fn.startswith('fd ')):
        return None

    # Step 1: hint-based match (exact phrase in filename+preview)
    for amc, hints in _SOURCE_HINTS.items():
        if any(h in text for h in hints):
            return amc

    # Step 2: generic AMC-key-in-filename (legacy + unambiguous new AMCs)
    # Short keys (≤3 chars) must appear as whole word to avoid substring false-positives
    # e.g. 'iti' inside 'opportunities', 'dsp' inside 'dsp-123' is fine but 'iti' in 'entities' is not
    for amc in _SOURCE_HINTS:
        if amc not in fn: continue
        if len(amc) <= 3 and not re.search(r'(?<![a-z])' + re.escape(amc) + r'(?![a-z])', fn):
            continue
        return amc

    # Step 3: specific fallbacks for edge-case filenames
    if 'franklin' in fn:                return 'ft'
    if 'bank of india' in fn:           return 'boi'
    if 'old bridge' in fn or 'old btidge' in fn: return 'old_bridge'
    if 'bajaj finserv' in fn:           return 'bajaj'
    if 'bajaj' in fn and 'brokerage' in fn: return 'bajaj'
    if 'baroda bnp' in fn:              return 'baroda'
    if 'edelweis' in fn:                return 'edelweiss'
    if 'iti mf' in fn or ('iti' in fn and 'brokerage' in fn): return 'iti'
    if 'jm mf' in fn or ('jm' in fn and 'brokerage' in fn):   return 'jm'
    if 'quant' in fn and ('mf' in fn or 'brokerage' in fn):   return 'quant'
    if 'uti brokerage' in fn or 'uti commission' in fn:        return 'uti'
    if 'whiteoak' in fn:                return 'whiteoak'
    if '360 one' in fn or (fn.startswith('360 ') and 'brokerage' in fn): return '360one'
    if 'kotak' in fn and ('brokerage' in fn or 'commission' in fn): return 'kotak'
    if 'groww' in fn:                   return 'groww'
    return None

def detect_target_amc(wb) -> str | None:
    ws = wb.active
    hdr = {ws.cell(1, c).value: c for c in range(1, ws.max_column + 1)}
    sc = hdr.get('SchemeName', 4)
    for row in range(2, min(6, ws.max_row + 1)):
        scheme = str(ws.cell(row, sc).value or '').lower()
        for amc, prefixes in _TARGET_PREFIXES.items():
            if any(scheme.startswith(p) for p in prefixes):
                return amc
    return None

# ── Parsers ───────────────────────────────────────────────────────────────────

def parse_absl() -> dict:
    """ABSL: hardcoded EX-GST rates for April–June 2026."""
    raw = {
        # Equity
        'ABSL Banking And Financial Services Fund': 0.80,
        'ABSL Business Cycle Fund': 0.85,
        'ABSL Conglomerate Fund': 0.85,
        'ABSL Consumption Fund': 0.75,
        'ABSL Digital India Fund': 0.75,
        'ABSL Dividend Yield Fund': 0.90,
        'ABSL ESG Integration Strategy Fund': 1.00,
        'ABSL Flexi Cap Fund': 0.65,
        'ABSL Focused Fund': 0.75,
        'ABSL Infrastructure Fund': 0.95,
        'ABSL Large & Mid Cap Fund': 0.75,
        'ABSL Large Cap Fund': 0.65,
        'ABSL Manufacturing Equity Fund': 0.90,
        'ABSL Midcap Fund': 0.75,
        'ABSL MNC Fund': 0.80,
        'ABSL Multi-Cap Fund': 0.75,
        'ABSL Pharma & Healthcare Fund': 0.95,
        'ABSL PSU Equity Fund': 0.75,
        'ABSL Quant Fund': 0.85,
        'ABSL Small Cap Fund': 0.80,
        'ABSL Special Opportunities Fund': 0.95,
        'ABSL Transportation and Logistics Fund': 0.90,
        'ABSL Value Fund': 0.80,
        # Hybrid
        'ABSL Balanced Advantage Fund': 0.75,
        'ABSL Multi Asset Allocation Fund': 0.75,
        'ABSL Regular Savings Fund': 0.80,
        'ABSL Equity Savings Fund': 0.45,
        "ABSL Equity Hybrid '95 Fund": 0.75,
        # Liquid
        'ABSL Overnight Fund': 0.09,
        'ABSL Liquid Fund': 0.09,
        # Debt
        'ABSL Money Manager Fund': 0.08,
        'ABSL Floating Rate Fund': 0.15,
        'ABSL Savings Fund': 0.18,
        'ABSL Low Duration Fund': 0.65,
        'ABSL Corporate Bond Fund': 0.17,
        'ABSL Banking & PSU Debt Fund': 0.30,
        'ABSL Short Term Fund': 0.40,
        'ABSL Dynamic Bond Fund': 0.50,
        'ABSL Credit Risk Fund': 0.65,
        'ABSL Medium Term Plan': 0.65,
        'ABSL Income Fund': 0.40,
        'ABSL Long Duration Fund': 0.50,
        'ABSL Government Securities Fund': 0.45,
        # Arbitrage & Solution
        'ABSL Arbitrage Fund': 0.45,
        'ABSL Bal Bhavishya Yojna': 0.95,
        'ABSL Retirement Fund - The 30s Plan': 1.05,
        'ABSL Retirement Fund - The 40s Plan': 1.05,
        'ABSL Retirement Fund - The 50s Plan': 0.90,
        'ABSL Retirement Fund - The 50s Plus Debt Plan': 0.95,
        'ABSL ELSS Tax Saver Fund': 0.70,
        # Equity Index
        'ABSL Nifty 50 Index Fund': 0.22,
        'ABSL Nifty 50 Equal Weight Index Fund': 0.45,
        'ABSL Nifty Next 50 Index Fund': 0.35,
        'ABSL Nifty Midcap 150 Index Fund': 0.30,
        'ABSL Nifty Smallcap 50 Index Fund': 0.35,
        'ABSL Nifty India Defence Index Fund': 0.45,
        'ABSL BSE India Infrastructure Index Fund': 0.50,
        'ABSL BSE 500 Momentum 50 Index Fund': 0.50,
        'ABSL BSE 500 Quality 50 Index Fund': 0.50,
        # Debt Index
        'ABSL CRISIL-IBX Financial Services 3 to 6 Months Debt Index Fund': 0.15,
        'ABSL CRISIL-IBX Financial Services 9-12 Months Debt Index Fund': 0.15,
        'ABSL CRISIL IBX Gilt - April 2026 Index Fund': 0.12,
        'ABSL CRISIL IBX 60:40 SDL + AAA PSU Apr 2026 Index Fund': 0.15,
        'ABSL Nifty SDL Plus PSU Bond Sep 2026 60:40 Index Fund': 0.12,
        'ABSL CRISIL IBX AAA NBFC HFC Sep 2026 Index Fund': 0.20,
        'ABSL CRISIL-IBX AAA NBFC-HFC Index - Sep 2026 Fund': 0.20,
        'ABSL CRISIL IBX 60:40 SDL + AAA PSU - Apr 2027 Index Fund': 0.12,
        'ABSL Nifty SDL Apr 2027 Index Fund': 0.18,
        'ABSL CRISIL IBX Gilt June 2027 Index Fund': 0.15,
        'ABSL Nifty SDL Sep 2027 Index Fund': 0.18,
        'ABSL Crisil-IBX AAA Financial Services Index-Sep 2027 Fund': 0.15,
        'ABSL CRISIL IBX Gilt Apr 2028 Index Fund': 0.20,
        'ABSL CRISIL IBX 50:50 Gilt Plus SDL Apr 2028 Index Fund': 0.10,
        'ABSL CRISIL IBX Gilt Apr 2029 Index Fund': 0.15,
        'ABSL CRISIL IBX SDL Jun 2032 Index Fund': 0.18,
        'ABSL CRISIL IBX Gilt Apr 2033 Index Fund': 0.15,
        'ABSL Crisil IBX Gilt April 2033 Index Fund': 0.15,
        # FOF
        'ABSL Gold Fund': 0.20,
        'ABSL Silver ETF FOF': 0.25,
        'ABSL Multi Asset Passive FoF': 0.18,
        'ABSL Income Plus Arbitrage Active FoF': 0.20,
        'ABSL Conservative Hybrid Active FOF': 0.35,
        'ABSL Aggressive Hybrid Omni FOF': 0.45,
        'ABSL Dynamic Asset Allocation Omni FOF': 0.50,
        'ABSL Multi-Asset Omni FOF': 0.60,
    }
    return {k: (v, v, v) for k, v in raw.items()}


def parse_axis(data: bytes, pwd: str = '') -> dict:
    txt = read_pdf(data, pwd)
    result = {}
    skip_words = {'name', 'scheme', 'trail', 'year', 'fund', 'type', 'plan', 'category',
                  'commission', 'brokerage', 'structure', 'applicable', 'rate', 'total',
                  'addnl', 'special', 'n/a', 'nil'}
    for line in txt.split('\n'):
        line = line.strip()
        if not line: continue
        ll = line.lower()
        if not (ll.startswith('axis') or ('axis' in ll and any(
                w in ll for w in ['fund', 'etf', 'sip', 'elss']))): continue
        if ll.startswith('axis') and ll.split()[1:2] in [['mutual'], ['mf']]: continue
        nums = get_floats(line)
        if len(nums) >= 2:
            # Strip from first decimal number onwards (handles "N/A", "NIL" in-between)
            m_num = re.search(r'\s+\d+\.\d+', line)
            name = line[:m_num.start()].strip() if m_num else line
            name = re.sub(r'\s*\([^)]+\)\s*$', '', name).strip()  # strip trailing "(formerly...)"
            if not name or len(name) < 5: continue
            if name.lower().split()[0] not in ('axis',): continue
            result[name] = (nums[0], nums[0], nums[0], nums[-1])
    return result


def parse_bandhan(data: bytes, pwd: str = '') -> dict:
    txt = read_pdf(data, pwd)
    result = {}
    lines = txt.split('\n')

    # Category words shared across all three Bandhan formats
    _CAT = sorted([
        'Sectoral/Thematic Funds', 'Flexi Cap Fund', 'Focused Fund', 'ELSS',
        'Large & Mid Cap Fund', 'Large Cap Fund', 'Mid Cap Fund', 'Multi Cap Fund',
        'Small Cap Fund', 'Aggressive Hybrid Fund', 'Arbitrage Fund',
        'Balanced Advantage Fund', 'Conservative Hybrid Fund', 'Equity Savings Fund',
        'Multi Asset Allocation Fund', 'Multi Asset Allocation',
        'Low Duration Fund', 'Money Market Fund', 'Ultra Short Duration Fund',
        'Banking and PSU Fund', 'Corporate Bond Fund', 'Credit Risk Fund',
        'Dynamic Bond Fund', 'Floater Fund', 'Gilt Fund', 'Long Duration Fund',
        'Medium Duration Fund', 'Medium to Long Duration Fund', 'Short Duration Fund',
        'Liquid Fund', 'Overnight Fund', 'Index Fund', 'Fund of Fund',
        'Domestic Fund', 'Retirement Fund', 'Value Fund/Contra Fund', 'Value Fund',
    ], key=len, reverse=True)  # longest first so "Value Fund/Contra Fund" strips before "Value Fund"

    all_count    = sum(1 for l in lines if re.search(r'\bALL\b', l))
    may_4val     = sum(1 for l in lines if 'Bandhan' in l and len(get_floats(l)) >= 4)

    if all_count >= 3:
        # ── April format: "Bandhan Scheme Name ... ALL T1 T2 T3 T4 ..."
        for line in lines:
            line = line.strip()
            m_all = re.search(r'\bALL\b', line)
            if not line or not m_all: continue
            nums = get_floats(line)
            if len(nums) >= 4:
                name = line[:m_all.start()].strip()
                if name and len(name) > 5:
                    v = nums[0]
                    result[name] = (v, v, v, v)

    elif may_4val >= 5:
        # ── May format: "Bandhan Scheme Category T1 T2 T3 T4" (4 distinct columns)
        for line in lines:
            line = line.strip()
            nums = get_floats(line)
            if len(nums) >= 4:
                name = line
                for cw in _CAT:
                    name = name.replace(cw, '').strip()
                name = re.sub(r'[\d.%\s]+$', '', name).strip()
                if name and len(name) > 5:
                    result[name] = (nums[0], nums[1], nums[2], nums[3])

    else:
        # ── Jan-Mar format: "Bandhan Scheme Fund CategoryWord 1.25%" or "...1.20%1.05%"
        # Strip trailing values first, then strip category word from the end of what remains.
        for line in lines:
            line = line.strip()
            if not line.startswith('Bandhan'): continue
            nums = get_floats(line)
            if not nums: continue
            name = re.sub(r'[\d.%\s]+$', '', line).strip()
            for cw in _CAT:
                if name.endswith(cw):
                    name = name[:-len(cw)].strip()
                    break
            if not name or len(name) <= 5: continue
            if len(nums) >= 2:
                result[name] = (nums[0], nums[0], nums[0], nums[1])
            else:
                v = nums[0]
                result[name] = (v, v, v, v)

    return result


def parse_boi(data: bytes, pwd: str = '') -> dict:
    """BOI – IN-GST ('inclusive of statutory levies and taxes') → EX-GST after /1.18."""
    txt = read_pdf(data, pwd)
    result = {}
    lines = txt.split('\n')
    # Merge only when there are NO numbers at all (integer or decimal) on the line
    merged, i = [], 0
    while i < len(lines):
        line = lines[i].strip()
        if (line.startswith('Bank of India')
                and not get_all_nums(line)
                and i + 1 < len(lines)):
            merged.append(line + ' ' + lines[i+1].strip()); i += 2
        else:
            merged.append(line); i += 1
    for line in merged:
        line = re.sub(r'Click\s*[Hh]ere', '', line).strip()
        if not line.startswith('Bank of India'): continue
        nums = [n for n in get_all_nums(line) if 0.01 <= n <= 5.0]
        if len(nums) >= 1:
            name = re.sub(r'[\d.%\s]+$', '', line).strip()
            v = round(nums[0] / 1.18, 4)
            t2 = round(nums[1] / 1.18, 4) if len(nums) >= 2 else v
            result[name] = (v, v, v, t2)
    return result


def parse_canara(data: bytes, pwd: str = '') -> dict:
    txt = read_pdf(data, pwd)
    result = {}
    lines = txt.split('\n')
    merged, i = [], 0
    while i < len(lines):
        line = lines[i].strip()
        if line.startswith('Canara Robeco') and not get_floats(line) and i + 1 < len(lines):
            merged.append(line + ' ' + lines[i+1].strip()); i += 2
        else:
            merged.append(line); i += 1
    for line in merged:
        if not line.startswith('Canara Robeco'): continue
        nums = get_floats(line)
        if len(nums) >= 2:
            name = re.sub(r'[\d.%\s]+$', '', line)
            name = re.sub(r'\d.*', '', name).strip()
            if not name: continue
            # CANARA triplet: [Total_IN-GST, Base_EX-GST, GST_amount]
            # When exit load has decimal (e.g. "1.00%"), get_floats picks it up prepended.
            # nums[-2] always gives Base_EX-GST regardless of whether exit load is present.
            base = nums[-2]
            result[name] = (base,)
    return result


def parse_dsp(data: bytes, pwd: str = '') -> dict:
    txt = read_pdf(data, pwd)
    result = {}
    for line in txt.split('\n'):
        line = line.strip()
        if not line.startswith('DSP'): continue
        nums = get_floats(line)
        if len(nums) >= 6:
            name = re.sub(r'[\d.%\s]+$', '', line).strip()
            result[name] = (nums[0], nums[0], nums[0], nums[3])
        elif len(nums) >= 3:
            name = re.sub(r'[\d.%\s]+$', '', line).strip()
            result[name] = (nums[0], nums[0], nums[0], nums[0])
    return result


def parse_ft(data: bytes, pwd: str = '') -> dict:
    """Franklin Templeton – IN-GST ("inclusive of GST") → EX-GST after /1.18.
    Format: '0.00 0.90  1) FUND NAME (CODE)[#*] 0.90 0.90 ExitLoadText CATEGORY'
    The fund code (ALL-CAPS in parens) is the anchor; trail rates follow it.
    """
    txt = read_pdf(data, pwd)
    result = {}
    for line in txt.split('\n'):
        line = line.strip()
        # Primary: fund with code in parens: "N) FUND NAME (CODE)[#*] rate1 rate2 ..."
        m = re.search(
            r'(?<!\w)\d+\)\s+(.+?)\s*\([A-Z0-9]+\)[#*]*\s+([\d.]+)\s+([\d.]+)', line)
        if m:
            name = m.group(1).strip().rstrip('#*').strip()
            t1 = round(float(m.group(2)) / 1.18, 4)
            t2 = round(float(m.group(3)) / 1.18, 4)
            result[name] = (t1, t1, t1, t2)
            nu = name.upper()
            if 'INDEX FUND' in nu and 'NIFTY 50' in nu:
                result['FRANKLIN INDIA NSE NIFTY 50 INDEX FUND'] = (t1, t1, t1, t2)
        else:
            # Fallback: simpler "N)  FUND NAME  rate1 rate2"
            m2 = re.search(r'(?<!\w)\d+\)\s+([A-Z].+?)\s{2,}([\d.]+)(?:\s+([\d.]+))?', line)
            if m2:
                name = m2.group(1).strip().rstrip('#*').strip()
                t1 = round(float(m2.group(2)) / 1.18, 4)
                t2 = round(float(m2.group(3)) / 1.18, 4) if m2.group(3) else t1
                result[name] = (t1, t1, t1, t2)
    return result


def parse_hdfc(data_list: list, pwd: str = '') -> dict:
    """data_list: one or more HDFC PDF byte strings (main + optional NFO)."""
    txt = ''
    for d in data_list:
        try: txt += '\n' + read_pdf(d, pwd)
        except Exception: pass
    result = {}
    for line in txt.split('\n'):
        line = line.strip()
        if not line.startswith('HDFC'): continue
        nums = get_floats(line)
        if len(nums) >= 3:
            name = re.sub(r'\s*#.*', '', line).strip()
            name = re.sub(r'\s+(?:NIL|\d+\s*(?:Months?|Days?|Year|Yrs?|yr))\s*.*$', '', name, flags=re.I).strip()
            name = re.sub(r'[\d.%\s]+$', '', name).strip()
            if not name or len(name) < 4: continue
            result[name] = (nums[0], nums[0], nums[0], nums[1])
    return result


def parse_hsbc(data: bytes, pwd: str = '') -> dict:
    txt = read_pdf(data, pwd)
    result = {}
    lines = txt.split('\n')
    # Merge split lines: handles both "HSBC Scheme\n1.25 1.25 ..." (Apr-Jun)
    # and "1.25 1.25 HSBC Scheme 1.15 ..." (Jan-Mar) formats.
    merged, i = [], 0
    while i < len(lines):
        line = lines[i].strip()
        hi = line.find('HSBC')
        if hi >= 0 and not get_floats(line[hi:]) and i + 1 < len(lines):
            merged.append(line + ' ' + lines[i+1].strip()); i += 2
        else:
            merged.append(line); i += 1
    for line in merged:
        if 'HSBC' not in line: continue
        hi = line.find('HSBC')
        part = line[hi:]          # everything from "HSBC" onward; discards pre-name values
        nums = get_floats(part)
        if len(nums) >= 2:
            name = re.sub(r'[\d.%\s]+$', '', part).strip()
            if not name or len(name) < 4: continue
            result[name] = (nums[0], nums[0], nums[0], nums[1])
    return result


def parse_invesco(data: bytes, pwd: str = '') -> dict:
    txt = read_pdf(data, pwd)
    result = {}
    for line in txt.split('\n'):
        line = line.strip()
        if 'Invesco' not in line: continue
        nums = get_floats(line)
        # Apr-Jun: 4 columns (T1, T2, T3, T4+);  Jan-Mar: 3 columns (T1, T2&T3, T4+)
        if len(nums) >= 3:
            idx = line.find('Invesco')
            part = line[idx:]
            name = re.sub(r'[\d.\s]+$', '', part).strip()
            if not name: continue
            result[name] = (nums[0], nums[0], nums[0], nums[-1])
    return result


def parse_lic(data: bytes, pwd: str = '') -> dict:
    txt = read_pdf(data, pwd)
    # Fix PDF split-decimal artifacts: "1. 14" → "1.14", "1 .27" → "1.27"
    txt = re.sub(r'(\d)\.\s+(\d)', r'\1.\2', txt)
    txt = re.sub(r'(\d)\s+\.(\d)', r'\1.\2', txt)
    result = {}
    cat_re = re.compile(
        r'\b(FLEXI CAP FUND|LARGE CAP FUND|LARGE & MIDCAP FUND|CHILDREN\'?S FUND|'
        r'MULTI ?CAP FUND|MID ?CAP FUND|SMALL CAP FUND|EQUITY DIVIDEND YIELD|'
        r'EQUITY FOCUSED FUND|EQUITY VALUE FUND|SECTORAL/?THEMATIC FUND|'
        r'EQUITY SAVINGS FUND|ARBITRAGE FUND|DYNAMIC ASSET ALLOCAT\w*|ELSS|'
        r'INDEX FUND|GOLD FUND|MEDIUM TO LONG DURATION FUND|MONEY MARKET FUND|'
        r'BANKING & PSU DEBT FUND|GILT FUND|LOW DURATION FUND|SHORT DURATION FUND|'
        r'OVERNIGHT FUND|ULTRA SHORT DURATION\s+FUND|LIQUID FUND|'
        r'CONSERVATIVE HYBRID\s+FUND|AGGRESSIVE HYBRID FUND)\b', re.I)
    for line in txt.split('\n'):
        line = line.strip()
        if not line.startswith('LIC'): continue
        nums = get_floats(line)
        if len(nums) >= 4:
            name = re.sub(r'[\d.\s]+$', '', line).strip()
            name = re.sub(r'\s*See\s+Overleaf\s*', '', name, flags=re.I).strip()
            # Strip only the LAST occurrence of a category word (PDF category column).
            matches = list(cat_re.finditer(name))
            if len(matches) >= 2:
                name = name[:matches[-1].start()].strip()
            if not name: continue
            # LIC rates are IN-GST ("brokerage shall be inclusive of GST") → convert to EX-GST
            result[name] = tuple(round(n / 1.18, 4) for n in nums[:4])
        elif len(nums) >= 2:
            name = re.sub(r'[\d.\s]+$', '', line).strip()
            name = re.sub(r'\s*See\s+Overleaf\s*', '', name, flags=re.I).strip()
            matches = list(cat_re.finditer(name))
            if len(matches) >= 2:
                name = name[:matches[-1].start()].strip()
            if not name: continue
            v0, vn = round(nums[0] / 1.18, 4), round(nums[-1] / 1.18, 4)
            result[name] = (v0, v0, v0, vn)
    return result


def parse_mahindra(data: bytes, pwd: str = '') -> dict:
    """Mahindra Manulife – IN-GST ("inclusive of any tax, GST") → EX-GST after /1.18."""
    txt = read_pdf(data, pwd)
    result = {}
    cat_words = ['ELSS (Tax Saver)','Large-Cap','Mid-Cap','Small Cap','Large & Mid Cap',
                 'Multi-Cap','Flexi Cap','Focused','Thematic','Value Fund','Sectoral Fund',
                 'Equity Savings','Balanced Advantage','Aggressive Hybrid','Hybrid',
                 'Fund of Funds','FOF Domestic','Dynamic Bond','Liquid','Debt',
                 'Low Duration','Short Duration']
    for line in txt.split('\n'):
        line = line.strip()
        if not line.startswith('Mahindra'): continue
        nums = get_floats(line)
        if len(nums) >= 2:
            name = re.sub(r'[\d.\s%]+$', '', line).strip()
            for w in sorted(cat_words, key=len, reverse=True):
                name = name.replace(w, '').strip()
            name = re.sub(r'\s{2,}', ' ', name).strip()
            if not name or len(name) < 5: continue
            v0, vn = round(nums[0] / 1.18, 4), round(nums[-1] / 1.18, 4)
            result[name] = (v0, v0, v0, vn)
    return result


def parse_mirae(data: bytes, pwd: str = '') -> dict:
    txt = re.sub(r'\b00\.', '0.', read_pdf(data, pwd))
    result = {}
    lines = txt.split('\n')
    i = 0
    while i < len(lines):
        line = lines[i].strip()
        if line.lower().startswith('mirae asset'):
            combined = line
            j = i + 1
            while j < len(lines) and j < i + 25:
                combined += ' ' + lines[j].strip(); j += 1
            nums = get_floats(combined)
            if len(nums) >= 5:
                name = re.sub(r'[\d.%±\s]+$', '', combined).strip()
                name = re.sub(r'\d.*', '', name).strip()
                result[name] = (nums[1], nums[2], nums[3], nums[4])
            elif len(nums) >= 1:
                name = re.sub(r'[\d.%±\s]+$', '', line).strip()
                if name: result[name] = (nums[0], nums[0], nums[0], nums[0])
        i += 1
    return result


def parse_motilal(data: bytes, pwd: str = '') -> dict:
    """Motilal Oswal – IN-GST ('inclusive of applicable GST') → EX-GST after /1.18.
    Format: 'Motilal Oswal Scheme Name  T1 T2 T3 T4+ Cumulative' (BPS)
    Fund names may embed numbers (Nifty 150, Nifty 500) → always use last 5 values.
    """
    txt = read_pdf(data, pwd)
    result = {}
    for line in txt.split('\n'):
        line = line.strip()
        if not line.startswith('Motilal'): continue
        nums = get_all_nums(line)
        if len(nums) >= 5:
            name = re.sub(r'[\d.\s]+$', '', line).strip()
            if not name or len(name) < 5: continue
            t = nums[-5:]   # last 5 = [T1, T2, T3, T4+, cumulative]
            t1 = round((t[0] / 100) / 1.18, 4)
            t4 = round((t[3] / 100) / 1.18, 4)
            result[name] = (t1, t1, t1, t4)
        elif len(nums) >= 2:
            name = re.sub(r'[\d.\s]+$', '', line).strip()
            if not name or len(name) < 5: continue
            t1 = round((nums[-2] / 100) / 1.18, 4)
            t4 = round((nums[-1] / 100) / 1.18, 4)
            result[name] = (t1, t1, t1, t4)
    return result


def parse_nippon(data: bytes, pwd: str = '') -> dict:
    txt = read_pdf(data, pwd)
    result = {}
    for line in txt.split('\n'):
        line = line.strip()
        if 'NIPPON INDIA' not in line.upper(): continue
        nums = get_floats(line)
        if len(nums) >= 2:
            m = re.search(r'NIPPON\s+INDIA', line, re.I)
            if m:
                part = line[m.start():]
                name = re.sub(r'\s+(?:NIL|\d+\s*(?:Month|Day|Year|Yr|yr|lock))[^\d]*.*$', '', part, flags=re.I).strip()
                name = re.sub(r'[\d.%\s]+$', '', name).strip()
            else:
                name = re.sub(r'[\d.%\s]+$', '', line).strip()
            if not name: continue
            # T1=T2=T3 use first rate, T4+ use last (may differ in newer PDFs)
            result[name] = (nums[0], nums[0], nums[0], nums[-1])
            nu = name.upper()
            if 'POWER' in nu and 'INFRA' in nu:
                alias = re.sub(r'POWER\s*(?:AND|&)\s*INFRA', 'INFRASTRUCTURE', nu)
                result[alias] = result[name]
            if nu == 'NIPPON INDIA FOCUSED FUND':
                result['NIPPON INDIA FOCUSED LARGE CAP FUND'] = result[name]
            if 'MULTI ASSET FUND' in nu and 'OMNI' not in nu:
                result[nu.replace('MULTI ASSET FUND', 'MULTI ASSET ALLOCATION FUND')] = result[name]
    return result


def parse_pgim(data: bytes, pwd: str = '') -> dict:
    txt = read_pdf(data, pwd)
    result = {}
    for line in txt.split('\n'):
        line = line.strip()
        if not line.startswith('PGIM'): continue
        nums = get_floats(line)
        name = re.sub(r'[\d.%\s]+$', '', line).strip()
        name = re.sub(r'\d+\.?\d*%.*', '', name).strip()
        if not name: continue
        if len(nums) >= 4:
            # PGIM columns: [Total_Y1, Base, Add_Trail, T4+]
            # Equity funds prepend exit load (e.g. "0.50%") → 5 floats total.
            # Use the last 4 to always get [Total_Y1, Base, Add, T4+] correctly.
            t = nums[-4:]
            result[name] = (t[0], t[0], t[0], t[3])
        elif len(nums) == 3:
            result[name] = (nums[0], nums[0], nums[0], nums[0])
    return result


def parse_sbi(data: bytes, pwd: str = '') -> dict:
    """SBI – supports both PDF and Excel (.xlsx) source files.
    Excel layout: col2=Category, col3=SchemeCode, col4=SchemeName,
                  col5=1stYear, col6=2nd&3rdYear, col7=3YrPricing, col8=4thYearOnwards
    Values in Excel are already in % (e.g. 0.67 means 0.67%).
    """
    if data[:2] == b'PK':
        return _parse_sbi_xl(data)
    txt = read_pdf(data, pwd)
    result = {}
    for line in txt.split('\n'):
        line = line.strip()
        if 'SBI' not in line: continue
        idx = line.find('SBI ')
        if idx < 0: continue
        part = line[idx:]
        nums = get_floats(part)
        if len(nums) >= 2:
            name = re.sub(r'[\d.%\s]+$', '', part).strip()
            name = re.sub(r'\s+\d+\.?\d*\s*(?:%|years?|yrs?|months?|days?).*$', '', name, flags=re.I).strip()
            if not name or not name.startswith('SBI'): continue
            result[name] = (nums[0], nums[1])
    return result

def _parse_sbi_xl(data: bytes) -> dict:
    wb = read_excel_wb(data)
    ws = wb.active
    result = {}
    def cv(x): return round(float(x), 4) if isinstance(x, (int, float)) else None
    for row in range(1, ws.max_row + 1):
        name = ws.cell(row, 4).value  # Scheme Name in col 4
        t1   = ws.cell(row, 5).value  # 1st Year Trail
        t2   = ws.cell(row, 6).value  # 2nd & 3rd Year Trail
        t4   = ws.cell(row, 8).value  # 4th Year Onwards
        if not isinstance(name, str): continue
        name = name.strip()
        if not name.startswith('SBI'): continue
        if not isinstance(t1, (int, float)): continue
        result[name] = (cv(t1), cv(t2), cv(t2), cv(t4))
    return result


def parse_sundaram(data: bytes, pwd: str = '') -> dict:
    txt = read_pdf(data, pwd)   # empty password handled in read_pdf
    result = {}
    for line in txt.split('\n'):
        line = line.strip()
        if not line.startswith('Sundaram'): continue
        nums = get_floats(line)
        if len(nums) >= 1:
            name = re.sub(r'[\d.\s*]+$', '', line).strip().rstrip('*').strip()
            if not name: continue
            result[name] = (nums[0],)
            expanded = name.replace('Fin. Services', 'Financial Services').replace('Opps', 'Opportunities')
            if expanded != name: result[expanded] = (nums[0],)
    return result


def parse_tata(data: bytes, pwd: str = '') -> dict:
    txt = read_pdf(data, pwd)
    result = {}
    for line in txt.split('\n'):
        line = line.strip()
        if not line.startswith('Tata'): continue
        nums = get_floats(line)
        if len(nums) >= 1:
            name = re.sub(r'[\d.%\s]+$', '', line).strip()
            if not name: continue
            name = (name
                    .replace('Tata Retirement Savings Fund -PP', 'Tata Retirement Saving Fund - Progressive')
                    .replace('Tata Retirement Savings Fund -MP', 'Tata Retirement Saving Fund - Moderate')
                    .replace('Tata Retirement Savings Fund -CP', 'Tata Retirement Saving Fund - Conservative'))
            result[name] = (nums[0],)
    return result


def parse_trust(data: bytes, pwd: str = '') -> dict:
    """TRUSTMF – IN-GST ('commission rates are inclusive of GST') → EX-GST after /1.18.
    Format A (recent): each line starts with 'TRUSTMF ...'
    Format B (Q1): 'Category  TRUSTMF Scheme Name  rate  rate'
    """
    txt = read_pdf(data, pwd)
    result = {}
    for line in txt.split('\n'):
        line = re.sub(r'\s+', ' ', line.strip())
        # Find TRUSTMF anywhere in the line
        idx = line.upper().find('TRUSTMF')
        if idx < 0: continue
        part = line[idx:]
        nums = get_floats(part)
        if len(nums) >= 1:
            name = re.sub(r'[\d.%\s]+$', '', part).strip()
            if not name or len(name) < 5: continue
            v = round(nums[0] / 1.18, 4)
            result[name] = (v, v, v, v)
    return result


def parse_icici(data: bytes, pwd: str = '') -> dict:
    """ICICI source is an Excel file.
    Layout: col1=SchemeName, col2=T1, col3=T2, col4=T3, col5=T4+
    Values are decimal fractions (0.0091 = 0.91%) → multiply by 100.
    NFO PDFs may also be passed — return empty dict for those.
    """
    if data[:2] != b'PK':   # Not a zip/xlsx → skip PDFs gracefully
        return {}
    wb = read_excel_wb(data)
    ws = wb.active
    result = {}
    def cv(x): return round(float(x) * 100, 4) if isinstance(x, (int, float)) else None
    for row in range(1, ws.max_row + 1):
        name = ws.cell(row, 1).value
        t1   = ws.cell(row, 2).value
        t2   = ws.cell(row, 3).value
        t3   = ws.cell(row, 4).value
        t4   = ws.cell(row, 5).value
        if not isinstance(name, str): continue
        name = name.strip()
        if not name.startswith('ICICI'): continue
        if not isinstance(t1, (int, float)): continue
        vals = (cv(t1), cv(t2), cv(t3), cv(t4))
        result[name] = vals
        ku = name.upper()
        if 'REGULAR GOLD SAVINGS' in ku:
            result['ICICI Prudential Gold ETF FOF'] = vals
            result['ICICI Prudential Gold ETF'] = vals
    return result


_BAJAJ_TYPES = sorted([
    'Flexi Cap Fund','Liquid Fund','Overnight Fund','Money Market Fund','Banking & PSU Fund',
    'Arbitrage Fund','Large Cap Fund','Mid Cap Fund','Small Cap Fund','ELSS Tax Saver Fund',
    'Multi Asset Allocation Fund','Balanced Advantage Fund','Low Duration Fund',
    'Short Duration Fund','Ultra Short Duration Fund','Conservative Hybrid Fund',
    'Multi Cap Fund','Dynamic Bond Fund','Focused Fund','Equity Savings Fund',
    'Value Fund','Small Cap Fund Regular',
], key=len, reverse=True)

def parse_bajaj(data: bytes, pwd: str = '') -> dict:
    """Bajaj Finserv – IN-GST rates. Format: 'Bajaj Finserv X Fund [Cat] [ExitPeriod] T% T% T% Total%'
    Some PDF layouts split the scheme name across two lines; merge those before parsing.
    """
    txt = read_pdf(data, pwd)
    result = {}
    lines = txt.split('\n')
    merged, i = [], 0
    while i < len(lines):
        line = lines[i].strip()
        if line.startswith('Bajaj Finserv') and not get_floats(line) and i + 1 < len(lines):
            merged.append(line + ' ' + lines[i+1].strip()); i += 2
        else:
            merged.append(line); i += 1
    for line in merged:
        line = line.strip()
        if not line.startswith('Bajaj Finserv'): continue
        nums = get_floats(line)
        if len(nums) < 3: continue
        # Find where first float starts → everything before it is name + category + exit period
        m = re.search(r'\d+\.\d+', line)
        if not m: continue
        pre = line[:m.start()].strip()
        # Strip exit-load period (e.g. "6 Months", "7 days", "NIL", "15 Days")
        pre = re.sub(r'\s+(?:NIL|\d+\s*(?:months?|days?|years?|day))\s*$', '', pre, flags=re.I).strip()
        # Strip duplicate category suffix (e.g. "Flexi Cap Fund" repeated)
        for t in _BAJAJ_TYPES:
            if pre.endswith(' ' + t):
                pre = pre[:-len(t)-1].strip()
                break
        if not pre or not pre.startswith('Bajaj Finserv') or len(pre) < 14: continue
        # nums[0]=T1=T2=T3, nums[-1]=Total(3yr, ignore)
        v = round(nums[0] / 1.18, 4)
        result[pre] = (v, v, v, v)
    return result


def parse_baroda(data: bytes, pwd: str = '') -> dict:
    """Baroda BNP Paribas – EX-GST. Format: 'Baroda BNP Paribas X Fund T1% T4+% Total%'"""
    txt = read_pdf(data, pwd)
    result = {}
    lines = txt.split('\n')
    merged, i = [], 0
    while i < len(lines):
        line = lines[i].strip()
        if line.startswith('Baroda BNP Paribas') and not get_floats(line) and i + 1 < len(lines):
            merged.append(line + ' ' + lines[i+1].strip()); i += 2
        else:
            merged.append(line); i += 1
    for line in merged:
        if not line.startswith('Baroda BNP Paribas'): continue
        nums = get_floats(line)
        if len(nums) >= 2:
            name = re.sub(r'[\d.%\s]+$', '', line).strip()
            if not name: continue
            result[name] = (nums[0], nums[0], nums[0], nums[1])
    return result


def parse_edelweiss(data: bytes, pwd: str = '') -> dict:
    """Edelweiss – IN-GST. Format: 'Edelweiss X Fund [ExitLoad] T1% T2% T3%'"""
    txt = read_pdf(data, pwd)
    result = {}
    for line in txt.split('\n'):
        line = line.strip()
        if 'Edelweis' not in line: continue
        idx = line.find('Edelweis')
        part = line[idx:]
        nums = get_floats(part)
        if not nums: continue
        # Strip trailing numbers to get name; also strip exit load description
        name = re.sub(r'[\d.%\s]+$', '', part).strip()
        name = re.sub(r'\s+Exit\s+load.*$', '', name, flags=re.I).strip()
        if not name or len(name) < 5: continue
        v = round(nums[0] / 1.18, 4)
        result[name] = (v, v, v, v)
    return result


def parse_iti(data: bytes, pwd: str = '') -> dict:
    """ITI MF – EX-GST. Format: '[Category] ITI X Fund T1% T1%'"""
    txt = read_pdf(data, pwd)
    result = {}
    for line in txt.split('\n'):
        line = line.strip()
        if 'ITI ' not in line: continue
        idx = line.find('ITI ')
        part = line[idx:]
        nums = get_floats(part)
        if not nums: continue
        name = re.sub(r'[\d.%\s]+$', '', part).strip()
        if not name or len(name) < 5: continue
        v = nums[0]
        result[name] = (v, v, v, v)
    return result


def parse_jm(data: bytes, pwd: str = '') -> dict:
    """JM Financial – IN-GST. Format: '[Category]JM X Fund [ExitLoad%] T1 T2'"""
    txt = read_pdf(data, pwd)
    result = {}
    # Split entire text at each "JM " boundary to handle PDF-concatenated rows
    for frag in re.split(r'(?=\bJM\s)', txt):
        frag = frag.strip()
        if not frag.startswith('JM'): continue
        nums = get_floats(frag)
        if len(nums) < 2: continue
        name = re.sub(r'[\d.%\s]+$', '', frag).strip()
        name = re.sub(r'\s*(?:NIL|refer\s+link\s+below).*$', '', name, flags=re.I).strip()
        name = re.sub(r'[*#(].*$', '', name).strip()
        if not name or len(name) < 5: continue
        t1 = round(nums[-2] / 1.18, 4)
        t2 = round(nums[-1] / 1.18, 4)
        result[name] = (t1, t1, t1, t2)
    return result


def parse_quant(data: bytes, pwd: str = '') -> dict:
    """Quant MF – IN-GST. Three AUM tiers: Base+, Base, Open. Use Base (middle) tier."""
    txt = read_pdf(data, pwd)
    result = {}
    for line in txt.split('\n'):
        line = line.strip()
        ll = line.lower()
        if not (lll := ll).startswith('quant'): continue
        if any(w in lll for w in ('brokerage', 'category', 'taxation', 'exit load',
                                   'base plus', 'base  plus', 'open', 'powered by',
                                   'note', 'general', 'statutory')): continue
        nums = get_floats(line)
        if len(nums) < 3: continue
        # Name = part before first double-space (two-space separator used in Quant PDFs)
        parts = re.split(r'\s{2,}', line)
        name = parts[0].strip() if parts else ''
        if not name or len(name) < 5: continue
        # Last 3 values are Base+, Base, Open → use Base (2nd from last)
        v = round(nums[-2] / 1.18, 4)
        result[name] = (v, v, v, v)
    return result


_UTI_CLASS = [
    'Flexi Cap Fund','Large Cap Fund','Mid Cap Fund','Small Cap Fund','Value Fund',
    'Dividend Yield Fund','Sectoral/ Thematic','Sectoral/Thematic','Focused Fund','ELSS',
    'Liquid Fund','Credit Risk Fund','Long Duration Debt Fund','Gilt Fund',
    'Index Funds','INDEX FUND','Gold ETF','Other ETF','Thematic Fund',
    'Large &Mid Cap Fund','Large & Mid Cap Fund','Short Duration Debt Fund',
    'Conservative Hybrid Fund','Money Market Fund','Dynamic Asset Allocation Fund',
    'Aggressive Hybrid Fund','Equity Savings Fund','Multi Cap Fund',
    'Banking & PSU Debt Fund','Hybrid Fund','Debt Fund',
]

def parse_uti(data: bytes, pwd: str = '') -> dict:
    """UTI – IN-GST. Format: 'UTI X Fund [Classification] [ExitLoad] T1 T2 [B30]'"""
    txt = read_pdf(data, pwd)
    result = {}
    lines = txt.split('\n')
    merged, i = [], 0
    while i < len(lines):
        line = lines[i].strip()
        if line.startswith('UTI ') and not get_floats(line) and i + 1 < len(lines):
            merged.append(line + ' ' + lines[i+1].strip()); i += 2
        else:
            merged.append(line); i += 1
    for line in merged:
        if not line.startswith('UTI '): continue
        nums = get_floats(line)
        if len(nums) < 2: continue
        name = re.sub(r'[\d.%\s-]+$', '', line).strip()
        # Strip exit-load period "<1 Year - 1%" etc.
        name = re.sub(r'\s*[<>]?\d.*$', '', name).strip()
        name = re.sub(r'\s*NIL\s*$', '', name, flags=re.I).strip()
        # Strip "(Formerly UTI ...)" notes
        name = re.sub(r'\s*\(Formerly.*?\)', '', name).strip()
        # Strip classification suffix
        for cls in sorted(_UTI_CLASS, key=len, reverse=True):
            if name.endswith(cls):
                name = name[:-len(cls)].strip()
                break
        if not name or not name.startswith('UTI'): continue
        t1 = round(nums[0] / 1.18, 4)
        t2 = round(nums[1] / 1.18, 4)
        result[name] = (t1, t1, t1, t2)
    return result


def parse_whiteoak(data: bytes, pwd: str = '') -> dict:
    """WhiteOak Capital – IN-GST. Format: '[Category] WHITEOAK X FUND T1 T2 T3 T4 [ExitLoad]'"""
    txt = read_pdf(data, pwd)
    result = {}
    for line in txt.split('\n'):
        line = line.strip()
        if 'WHITEOAK' not in line.upper() and 'WhiteOak' not in line: continue
        idx = line.upper().find('WHITEOAK')
        if idx < 0: idx = line.find('WhiteOak')
        part = line[idx:]
        # Strip exit load description before extracting numbers
        part_clean = re.sub(r'\s*(?:Lock\s+in|\d+%\s+before|NIL|Refer).*$', '', part, flags=re.I).strip()
        nums = get_floats(part_clean) or get_floats(part)
        if not nums: continue
        name = re.sub(r'[\d.%\s]+$', '', part_clean).strip()
        name = re.sub(r'\s*\([A-Z]{2,6}\)\s*', ' ', name).strip()  # Strip "(YFCF)" ticker codes
        if not name or len(name) < 5: continue
        if len(nums) >= 4:
            v1 = round(nums[0] / 1.18, 4)
            v4 = round(nums[3] / 1.18, 4)
            result[name] = (v1, v1, v1, v4)
        else:
            v = round(nums[0] / 1.18, 4)
            result[name] = (v, v, v, v)
    return result


def parse_360one(data: bytes, pwd: str = '') -> dict:
    """360 ONE – IN-GST.
    Format A (older): '360 ONE Fund Name' on its own line, rate on next line.
    Format B (newer): 'Category  360 ONE Fund Name  rate%' all on one line.
    Additional Trail section at bottom is skipped (it uses smaller booster rates).
    """
    txt = read_pdf(data, pwd)
    result = {}
    in_additional = False
    lines = txt.split('\n')
    i = 0
    while i < len(lines):
        line = lines[i].strip()
        ll = line.lower()
        if 'additional trail' in ll or 'additional brokerage' in ll:
            in_additional = True
        if in_additional:
            i += 1
            continue
        has_360 = '360 ONE' in line or '360 One' in line
        if has_360:
            idx = line.find('360 ONE')
            if idx < 0: idx = line.find('360 One')
            part = line[idx:]
            nums_inline = get_floats(part)
            if nums_inline:
                # Format B: rate on same line
                name = re.sub(r'[\d.%\s]+$', '', part).strip()
                if name and len(name) > 5:
                    v = round(nums_inline[0] / 1.18, 4)
                    result[name] = (v, v, v, v)
            else:
                # Format A: rate on next line
                name = part.strip()
                for j in range(i + 1, min(i + 4, len(lines))):
                    nxt = lines[j].strip()
                    if '360' in nxt: break
                    nums = get_floats(nxt)
                    if nums:
                        v = round(nums[0] / 1.18, 4)
                        result[name] = (v, v, v, v)
                        break
        i += 1
    return result


def parse_kotak(data: bytes, pwd: str = '') -> dict:
    """Kotak – EX-GST. Complex table: 'Lump sum 1 to MAX [thld] T1 T2 T3 T4 T5' before scheme name."""
    txt = read_pdf(data, pwd)
    # Normalise "Lump sum1" / "Systematic1" → add space
    txt = txt.replace('Lump sum1', 'Lump sum 1').replace('Systematic1', 'Systematic 1')
    # Merge split scheme names: "...Index\nFund12-Apr-..." / "...50\nIndex Fund\n01..."
    txt = re.sub(r'([A-Za-z0-9])\n((?:[Ii]ndex\s+)?(?:FUND|Fund))', r'\1 \2', txt)
    txt = re.sub(r'([A-Za-z0-9])\n(INDEX)', r'\1 \2', txt)

    result = {}
    lines = txt.split('\n')
    i = 0
    while i < len(lines):
        line = lines[i].strip()
        m = re.match(
            r'Lump\s+sum\s+\d+\s+to\s+MAX\s+[\d.]+\s+([\d.]+)\s+([\d.]+)\s+([\d.]+)\s+([\d.]+)\s+([\d.]+)',
            line, re.I)
        if m:
            t = [float(m.group(x)) for x in range(1, 6)]
            # Scheme name appears in the next 1-3 lines
            for j in range(i + 1, min(i + 5, len(lines))):
                nxt = lines[j].strip()
                if 'KOTAK' not in nxt.upper(): continue
                name = nxt
                # Strip trailing date "01-Apr-2024 to 30-Jun-2024 ..."
                name = re.sub(r'\s*\d{2}-[A-Za-z]+-\d{4}.*$', '', name).strip()
                # Strip type labels directly appended without space: "FundINDEX" / "FUNDINDEX" → "Fund"
                name = re.sub(r'(?<=[A-Za-z])(?:INDEX|EQUITY|DEBT|HYBRID|FIXED).*$', '', name).strip()
                # Strip space-separated double type labels: "Fund INDEX INDEX" → "Fund"
                name = re.sub(r'\s+(?:INDEX|EQUITY|DEBT|HYBRID)\s+(?:INDEX|EQUITY|DEBT|HYBRID).*$',
                               '', name).strip()
                name = re.sub(r'\s*FIXED\s+NO.*$', '', name, flags=re.I).strip()
                # Strip trailing single type label that appears after "Fund": "...Fund INDEX" → "...Fund"
                name = re.sub(r'(?i)(fund)\s+(?:INDEX|EQUITY|DEBT|HYBRID|LIQUID)\s*$', r'\1', name).strip()
                if name and len(name) > 5:
                    result[name] = (t[0], t[1], t[2], t[3])
                break
        i += 1
    return result


def parse_groww(data: bytes, pwd: str = '') -> dict:
    """Groww – IN-GST. Format: 'Groww X Fund\\nT1%\\nT2%\\nT3%'"""
    txt = read_pdf(data, pwd)
    result = {}
    lines = txt.split('\n')
    i = 0
    while i < len(lines):
        line = lines[i].strip()
        if line.startswith('Groww ') or line.startswith(' Groww '):
            name = line.strip()
            nums = get_floats(line)
            if not nums and i + 1 < len(lines):
                nums = get_floats(lines[i + 1])
            if nums:
                v = round(nums[0] / 1.18, 4)
                result[name] = (v, v, v, v)
        i += 1
    return result


def parse_old_bridge(data: bytes, pwd: str = '') -> dict:
    """Old Bridge – IN-GST. Format: 'Old Bridge X Fund  T%  p.a.*'"""
    txt = read_pdf(data, pwd)
    result = {}
    for line in txt.split('\n'):
        line = line.strip()
        if not line.startswith('Old Bridge'): continue
        nums = get_floats(line)
        if not nums: continue
        name = re.sub(r'\s*[\d.%]+.*$', '', line).strip()
        name = name.rstrip('*').strip()
        if not name or len(name) < 5: continue
        v = round(nums[0] / 1.18, 4)
        result[name] = (v, v, v, v)
    return result


# ── Parser registry ───────────────────────────────────────────────────────────

def get_parser(amc: str):
    """Return (parser_fn, needs_bytes) for the given AMC identifier."""
    return {
        'absl':      (parse_absl,      False),
        'axis':      (parse_axis,      True),
        'bandhan':   (parse_bandhan,   True),
        'boi':       (parse_boi,       True),
        'canara':    (parse_canara,    True),
        'dsp':       (parse_dsp,       True),
        'ft':        (parse_ft,        True),
        'hdfc':      (parse_hdfc,      True),
        'hsbc':      (parse_hsbc,      True),
        'invesco':   (parse_invesco,   True),
        'lic':       (parse_lic,       True),
        'mahindra':  (parse_mahindra,  True),
        'mirae':     (parse_mirae,     True),
        'motilal':   (parse_motilal,   True),
        'nippon':    (parse_nippon,    True),
        'pgim':      (parse_pgim,      True),
        'sbi':       (parse_sbi,       True),
        'sundaram':  (parse_sundaram,  True),
        'tata':      (parse_tata,      True),
        'trust':     (parse_trust,     True),
        'icici':     (parse_icici,     True),
        'old_bridge':(parse_old_bridge,True),
        'bajaj':     (parse_bajaj,     True),
        'baroda':    (parse_baroda,    True),
        'edelweiss': (parse_edelweiss, True),
        'iti':       (parse_iti,       True),
        'jm':        (parse_jm,        True),
        'quant':     (parse_quant,     True),
        'uti':       (parse_uti,       True),
        'whiteoak':  (parse_whiteoak,  True),
        '360one':    (parse_360one,    True),
        'kotak':     (parse_kotak,     True),
        'groww':     (parse_groww,     True),
    }.get(amc, (None, False))
