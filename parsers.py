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
    return openpyxl.load_workbook(BytesIO(data))

# ── Number helpers ─────────────────────────────────────────────────────────────

def get_floats(line):
    return [float(x) for x in re.findall(r'\d+\.\d+', line.replace('%', ''))]

def get_all_nums(line):
    return [float(x) for x in re.findall(r'\d+(?:\.\d+)?', line.replace('%', ''))]

# ── Name normalisation ─────────────────────────────────────────────────────────

def norm(s: str) -> str:
    s = str(s).lower().strip()
    s = s.replace('aditya birla sun life', 'absl')
    s = s.replace('icici prudential', 'icici pru')
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
    s = re.sub(r'\b([a-z])(\s[a-z])+\b', lambda m: m.group(0).replace(' ', ''), s)
    return re.sub(r'\s+', ' ', s).strip()

# ── Fuzzy scheme matching ──────────────────────────────────────────────────────

def best_match(name, data):
    if name in data: return data[name]
    nl = name.lower()
    for k, v in data.items():
        if k.lower() == nl: return v
    nn = norm(name)
    for k, v in data.items():
        if norm(k) == nn: return v
    for k, v in data.items():
        kn = norm(k)
        if nn in kn or kn in nn: return v
    def _strip(text, word):
        return re.sub(r'\s+', ' ', re.sub(rf'\b{word}\b\s*', '', text)).strip()
    for strip_word in ('india', 'mf', 'elss'):
        nn_s = _strip(nn, strip_word)
        for k, v in data.items():
            kn_s = _strip(norm(k), strip_word)
            if nn_s == kn_s or nn_s in kn_s or kn_s in nn_s: return v
            if strip_word == 'elss':
                nw, kw = set(nn_s.split()), set(kn_s.split())
                if len(nw) >= 3 and len(kw) >= 3:
                    if nw.issubset(kw) and len(kw)-len(nw) <= 2: return v
                    if kw.issubset(nw) and len(nw)-len(kw) <= 2: return v
    nw = set(nn.split())
    for k, v in data.items():
        kw = set(norm(k).split())
        if len(nw) >= 3 and len(kw) >= 3:
            if nw.issubset(kw) and len(kw)-len(nw) <= 2: return v
            if kw.issubset(nw) and len(nw)-len(kw) <= 2: return v
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
    'icici':    ['icici prudential mutual fund commission','commission structure'],
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
}

def detect_source_amc(filename: str, preview: str = '') -> str | None:
    text = (filename + ' ' + preview[:800]).lower()
    for amc, hints in _SOURCE_HINTS.items():
        if any(h in text for h in hints):
            return amc
    fn = filename.lower()
    for amc in _SOURCE_HINTS:
        if amc in fn: return amc
    if 'franklin' in fn: return 'ft'
    if 'bank of india' in fn: return 'boi'
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
    for line in txt.split('\n'):
        line = line.strip()
        if not line: continue
        nums = get_floats(line)
        if len(nums) >= 4:
            name_part = re.sub(r'[\d.%\s]+$', '', line).strip()
            if name_part and len(name_part) > 5:
                v = nums[0]
                result[name_part] = (v, v, v, v)
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
    txt = read_pdf(data, pwd)
    result = {}
    lines = txt.split('\n')
    merged, i = [], 0
    while i < len(lines):
        line = lines[i].strip()
        if line.startswith('Bank of India') and not get_floats(line) and i + 1 < len(lines):
            merged.append(line + ' ' + lines[i+1].strip()); i += 2
        else:
            merged.append(line); i += 1
    for line in merged:
        line = re.sub(r'Click\s*[Hh]ere', '', line).strip()
        if not line.startswith('Bank of India'): continue
        nums = get_floats(line)
        if len(nums) >= 2:
            name = re.sub(r'[\d.%\s]+$', '', line).strip()
            result[name] = (nums[0], nums[0], nums[0], nums[1])
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
            base = nums[1] if len(nums) >= 3 else nums[0]
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
    txt = read_pdf(data, pwd)
    result = {}
    for line in txt.split('\n'):
        line = line.strip()
        m = re.match(r'\d+\)\s+(.+?)\s+([\d.]+)\s+([\d.]+)\s+([\d.]+)\s+([\d.]+)\s*$', line)
        if m:
            name = m.group(1).strip()
            name = re.sub(r'\s*\([A-Z]+\)[#*\s]*$', '', name).strip().rstrip('#*').strip()
            t1, t2 = float(m.group(2)), float(m.group(5))
            result[name] = (t1, t2)
            nu = name.upper()
            if 'INDEX FUND' in nu and 'NSE NIFTY' in nu and 'PLAN' in nu:
                result['FRANKLIN INDIA NSE NIFTY 50 INDEX FUND'] = (t1, t2)
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
    result = {}
    cat_re = re.compile(
        r'\b(FLEXI CAP FUND|LARGE CAP FUND|LARGE & MIDCAP FUND|CHILDREN\'S FUND|'
        r'MULTI CAP FUND|MIDCAP FUND|SMALL CAP FUND|EQUITY DIVIDEND YIELD|'
        r'EQUITY FOCUSED FUND|EQUITY VALUE FUND|SECTORAL/THEMATIC FUND|'
        r'EQUITY SAVINGS FUND|ARBITRAGE FUND|DYNAMIC ASSET ALLOCATION|ELSS|'
        r'INDEX FUND|GOLD FUND|MEDIUM TO LONG DURATION FUND|MONEY MARKET FUND|'
        r'BANKING & PSU DEBT FUND|GILT FUND|LOW DURATION FUND|SHORT DURATION FUND|'
        r'OVERNIGHT FUND|ULTRA SHORT DURATION FUND|LIQUID FUND|'
        r'CONSERVATIVE HYBRID FUND|AGGRESSIVE HYBRID FUND)\b', re.I)
    for line in txt.split('\n'):
        line = line.strip()
        if not line.startswith('LIC'): continue
        nums = get_floats(line)
        if len(nums) >= 12:
            name = re.sub(r'[\d.\s]+$', '', line).strip()
            name = cat_re.sub('', name).strip()
            if not name: continue
            result[name] = (nums[1], nums[4], nums[7], nums[10])
    return result


def parse_mahindra(data: bytes, pwd: str = '') -> dict:
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
        if len(nums) >= 3:
            name = re.sub(r'[\d.\s]+$', '', line).strip()
            for w in cat_words: name = name.replace(w, '').strip()
            if not name: continue
            result[name] = (nums[0],)
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
    txt = read_pdf(data, pwd)
    result = {}
    for line in txt.split('\n'):
        line = line.strip()
        if not line.startswith('Motilal'): continue
        nums = get_all_nums(line)
        if len(nums) >= 6:
            name = re.sub(r'[\d.\s]+$', '', line).strip()
            result[name] = (round(nums[0]/100,4),)*3 + (round(nums[3]/100,4),)
        elif len(nums) == 2:
            name = re.sub(r'[\d.\s]+$', '', line).strip()
            result[name] = (round(nums[0]/100,4),)*3 + (round(nums[1]/100,4),)
    return result


def parse_nippon(data: bytes, pwd: str = '') -> dict:
    txt = read_pdf(data, pwd)
    result = {}
    for line in txt.split('\n'):
        line = line.strip()
        if 'NIPPON INDIA' not in line.upper(): continue
        nums = get_floats(line)
        if len(nums) >= 8:
            m = re.search(r'NIPPON\s+INDIA', line, re.I)
            if m:
                part = line[m.start():]
                name = re.sub(r'\s+(?:NIL|\d+\s*(?:Month|Day|Year|Yr|yr|lock))[^\d]*.*$', '', part, flags=re.I).strip()
                name = re.sub(r'[\d.%\s]+$', '', name).strip()
            else:
                name = re.sub(r'[\d.%\s]+$', '', line).strip()
            if not name: continue
            result[name] = (nums[0], nums[1], nums[1], nums[2])
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
        if len(nums) >= 4:
            name = re.sub(r'[\d.%\s]+$', '', line).strip()
            name = re.sub(r'\d+\.?\d*%.*', '', name).strip()
            if not name: continue
            result[name] = (nums[0], nums[0], nums[0], nums[3])
        elif len(nums) == 3:
            name = re.sub(r'[\d.%\s]+$', '', line).strip()
            name = re.sub(r'\d+\.?\d*%.*', '', name).strip()
            if not name: continue
            result[name] = (nums[0], nums[0], nums[0], nums[0])
    return result


def parse_sbi(data: bytes, pwd: str = '') -> dict:
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
    txt = read_pdf(data, pwd)
    result = {}
    for line in txt.split('\n'):
        line = line.strip()
        if not line.upper().startswith('TRUST'): continue
        nums = get_floats(line)
        if len(nums) >= 1:
            name = re.sub(r'[\d.%\s]+$', '', line).strip()
            if not name: continue
            result[name] = (nums[0],)
    return result


def parse_icici(data: bytes) -> dict:
    """ICICI source is an Excel file; values are decimal fractions (×100 → %)."""
    wb = read_excel_wb(data)
    ws = wb.active
    result = {}
    for row in range(4, ws.max_row + 1):
        name = ws.cell(row, 2).value
        t1   = ws.cell(row, 3).value
        t2   = ws.cell(row, 4).value
        t3   = ws.cell(row, 5).value
        t4   = ws.cell(row, 6).value
        if not name or not isinstance(t1, (int, float)): continue
        vals = tuple(round(float(x) * 100, 4) for x in (t1, t2, t3, t4))
        key  = str(name).strip()
        result[key] = vals
        ku = key.upper()
        if 'REGULAR GOLD SAVINGS' in ku:
            result['ICICI Prudential Gold ETF FOF'] = vals
            result['ICICI Prudential Gold ETF'] = vals
    return result


# ── Parser registry ───────────────────────────────────────────────────────────

def get_parser(amc: str):
    """Return (parser_fn, needs_bytes) for the given AMC identifier."""
    return {
        'absl':     (parse_absl,     False),
        'axis':     (parse_axis,     True),
        'bandhan':  (parse_bandhan,  True),
        'boi':      (parse_boi,      True),
        'canara':   (parse_canara,   True),
        'dsp':      (parse_dsp,      True),
        'ft':       (parse_ft,       True),
        'hdfc':     (parse_hdfc,     True),
        'hsbc':     (parse_hsbc,     True),
        'invesco':  (parse_invesco,  True),
        'lic':      (parse_lic,      True),
        'mahindra': (parse_mahindra, True),
        'mirae':    (parse_mirae,    True),
        'motilal':  (parse_motilal,  True),
        'nippon':   (parse_nippon,   True),
        'pgim':     (parse_pgim,     True),
        'sbi':      (parse_sbi,      True),
        'sundaram': (parse_sundaram, True),
        'tata':     (parse_tata,     True),
        'trust':    (parse_trust,    True),
        'icici':    (parse_icici,    True),
    }.get(amc, (None, False))
