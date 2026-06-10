"""
fill_brokerage.py - Fill brokerage Excel files from corresponding PDFs.
Handles all 25 AMC Excel files for April-June 2026.
"""
import PyPDF2, openpyxl, re, os, sys

sys.stdout.reconfigure(encoding='utf-8')

BASE_PDF = r'c:\Users\tejas.rai\Desktop\mutual fund\Brokerage Structure & Formats for the month of April to June 2026'
BASE_XLS = os.path.join(BASE_PDF, 'April to June 2026', 'April to June 2026')
PDF_PWD = 'AADCH3479E'

# ─── Helpers ──────────────────────────────────────────────────────────────────

def read_pdf(fname, pwd=''):
    path = os.path.join(BASE_PDF, fname)
    with open(path, 'rb') as f:
        r = PyPDF2.PdfReader(f)
        if r.is_encrypted:
            r.decrypt(pwd)
        return '\n'.join(p.extract_text() or '' for p in r.pages)

def norm(s):
    s = str(s).lower().strip()
    s = s.replace('aditya birla sun life', 'absl')
    s = re.sub(r'\s*&\s*', ' and ', s)   # normalize all "&" variants
    s = re.sub(r"'", ' ', s)              # apostrophe → space
    s = re.sub(r'[-]', ' ', s)
    s = re.sub(r'[^a-z0-9 ]', ' ', s)
    # Compound word normalization
    s = s.replace('flexicap', 'flexi cap').replace('midcap', 'mid cap').replace('multicap', 'multi cap')
    s = s.replace('smallcap', 'small cap').replace('largecap', 'large cap')
    s = s.replace('nifty100', 'nifty 100').replace('nifty50', 'nifty 50').replace('nifty500', 'nifty 500')
    s = s.replace('healthcare', 'health care')
    s = s.replace('fof', 'fund of fund').replace('foف', 'fund of fund')
    # Plural normalization for common fund name words
    s = s.replace('savings fund', 'saving fund')
    s = s.replace('business groups', 'business group')
    s = s.replace('capital markets', 'capital market')
    s = re.sub(r'\s+', ' ', s).strip()
    # Join isolated single letters (e.g. T.I.G.E.R → tiger)
    s = re.sub(r'\b([a-z])(\s[a-z])+\b', lambda m: m.group(0).replace(' ', ''), s)
    return re.sub(r'\s+', ' ', s).strip()

def best_match(name, data):
    # 1. exact
    if name in data: return data[name]
    # 2. case-insensitive
    nl = name.lower()
    for k, v in data.items():
        if k.lower() == nl: return v
    # 3. normalised
    nn = norm(name)
    for k, v in data.items():
        if norm(k) == nn: return v
    # 4. partial: shorter name fully contained in longer
    for k, v in data.items():
        kn = norm(k)
        if nn in kn or kn in nn:
            return v
    # 5. strip "india" from both sides and retry
    nn_ni = re.sub(r'\s+', ' ', re.sub(r'\bindia\b\s*', '', nn)).strip()
    for k, v in data.items():
        kn_ni = re.sub(r'\s+', ' ', re.sub(r'\bindia\b\s*', '', norm(k))).strip()
        if nn_ni == kn_ni or nn_ni in kn_ni or kn_ni in nn_ni:
            return v
    # 6. strip "mf" from both sides (handles "LIC MF" vs "LIC")
    nn_nm = re.sub(r'\s+', ' ', re.sub(r'\bmf\b\s*', '', nn)).strip()
    for k, v in data.items():
        kn_nm = re.sub(r'\s+', ' ', re.sub(r'\bmf\b\s*', '', norm(k))).strip()
        if nn_nm == kn_nm or nn_nm in kn_nm or kn_nm in nn_nm:
            return v
    # 7. strip "elss" from both sides (handles "ELSS Tax Saver Fund" naming variants)
    nn_ne = re.sub(r'\s+', ' ', re.sub(r'\belss\b\s*', '', nn)).strip()
    for k, v in data.items():
        kn_ne = re.sub(r'\s+', ' ', re.sub(r'\belss\b\s*', '', norm(k))).strip()
        if nn_ne == kn_ne or nn_ne in kn_ne or kn_ne in nn_ne:
            return v
        # Also try word-subset after elss strip
        nn_ne_words = set(nn_ne.split())
        kn_ne_words = set(kn_ne.split())
        if len(nn_ne_words) >= 3 and len(kn_ne_words) >= 3:
            if nn_ne_words.issubset(kn_ne_words) and len(kn_ne_words) - len(nn_ne_words) <= 2:
                return v
            if kn_ne_words.issubset(nn_ne_words) and len(nn_ne_words) - len(kn_ne_words) <= 2:
                return v
    # 8. word-set subset: all words of shorter present in longer (max 2 extra words)
    nn_words = set(nn.split())
    for k, v in data.items():
        kn_words = set(norm(k).split())
        if len(nn_words) >= 3 and len(kn_words) >= 3:
            if nn_words.issubset(kn_words) and len(kn_words) - len(nn_words) <= 2:
                return v
            if kn_words.issubset(nn_words) and len(nn_words) - len(kn_words) <= 2:
                return v
    return None

def pct(s):
    """Strip % sign and convert to float, e.g. '1.06%' → 1.06"""
    return float(str(s).replace('%','').strip())

def get_floats(line):
    """Extract only decimal-point numbers (e.g. 0.75, 1.25%)"""
    line = line.replace('%', '')
    return [float(x) for x in re.findall(r'\d+\.\d+', line)]

def get_all_nums(line):
    """Extract all numbers (integers and decimals) from a line"""
    line = line.replace('%', '')
    return [float(x) for x in re.findall(r'\d+(?:\.\d+)?', line)]

def fill_excel(xls_name, data):
    """
    data: {scheme_name: tuple_of_trail_values}
    Each tuple maps to (T1, T2, T3, T4+) or shorter if fewer columns exist.
    """
    path = os.path.join(BASE_XLS, xls_name)
    wb = openpyxl.load_workbook(path)
    ws = wb.active

    # Column indices
    hdr = {}
    for c in range(1, ws.max_column + 1):
        v = ws.cell(1, c).value
        if v: hdr[v] = c

    scheme_col = hdr.get('SchemeName', 4)
    trail_order = [
        'Trail1stYear', 'T1',
        'Trail2ndYear', 'T2',
        'Trail3rdYear', 'T3',
        'Trail3rdYearOnwards',
        'Trail4thYearOnwards', 'T4+',
        'Trail1stYearOnwards',
        'Trail2ndYearOnwards',
    ]
    trail_cols = [(n, hdr[n]) for n in trail_order if n in hdr]

    filled, not_found = [], []
    for row in range(2, ws.max_row + 1):
        scheme = ws.cell(row, scheme_col).value
        if not scheme: continue
        match = best_match(str(scheme), data)
        if match is not None:
            vals = match if isinstance(match, tuple) else (match,) * 4
            for i, (col_name, col_idx) in enumerate(trail_cols):
                if i < len(vals) and vals[i] is not None:
                    ws.cell(row, col_idx).value = round(float(vals[i]), 4)
            filled.append(str(scheme))
        else:
            not_found.append(str(scheme))

    wb.save(path)
    return filled, not_found

# ─── Parsers ──────────────────────────────────────────────────────────────────

def parse_absl():
    """ABSL: hardcoded from Read-tool output (T1=T2=T3 for all schemes)."""
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

def parse_axis(pdf_fname, pwd=''):
    """AXIS: each line 'Scheme v1 v1 v1 v1' → T1=T2=T3=T4+=v1"""
    txt = read_pdf(pdf_fname, pwd)
    data = {}
    for line in txt.split('\n'):
        line = line.strip()
        if not line: continue
        nums = get_floats(line)
        if len(nums) >= 4:
            # Remove trailing numbers to get scheme name
            name_part = re.sub(r'[\d.%\s]+$', '', line).strip()
            if name_part and len(name_part) > 5:
                v = nums[0]
                data[name_part] = (v, v, v, v)
    return data

def parse_bandhan_april():
    """BANDHAN April: 'Scheme ALL t1excl% gst% t1incl% t2excl%' → all same as t1excl"""
    txt = read_pdf("BANDHAN Brokerage Structure effective from 1st Apr 2026 until further notice.pdf")
    data = {}
    for line in txt.split('\n'):
        line = line.strip()
        m_all = re.search(r'\bALL\b', line)
        if not line or not m_all: continue
        nums = get_floats(line)
        if len(nums) >= 4:
            name = line[:m_all.start()].strip()
            if name and len(name) > 5:
                v = nums[0]  # T1 EX-GST
                data[name] = (v, v, v, v)
    return data

def parse_bandhan_may():
    """BANDHAN May: 'Scheme Category t1% t2% t3% t4%' → 4 trail values"""
    txt = read_pdf("BANDHAN Brokerage Structure effective from 1st May 2026 until further notice.pdf")
    data = {}
    category_words = {'Sectoral/Thematic Funds','Flexi Cap Fund','Focused Fund','ELSS',
                      'Large & Mid Cap Fund','Large Cap Fund','Mid Cap Fund','Multi Cap Fund',
                      'Small Cap Fund','Aggressive Hybrid Fund','Arbitrage Fund',
                      'Balanced Advantage Fund','Conservative Hybrid Fund','Equity Savings Fund',
                      'Multi Asset Allocation Fund','Low Duration Fund','Money Market Fund',
                      'Ultra Short Duration Fund','Banking and PSU Fund','Corporate Bond Fund',
                      'Credit Risk Fund','Dynamic Bond Fund','Floater Fund','Gilt Fund',
                      'Long Duration Fund','Medium Duration Fund','Medium to Long Duration Fund',
                      'Short Duration Fund','Liquid Fund','Overnight Fund','Index Fund',
                      'Fund of Fund','Domestic Fund','Retirement Fund'}
    for line in txt.split('\n'):
        line = line.strip()
        nums = get_floats(line)
        if len(nums) >= 4:
            # Remove category words to find scheme name
            name = line
            for cw in category_words:
                name = name.replace(cw, '').strip()
            name = re.sub(r'[\d.%\s]+$', '', name).strip()
            if name and len(name) > 5:
                v1, v2, v3, v4 = nums[0], nums[1], nums[2], nums[3]
                data[name] = (v1, v2, v3, v4)
    return data

def parse_boi(pdf_fname, pwd=''):
    """BOI: split lines need joining; format 'Scheme t1-3% t4+% Click here'"""
    txt = read_pdf(pdf_fname, pwd)
    data = {}
    lines = txt.split('\n')
    # Join split lines: 'Bank of India ...' with no trailing numbers + next line
    merged = []
    i = 0
    while i < len(lines):
        line = lines[i].strip()
        if line.startswith('Bank of India') and not get_floats(line):
            nxt = lines[i+1].strip() if i+1 < len(lines) else ''
            merged.append(line + ' ' + nxt)
            i += 2
        else:
            merged.append(line)
            i += 1
    for line in merged:
        line = re.sub(r'Click\s*[Hh]ere', '', line).strip()
        if not line.startswith('Bank of India'): continue
        nums = get_floats(line)
        if len(nums) >= 2:
            name = re.sub(r'[\d.%\s]+$', '', line).strip()
            data[name] = (nums[0], nums[0], nums[0], nums[1])
    return data

def parse_canara():
    """CANARA: some schemes split over 2 lines; format 'Scheme exit_load total base gst'"""
    txt = read_pdf("CANARA ROBECO Brokerage Structure - April to June 26.pdf")
    data = {}
    lines = txt.split('\n')
    # Join split lines: 'Canara Robeco ...' with no numbers + next line
    merged = []
    i = 0
    while i < len(lines):
        line = lines[i].strip()
        if line.startswith('Canara Robeco') and not get_floats(line):
            nxt = lines[i+1].strip() if i+1 < len(lines) else ''
            merged.append(line + ' ' + nxt)
            i += 2
        else:
            merged.append(line)
            i += 1
    for line in merged:
        if not line.startswith('Canara Robeco'): continue
        nums = get_floats(line)
        if len(nums) >= 2:
            name = re.sub(r'[\d.%\s]+$', '', line)
            name = re.sub(r'\d.*', '', name).strip()
            if not name: continue
            base = nums[1] if len(nums) >= 3 else nums[0]
            data[name] = (base,)
    return data

def parse_dsp():
    """DSP: 'Scheme base1% gst1% total1% base4% gst4% total4%'
    → T1=T2=T3=base1, T4+=base4"""
    txt = read_pdf("DSP Commission Structure - 1st April 2026 - 30th June 2026.pdf")
    data = {}
    for line in txt.split('\n'):
        line = line.strip()
        if not line.startswith('DSP'): continue
        nums = get_floats(line)
        if len(nums) >= 6:
            name = re.sub(r'[\d.%\s]+$', '', line).strip()
            t13 = nums[0]  # Year 1-3 base EX-GST
            t4 = nums[3]   # Year 4+ base EX-GST
            data[name] = (t13, t13, t13, t4)
        elif len(nums) >= 3:
            name = re.sub(r'[\d.%\s]+$', '', line).strip()
            data[name] = (nums[0], nums[0], nums[0], nums[0])
    return data

def parse_ft():
    """FT: 'N) FUND NAME (CODE)# v1 v2 v3 v4'
    → T1=v1 (base EX-GST), T2+=v4 (same as v1)"""
    txt = read_pdf("FT Brokerage Effective 1st April 2026 to 30th June 2026.pdf")
    data = {}
    for line in txt.split('\n'):
        line = line.strip()
        m = re.match(r'\d+\)\s+(.+?)\s+([\d.]+)\s+([\d.]+)\s+([\d.]+)\s+([\d.]+)\s*$', line)
        if m:
            name = m.group(1).strip()
            # Remove "(CODE)#*" footnote patterns from end
            name = re.sub(r'\s*\([A-Z]+\)[#\*\s]*$', '', name).strip()
            name = name.rstrip('#*').strip()
            t1 = float(m.group(2))  # Col A: base trail
            t2 = float(m.group(5))  # Col D: 2nd yr onwards (same as T1 for FT)
            data[name] = (t1, t2)
            # Alias: "INDEX FUND NSE NIFTY PLAN" → "NSE NIFTY 50 INDEX FUND"
            nu = name.upper()
            if 'INDEX FUND' in nu and 'NSE NIFTY' in nu and 'PLAN' in nu:
                data['FRANKLIN INDIA NSE NIFTY 50 INDEX FUND'] = (t1, t2)
    return data

def parse_hdfc():
    """HDFC: 'Scheme Category exit_load t13% t4% 3yr%'
    → T1=T2=T3=t13, T4+=t4"""
    txt = read_pdf("HDFC Brokerage Structure - April to June 26.pdf")
    data = {}
    # Also parse NFO file
    try:
        txt += '\n' + read_pdf("HDFC Gold Silver Passive FOF Brokerage Structure - NFO.pdf")
    except Exception:
        pass
    for line in txt.split('\n'):
        line = line.strip()
        if not line.startswith('HDFC'): continue
        nums = get_floats(line)
        if len(nums) >= 3:
            # Strip footnote markers and everything after
            name = re.sub(r'\s*#.*', '', line).strip()
            # Strip exit load text and trailing numbers
            name = re.sub(r'\s+(?:NIL|\d+\s*(?:Months?|Days?|Year|Yrs?|yr))\s*.*$', '', name, flags=re.I).strip()
            name = re.sub(r'[\d.%\s]+$', '', name).strip()
            if not name or len(name) < 4: continue
            t13 = nums[0]  # Year 1-3 APM
            t4 = nums[1]   # Year 4+ APM
            data[name] = (t13, t13, t13, t4)
    return data

def parse_hsbc(pdf_fname):
    """HSBC: some lines split mid-name; join them before parsing.
    Format: 'HSBC Scheme Category t13% t4%' → T1=T2=T3=t13, T4+=t4"""
    txt = read_pdf(pdf_fname)
    data = {}
    # Join split lines: HSBC line with no numbers + next line with numbers
    lines = txt.split('\n')
    merged = []
    i = 0
    while i < len(lines):
        line = lines[i].strip()
        if line.startswith('HSBC') and not get_floats(line):
            # This line has the scheme name but no values; grab next line
            nxt = lines[i+1].strip() if i+1 < len(lines) else ''
            merged.append(line + ' ' + nxt)
            i += 2
        else:
            merged.append(line)
            i += 1
    for line in merged:
        line = line.strip()
        if not line.startswith('HSBC'): continue
        nums = get_floats(line)
        if len(nums) >= 2:
            name = re.sub(r'[\d.%\s]+$', '', line).strip()
            if not name or len(name) < 4: continue
            t13 = nums[0]
            t4 = nums[1]
            data[name] = (t13, t13, t13, t4)
    return data

def parse_invesco():
    """Invesco: 'Scheme t13_base gst total t4_base' OR 'CATEGORY Scheme ...'
    → T1=T2=T3=t13_base, T4+=t4_base"""
    txt = read_pdf("Invesco Brokerage Structure - April to June 26.pdf")
    data = {}
    for line in txt.split('\n'):
        line = line.strip()
        if 'Invesco' not in line: continue
        nums = get_floats(line)
        if len(nums) >= 4:
            # Find "Invesco" start; strip category prefix
            idx = line.find('Invesco')
            part = line[idx:]
            name = re.sub(r'[\d.\s]+$', '', part).strip()
            t13 = nums[0]
            t4 = nums[3]
            data[name] = (t13, t13, t13, t4)
    return data

def parse_lic():
    """LIC: 'SCHEME CATEGORY incl1 excl1 gst1 incl2 excl2 gst2 incl3 excl3 gst3 incl4 excl4 gst4'
    → T1=excl1, T2=excl2, T3=excl3, T4+=excl4"""
    txt = read_pdf("LIC_MF_Commission_Structure_Apr-26_to_Jun-26_-_SPECIAL.pdf")
    data = {}
    for line in txt.split('\n'):
        line = line.strip()
        if not line.startswith('LIC'): continue
        nums = get_floats(line)
        if len(nums) >= 12:
            name = re.sub(r'[\d.\s]+$', '', line).strip()
            # Remove category words
            name = re.sub(r'\b(FLEXI CAP FUND|LARGE CAP FUND|LARGE & MIDCAP FUND|'
                          r"CHILDREN'S FUND|MULTI CAP FUND|MIDCAP FUND|SMALL CAP FUND|"
                          r'EQUITY DIVIDEND YIELD|EQUITY FOCUSED FUND|EQUITY VALUE FUND|'
                          r'SECTORAL/THEMATIC FUND|EQUITY SAVINGS FUND|ARBITRAGE FUND|'
                          r'DYNAMIC ASSET ALLOCATION|ELSS|INDEX FUND|GOLD FUND|'
                          r'MEDIUM TO LONG DURATION FUND|MONEY MARKET FUND|'
                          r'BANKING & PSU DEBT FUND|GILT FUND|LOW DURATION FUND|'
                          r'SHORT DURATION FUND|OVERNIGHT FUND|ULTRA SHORT DURATION FUND|'
                          r'LIQUID FUND|CONSERVATIVE HYBRID FUND|AGGRESSIVE HYBRID FUND|'
                          r'MIDCAP FUND)\b', '', name, flags=re.IGNORECASE).strip()
            name = name.strip()
            if not name: continue
            t1 = nums[1]   # EX-GST Y1
            t2 = nums[4]   # EX-GST Y2
            t3 = nums[7]   # EX-GST Y3
            t4 = nums[10]  # EX-GST Y4+
            data[name] = (t1, t2, t3, t4)
    return data

def parse_mahindra():
    """MAHINDRA: 'Scheme Category base gst total' → single trail = base"""
    txt = read_pdf("MAHINDRA  Commission Structure - 1st April 2026 - FURTHER NOTICE.pdf")
    data = {}
    for line in txt.split('\n'):
        line = line.strip()
        if not line.startswith('Mahindra'): continue
        nums = get_floats(line)
        if len(nums) >= 3:
            name = re.sub(r'[\d.\s]+$', '', line).strip()
            for suffix in ['ELSS (Tax Saver)', 'Large-Cap', 'Mid-Cap', 'Small Cap',
                           'Large & Mid Cap', 'Multi-Cap', 'Flexi Cap', 'Focused',
                           'Thematic', 'Value Fund', 'Sectoral Fund', 'Equity Savings',
                           'Balanced Advantage', 'Aggressive Hybrid', 'Hybrid',
                           'Fund of Funds', 'FOF Domestic', 'Dynamic Bond', 'Liquid',
                           'Debt', 'Low Duration', 'Short Duration']:
                name = name.replace(suffix, '').strip()
            name = name.strip()
            if not name: continue
            base = nums[0]  # EX-GST
            data[name] = (base,)
    return data

def parse_mirae():
    """MIRAE: multi-line entries - extract scheme name and 4 trail values"""
    txt = read_pdf("MIRAE Brokerage Structure - April to June 26.pdf")
    data = {}
    # Replace "00." prefix artifacts
    txt = re.sub(r'\b00\.', '0.', txt)
    lines = txt.split('\n')
    i = 0
    while i < len(lines):
        line = lines[i].strip()
        # Asset class markers
        if line in ('Equity', 'Debt', 'Hybrid', 'Gold'):
            i += 1
            continue
        if line.lower().startswith('mirae asset'):
            # Collect this line and next line(s) for values
            combined = line
            j = i + 1
            while j < len(lines) and j < i + 25:
                combined += ' ' + lines[j].strip()
                j += 1
            nums = get_floats(combined)
            if len(nums) >= 5:
                name = re.sub(r'[\d.%±\s]+$', '', combined).strip()
                name = re.sub(r'\d.*', '', name).strip()
                # SIP trail is first, then T1, T2, T3, T4+
                t1, t2, t3, t4 = nums[1], nums[2], nums[3], nums[4]
                data[name] = (t1, t2, t3, t4)
            elif len(nums) >= 1:
                name = re.sub(r'[\d.%±\s]+$', '', line).strip()
                v = nums[0] if nums else None
                if v and name:
                    data[name] = (v, v, v, v)
        i += 1
    return data

def parse_motilal():
    """MOTILAL: BPS values (integers); 'Scheme bps_t13 gst total bps_t4 gst total ...'
    → T1=T2=T3=bps_t13/100, T4+=bps_t4/100"""
    txt = read_pdf("MOTILAL OSWAL Brokerage Structure - April to June 26.pdf")
    data = {}
    for line in txt.split('\n'):
        line = line.strip()
        if not line.startswith('Motilal'):
            continue
        nums = get_all_nums(line)
        if len(nums) >= 6:
            name = re.sub(r'[\d.\s]+$', '', line).strip()
            t13 = round(nums[0] / 100, 4)   # bps year 1-3
            t4  = round(nums[3] / 100, 4)   # bps year 4+
            data[name] = (t13, t13, t13, t4)
        elif len(nums) == 2:
            # Arbitrage: 'Fund  90 50' (two period-based rates)
            name = re.sub(r'[\d.\s]+$', '', line).strip()
            t13 = round(nums[0] / 100, 4)
            t4  = round(nums[1] / 100, 4)
            data[name] = (t13, t13, t13, t4)
    return data

def parse_nippon():
    """NIPPON: 'Category NIPPON INDIA SCHEME exit_load T1excl T2-3excl T4-5excl T6+ T1incl...'
    → T1=col1, T2=T3=col2, T4+=col3 (all EX-GST)"""
    txt = read_pdf("NIPPON Brokerage Structure - April to June 26.pdf")
    data = {}
    for line in txt.split('\n'):
        line = line.strip()
        if 'NIPPON INDIA' not in line.upper():
            continue
        nums = get_floats(line)
        if len(nums) >= 8:
            # Start from 'NIPPON INDIA' to strip category prefix
            m2 = re.search(r'NIPPON\s+INDIA', line, re.IGNORECASE)
            if m2:
                part = line[m2.start():]
                # Strip exit load text and trailing numbers
                name = re.sub(r'\s+(?:NIL|\d+\s*(?:Month|Day|Year|Yr|yr|lock))[^\d]*.*$', '', part, flags=re.I).strip()
                name = re.sub(r'[\d.%\s]+$', '', name).strip()
            else:
                name = re.sub(r'[\d.%\s]+$', '', line).strip()
            if not name: continue
            t1 = nums[0]
            t2 = nums[1]
            t4 = nums[2]
            data[name] = (t1, t2, t2, t4)
            # Add aliases for funds that were renamed
            nu = name.upper()
            if 'POWER' in nu and 'INFRA' in nu:
                alias = re.sub(r'POWER\s*(?:AND|&)\s*INFRA', 'INFRASTRUCTURE', nu, flags=re.I)
                data[alias] = (t1, t2, t2, t4)
            if nu == 'NIPPON INDIA FOCUSED FUND':
                data['NIPPON INDIA FOCUSED LARGE CAP FUND'] = (t1, t2, t2, t4)
            if 'MULTI ASSET FUND' in nu and 'OMNI' not in nu:
                data[nu.replace('MULTI ASSET FUND', 'MULTI ASSET ALLOCATION FUND')] = (t1, t2, t2, t4)
    return data

def parse_pgim():
    """PGIM: 'Scheme exit_load total_1yr base add_t3 t4+'
    → T1=T2=T3=total_1yr, T4+=t4+"""
    txt = read_pdf("PGIM Brokerage Structure - April to June 26.pdf")
    data = {}
    for line in txt.split('\n'):
        line = line.strip()
        if not line.startswith('PGIM'): continue
        nums = get_floats(line)
        if len(nums) >= 4:
            name = re.sub(r'[\d.%\s]+$', '', line).strip()
            # Remove exit load info
            name = re.sub(r'\d+\.?\d*%.*', '', name).strip()
            if not name: continue
            t13 = nums[0]  # Total Trail Year 1 (= T1=T2=T3)
            t4 = nums[3]   # Trail Year 4+
            data[name] = (t13, t13, t13, t4)
        elif len(nums) == 3:
            name = re.sub(r'[\d.%\s]+$', '', line).strip()
            name = re.sub(r'\d+\.?\d*%.*', '', name).strip()
            if not name: continue
            data[name] = (nums[0], nums[0], nums[0], nums[0])
    return data

def parse_sbi():
    """SBI: each line contains 'SBI <name> T1excl T2+excl T1incl T2+incl'
    → T1=T1excl, T2+=T2+excl"""
    txt = read_pdf("SBI MF Brokerage Structure for Q1 - APR TO JUNE'26.pdf")
    data = {}
    for line in txt.split('\n'):
        line = line.strip()
        if 'SBI' not in line: continue
        # Find the start of "SBI " in the line
        idx = line.find('SBI ')
        if idx < 0: continue
        part = line[idx:]
        nums = get_floats(part)
        if len(nums) >= 2:
            name = re.sub(r'[\d.%\s]+$', '', part).strip()
            # Strip trailing exit load text like "1 year" "365 days"
            name = re.sub(r'\s+\d+\.?\d*\s*(?:%|years?|yrs?|months?|days?).*$', '', name, flags=re.I).strip()
            if not name or not name.startswith('SBI'): continue
            data[name] = (nums[0], nums[1])
    return data

def parse_sundaram():
    """SUNDARAM: 'Scheme trail trail' → single trail"""
    txt = read_pdf("SUNDARAM Brokerage Structure - April to June 26.pdf", '')
    data = {}
    for line in txt.split('\n'):
        line = line.strip()
        if not line.startswith('Sundaram'): continue
        nums = get_floats(line)
        if len(nums) >= 1:
            name = re.sub(r'[\d.\s*]+$', '', line).strip()
            name = name.rstrip('*').strip()
            if not name: continue
            data[name] = (nums[0],)
            # Add full-name alias for abbreviated fund names
            expanded = name.replace('Fin. Services', 'Financial Services').replace('Opps', 'Opportunities')
            if expanded != name:
                data[expanded] = (nums[0],)
    return data

def parse_tata():
    """TATA: 'Scheme v%' → single trail; handle PP/MP/CP retirement variants"""
    txt = read_pdf("TATA Brokerage Structure - April" + chr(39) + "2026 to June" + chr(39) + "2026.pdf")
    data = {}
    for line in txt.split('\n'):
        line = line.strip()
        if not line.startswith('Tata'): continue
        nums = get_floats(line)
        if len(nums) >= 1:
            name = re.sub(r'[\d.%\s]+$', '', line).strip()
            if not name: continue
            # Expand retirement fund abbreviations and normalize "Savings" → "Saving"
            name = (name.replace('Tata Retirement Savings Fund -PP', 'Tata Retirement Saving Fund - Progressive')
                        .replace('Tata Retirement Savings Fund -MP', 'Tata Retirement Saving Fund - Moderate')
                        .replace('Tata Retirement Savings Fund -CP', 'Tata Retirement Saving Fund - Conservative'))
            data[name] = (nums[0],)
    return data

def parse_trust():
    """TRUST: 'Equity/Fixed/Hybrid TRUSTMF Scheme t1 gst total t2'
    → T1=col1, T2+=col4"""
    txt = read_pdf("TRUST Brokerage Structure - April to June 26.pdf")
    data = {}
    for line in txt.split('\n'):
        line = line.strip()
        if 'TRUSTMF' not in line: continue
        nums = get_floats(line)
        if len(nums) >= 4:
            name_m = re.search(r'(TRUSTMF\s+[\w\s]+?)(?:\s+[\d.])', line)
            if name_m:
                name = name_m.group(1).strip()
                data[name] = (nums[0], nums[3])
    return data

# ─── Main ──────────────────────────────────────────────────────────────────────

TASKS = [
    # (excel_file, data_getter_func)
    ("ABSL April to June.xlsx",        parse_absl),
    ("Axis -April 2026.xlsx",
        lambda: parse_axis("AXIS brokerage structure effective 1st April 2026 till 30th April 2026_.pdf", PDF_PWD)),
    ("Axis -May 2026.xlsx",
        lambda: parse_axis("AXIS brokerage structure effective MAY 2026_.pdf")),
    ("Bandhan-April 2026.xlsx",        parse_bandhan_april),
    ("Bandhan -May 2026.xlsx",         parse_bandhan_may),
    ("BOI-April 2026.xlsx",
        lambda: parse_boi("BOI Commission Structure - 1st April 2026 - 30th April 2026.pdf")),
    ("BOI-May to June 2026.xlsx",
        lambda: parse_boi("BOI Commission Structure - MAY TO JUNE 2026.pdf", PDF_PWD)),
    ("CANARA-April to June 2026.xlsx", parse_canara),
    ("DSP - April to June 2026.xlsx",  parse_dsp),
    ("FT - April to June 2026.xlsx",   parse_ft),
    ("HDFC - April to June 2026.xlsx", parse_hdfc),
    ("HSBC-April 2026.xlsx",
        lambda: parse_hsbc("HSBC Brokerage Structure - April'2026.pdf")),
    ("HSBC-May 2026.xlsx",
        lambda: parse_hsbc("HSBC Brokerage Structure - May'2026.pdf")),
    ("HSBC-June 2026.xlsx",
        lambda: parse_hsbc("HSBC Brokerage Structure - June'2026.pdf")),
    ("Invesco - April to June 2026.xlsx", parse_invesco),
    ("LIC- April to June 2026.xlsx",   parse_lic),
    ("Mahindra - April to June 2026.xlsx", parse_mahindra),
    ("Mirae -April to June 2026.xlsx", parse_mirae),
    ("Motilal - April to June 2026.xlsx", parse_motilal),
    ("Nippon-April to June 2026.xlsx", parse_nippon),
    ("PGIM - April to June 2026.xlsx", parse_pgim),
    ("SBI -April to June 2026.xlsx",   parse_sbi),
    ("SUNDARAM - April to June 2026.xlsx", parse_sundaram),
    ("TATA - April to June 2026.xlsx", parse_tata),
    ("Trust-April to June 2026.xlsx",  parse_trust),
]

overall_not_found = {}

for xls_name, getter in TASKS:
    print(f"\n{'='*60}")
    print(f"Processing: {xls_name}")
    try:
        data = getter()
        print(f"  PDF schemes found: {len(data)}")
        filled, not_found = fill_excel(xls_name, data)
        print(f"  Filled: {len(filled)} schemes")
        if not_found:
            print(f"  NOT MATCHED ({len(not_found)}):")
            for s in not_found:
                print(f"    - {s}")
            overall_not_found[xls_name] = not_found
    except Exception as e:
        import traceback
        print(f"  ERROR: {e}")
        traceback.print_exc()

print(f"\n{'='*60}")
print("DONE. Summary of unmatched schemes:")
for xls, schemes in overall_not_found.items():
    if schemes:
        print(f"\n  {xls}:")
        for s in schemes:
            print(f"    - {s}")
