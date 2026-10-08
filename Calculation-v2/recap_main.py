"""
update_recap.py  —  Fixed + Code_Mapping version
Fixes the merged-cell row corruption bug by:
  1. Capturing template styles as deep copies BEFORE any row manipulation
  2. Only inserting the EXTRA rows needed (not delete-all + insert-all)
Also resolves diamond / LGD / stone quality codes to descriptive names
using Code_Mapping.xlsx (sheets: Dia, LGD, Stones).
"""

import csv, copy, os, re, sys, zipfile
from openpyxl.utils import range_boundaries, get_column_letter
import openpyxl
from datetime import date

if sys.stdout and hasattr(sys.stdout, 'reconfigure'):
    try: sys.stdout.reconfigure(encoding='utf-8')
    except Exception: pass

# ─── Paths ────────────────────────────────────────────────────────────────────
MAPPING_FILE=r"C:\Users\ASUS\Documents\Manufacturing\3Marketing\Recap\files_BT\Code_Mapping.xlsx"
TOLERANCE_CSV = r"C:\Users\ASUS\Documents\Manufacturing\3Marketing\Recap\files_BT\tollerance.csv"   
RECAP_IN  =  r"C:\Users\ASUS\Documents\Manufacturing\3Marketing\Recap\files_BT\Belle's 35th Anniversary Spring test recap.xlsx"
CSV_FILE  =  r"C:\Users\ASUS\Documents\Manufacturing\3Marketing\Recap\outputs\Belles_35th_Anniversary\structured_styles.csv"
RECAP_OUT =  r"C:\Users\ASUS\Documents\Manufacturing\3Marketing\Recap\outputs\Belles_35th_Anniversary\Belles_35th_Anniversary_recap_UPDATED.xlsx"

# ─── Metal code → display ─────────────────────────────────────────────────────
METAL_CELL = {          # multi-line for the METAL cell
    'G14Y'  : '14K\nYELLOW\nGOLD',
    'G14W'  : '14K\nWHITE\nGOLD',
    'G14W/Y': '14K\nWHITE/YELLOW\nGOLD',
    'G14WY' : '14K\nWHITE/YELLOW\nGOLD',
    'G10Y'  : '10K\nYELLOW\nGOLD',
    'G10W'  : '10K\nWHITE\nGOLD',
    'G18Y'  : '18K\nYELLOW\nGOLD',
    'G18W'  : '18K\nWHITE\nGOLD',
    'SIL'   : 'STERLING\nSILVER',
    'PLT'   : 'PLATINUM',
}
METAL_DESC = {          # single-line for the Description field
    'G14Y'  : '14K Yellow Gold',
    'G14W'  : '14K White Gold',
    'G14W/Y': '14K White/Yellow Gold',
    'G14WY' : '14K White/Yellow Gold',
    'G10Y'  : '10K Yellow Gold',
    'G10W'  : '10K White Gold',
    'G18Y'  : '18K Yellow Gold',
    'G18W'  : '18K White Gold',
    'SIL'   : 'Sterling Silver',
    'PLT'   : 'Platinum',
}

def fmt_metal_cell(code): return METAL_CELL.get(code.strip(), code.strip())
def fmt_metal_desc(code): return METAL_DESC.get(code.strip(), code.strip())

# ─── Load Code_Mapping.xlsx (Dia / LGD / Stones) ─────────────────────────────
def load_code_mapping(filepath):
    """
    Returns three dicts: dia_map, lgd_map, stone_map
    Each dict: { code_upper: description_string }
    """
    wb_m = openpyxl.load_workbook(filepath, read_only=True, data_only=True)
    maps = {}
    for sheet in ['Dia', 'LGD', 'Stones']:
        d = {}
        ws_m = wb_m[sheet]
        first = True
        for row in ws_m.iter_rows(values_only=True):
            if first:           # skip header
                first = False
                continue
            if row[0] is not None:
                d[str(row[0]).strip().upper()] = str(row[1]).strip() if row[1] else str(row[0]).strip()
        maps[sheet] = d
        print(f"  Loaded {len(d):>5} codes from sheet '{sheet}'")
    wb_m.close()
    return maps['Dia'], maps['LGD'], maps['Stones']

print("Loading Code_Mapping...")
DIA_MAP, LGD_MAP, STONE_MAP = load_code_mapping(MAPPING_FILE)

# ─── Load Tolerance Ranges (tollerance.csv) ──────────────────────────────────
def load_tolerance_ranges(filepath):
    ranges = []
    if not os.path.exists(filepath):
        return ranges
    with open(filepath, encoding='utf-8') as f:
        for row in csv.reader(f):
            if not row or not row[0].strip():
                continue
            rng_str = row[0].strip()
            if 'no fraction' in rng_str.lower() or 'fraction range' in rng_str.lower():
                continue
            m = re.match(r'([\d\.]+)\s*-\s*([\d\.]+)', rng_str)
            if m and len(row) > 2 and row[2].strip():
                ranges.append((float(m.group(1)), float(m.group(2)), re.sub(r'\s+', ' ', row[2].strip())))
    return ranges

print("Loading Tolerance Ranges from tollerance.csv...")
TOLERANCE_RANGES = load_tolerance_ranges(TOLERANCE_CSV)

def get_fraction(wt, ranges):
    if not wt or wt <= 0: return ''
    for mn, mx, fr in ranges:
        if mn <= wt <= mx:
            return fr
    return f"{wt:.2f}"

def load_recap_metadata(filepath):
    wb_rec = openpyxl.load_workbook(filepath, data_only=True)
    ws_rec = wb_rec.active
    cttw_map = {}
    desc_map = {}
    dia_map  = {}
    for r in range(15, ws_rec.max_row + 1):
        st = ws_rec.cell(r, 2).value
        ct = ws_rec.cell(r, 8).value
        ds = ws_rec.cell(r, 4).value
        dq = ws_rec.cell(r, 7).value
        if st:
            m = re.search(r'\d{6}', str(st))
            if m:
                dig = m.group(0)
                if ct: cttw_map[dig] = str(ct).strip()
                if ds: desc_map[dig] = str(ds).strip()
                if dq: dia_map[dig]  = str(dq).strip()
    wb_rec.close()
    return cttw_map, desc_map, dia_map

print("Loading existing Recap metadata...")
RECAP_CTTW, RECAP_DESC, RECAP_DIA = load_recap_metadata(RECAP_IN)

def get_center_desc(dia1_size, dia1_wt):
    s = str(dia1_size or '').upper()
    wt_int = round(float(dia1_wt)) if dia1_wt else 2
    if 'MRQ' in s:  return f"Center Marqise {wt_int}ct"
    if 'OVAL' in s: return f"Center Oval {wt_int}ct"
    if 'CUS' in s:  return f"Center Cushion {wt_int}ct"
    if 'PER' in s:  return f"Center Pear {wt_int}ct"
    if 'RND' in s:  return f"Center Round {wt_int}ct"
    return ""


# Combined lookup order: LGD → Dia → Stones (for diamond quality codes)
# For gem quality codes:  Stones → Dia → LGD
def resolve_code(code, primary, secondary, tertiary):
    """
    Looks up `code` in primary → secondary → tertiary maps.
    Returns descriptive name or original code if not found.
    """
    if not code:
        return ''
    key = code.strip().upper()
    return (primary.get(key)
            or secondary.get(key)
            or tertiary.get(key)
            or code.strip())   # fallback: raw code

# ─── Compound quality-code parser (e.g. MLGDFVS2, RLGDFVS2, RBWNDA) ─────────
# These codes embed shape + diamond-type + quality in one token, e.g.:
#   M  LGD  F VS2  →  M=Marquise, LGD=Lab Grown Diamond, F VS2=quality grade
#   RB WNDA          →  RB=Round Black, WNDA=White Natural Diamond

_SHAPE_LGD = {                 # prefix before "LGD" in compound code
    'M'  : 'Lab Marquise Diamond',
    'R'  : 'Lab Round Diamond',
    'OW' : 'Lab Oval Diamond',
    'OV' : 'Lab Oval Diamond',
    'CU' : 'Lab Cushion Diamond',
    'CUL': 'Lab Cushion Diamond',
    'PR' : 'Lab Princess Diamond',
    'PER': 'Lab Pear Diamond',
    'EM' : 'Lab Emerald Cut Diamond',
    'RA' : 'Lab Radiant Diamond',
    'A'  : 'Lab Pear Diamond',
    'AL' : 'Lab Pear Diamond',
    'OWL': 'Lab Oval Diamond',
}
_SHAPE_NDA = {                 # prefix before "WNDA" / "NDA" / "BDA"
    'RB'  : 'Round Black Diamond',
    'R'   : 'Round Diamond',
    'M'   : 'Marquise Diamond',
    'OV'  : 'Oval Diamond',
    'CU'  : 'Cushion Diamond',
    'PR'  : 'Princess Diamond',
}

def _parse_quality(suffix):
    """'FVS2' → 'F VS2',  'GVS1' → 'G VS1',  'II1' → 'I I1', etc."""
    m = re.match(r'^([A-Z]{1,2})(VS2|VS1|SI1|SI2|I1|IF|VVS1|VVS2|FL)$', suffix, re.I)
    if m:
        return f'{m.group(1).upper()} {m.group(2).upper()}'
    return suffix

def parse_compound_quality(code):
    """
    Returns a human-readable string if `code` follows a known compound pattern.
    Returns None if the code is not parseable.
    """
    c = code.strip().upper()

    # Pattern 1 – Lab Grown Diamond:  {shape}LGD{quality}  e.g. MLGDFVS2
    if 'LGD' in c:
        idx   = c.index('LGD')
        shape = c[:idx]
        qual  = c[idx + 3:]
        # Longest-match for shape prefix
        shape_desc = None
        for k in sorted(_SHAPE_LGD.keys(), key=len, reverse=True):
            if shape == k:
                shape_desc = _SHAPE_LGD[k]
                break
        if not shape_desc:
            shape_desc = f'Lab Diamond'
        qual_desc = _parse_quality(qual) if qual else ''
        return f'{shape_desc} – {qual_desc}' if qual_desc else shape_desc

    # Pattern 2 – Natural / treated:  {shape}WNDA | {shape}NDA | {shape}BDA
    for suffix in ('WNDA', 'NDA', 'BDA'):
        if c.endswith(suffix):
            prefix = c[: len(c) - len(suffix)]
            shape_desc = _SHAPE_NDA.get(prefix, f'Diamond ({prefix})')
            label = {'WNDA': 'White Natural', 'NDA': 'Natural', 'BDA': 'Black'}[suffix]
            return f'{shape_desc} ({label})'

    return None   # cannot parse


def resolve_dia_quality(code):
    """Resolve diamond quality code → description.
    Priority: compound-parser → LGD map → Dia map → raw code.
    """
    if not code:
        return ''
    parsed = parse_compound_quality(code)
    if parsed:
        return parsed
    return resolve_code(code, LGD_MAP, DIA_MAP, STONE_MAP)

def resolve_gem_quality(code):
    """Resolve gemstone quality code → description.
    Priority: Stones map → Dia map → compound-parser → raw code.
    """
    if not code:
        return ''
    result = resolve_code(code, STONE_MAP, DIA_MAP, LGD_MAP)
    if result == code.strip():        # not found in any map
        parsed = parse_compound_quality(code)
        if parsed:
            return parsed
    return result

# ─── Load & deduplicate CSV (first occurrence of each Style_No) ───────────────
seen, ordered = {}, []
with open(CSV_FILE, encoding='utf-8') as f:
    for row in csv.DictReader(f):
        sn = row['Style_No'].strip()
        if sn and sn not in seen:
            seen[sn] = row
            ordered.append(sn)

print(f"Distinct styles: {len(ordered)}")

# ─── Build records ────────────────────────────────────────────────────────────
COL_SR       = 1
COL_STYLE    = 2
COL_IMAGE    = 3
COL_DESC     = 4
COL_METAL    = 5
COL_METAL_WT = 6   # METAL WT (g)
COL_DIA      = 7   # DIAMOND QUALITY
COL_CTTW     = 8   # CTTW
COL_GEM      = 9   # Gemstone Information
COL_PRICE    = 10
COL_CMT      = 11

records = []
for sr, sn in enumerate(ordered, start=1):
    r = seen[sn]
    metal_code  = r.get('Metal', '').strip()
    dia1_val    = r.get('Dia1_Value', '').strip()
    dia1_qly    = r.get('Dia1_Qly',  '').strip()
    gem_qly     = r.get('Gem_Qly',   '').strip()
    ny_sell_raw = r.get('NY_Sell_Incl_COOP', '').strip()

    # Metal Wt (g)
    m_wt_raw = r.get('Max_Wt_GOLD', '').strip() or r.get('Max_Wt_SIL', '').strip()
    try:    metal_wt = round(float(m_wt_raw), 2)
    except: metal_wt = m_wt_raw

    # CTTW (Resolution via existing recap string or tolerance table fraction)
    m = re.search(r'\d{6}', sn)
    num = m.group(0) if m else sn
    if num in RECAP_CTTW:
        cttw = RECAP_CTTW[num]
    else:
        d1 = float(r.get('Dia1_Wt') or 0)
        d2 = float(r.get('Dia2_Wt') or 0)
        gm = float(r.get('Gem_Wt') or 0)
        tot = d1 + d2 + gm
        frac = get_fraction(tot, TOLERANCE_RANGES)

        c_desc = ""
        cnt = float(r.get('Dia1_Count') or 0)
        if cnt == 1 and d1 >= 1.0:
            c_info = get_center_desc(r.get('Dia1_Size'), d1)
            if c_info: c_desc = f"      ({c_info})"

        cttw = f"{frac} cttw{c_desc}"

    # Resolve quality codes → descriptive names via Code_Mapping
    dia1_qly_desc = resolve_dia_quality(dia1_qly)   # e.g. MLGDFVS2 → "Lab Marquise Diamond"
    gem_qly_desc  = resolve_gem_quality(gem_qly)    # e.g. OVTZ2 → "Oval Tanzanite"

    # Description: Use recap template description if available, else build dynamically
    if num in RECAP_DESC:
        description = RECAP_DESC[num]
    else:
        parts = [fmt_metal_desc(metal_code)]
        if dia1_val:       parts.append(dia1_val)
        if dia1_qly_desc:  parts.append(dia1_qly_desc)
        if gem_qly_desc:   parts.append(gem_qly_desc)
        parts.append(sn)
        description = '  '.join(p for p in parts if p)

    final_dia_qly = RECAP_DIA.get(num, dia1_qly_desc)

    try:    price = round(float(ny_sell_raw), 2)
    except: price = 'TBD' if not ny_sell_raw else ny_sell_raw

    records.append({
        'sr': sr, 'style_no': sn,
        'description': description,
        'metal'    : fmt_metal_cell(metal_code),
        'metal_wt' : metal_wt,
        'dia_qly'  : final_dia_qly,    # descriptive name for Diamond Quality col
        'cttw'     : cttw,             # CTTW with fraction
        'gem_info' : gem_qly_desc,     # descriptive name for Gemstone Info col
        'unit_price': price,
    })

# ─── Load workbook ────────────────────────────────────────────────────────────
wb  = openpyxl.load_workbook(RECAP_IN)
ws  = wb.active
MAX_COL = ws.max_column

DATA_START = 15   # first data row

# ─── 1. Update date (Row 6, col 7) ───────────────────────────────────────────
today = date.today()
ws.cell(row=6, column=7).value = today.strftime("%B, %d %Y")
print(f"Date -> {today.strftime('%B, %d %Y')}")

# ─── 2. Locate T&C row ───────────────────────────────────────────────────────
tc_row = None
for r in range(DATA_START, ws.max_row + 1):
    v = ws.cell(row=r, column=1).value
    if v and 'TERMS' in str(v).upper():
        tc_row = r
        break
assert tc_row, "Could not find TERMS & CONDITIONS row!"
print(f"T&C at row {tc_row}")

# ─── 3. Capture template styles as DEEP COPIES (before any row ops) ──────────
#  Use existing data row 15 (odd) and row 16 (even) as style templates.
template_styles  = {}   # {0: {col: style_dict}, 1: {col: style_dict}}
template_heights = {}   # {0: height, 1: height}

for alt, tr in enumerate([DATA_START, DATA_START + 1]):
    template_styles[alt]  = {}
    template_heights[alt] = ws.row_dimensions[tr].height
    for c in range(1, MAX_COL + 1):
        cell = ws.cell(row=tr, column=c)
        template_styles[alt][c] = {
            'font'         : copy.copy(cell.font)         if cell.has_style else None,
            'fill'         : copy.copy(cell.fill)         if cell.has_style else None,
            'border'       : copy.copy(cell.border)       if cell.has_style else None,
            'alignment'    : copy.copy(cell.alignment)    if cell.has_style else None,
            'number_format': cell.number_format,
        }

def apply_style(dst_cell, style_dict):
    if style_dict.get('font'):          dst_cell.font          = style_dict['font']
    if style_dict.get('fill'):          dst_cell.fill          = style_dict['fill']
    if style_dict.get('border'):        dst_cell.border        = style_dict['border']
    if style_dict.get('alignment'):     dst_cell.alignment     = style_dict['alignment']
    if style_dict.get('number_format'): dst_cell.number_format = style_dict['number_format']

# ─── 4. Target T&C row directly after data rows ───────────────────────────────
new_count     = len(records)
target_tc_row = DATA_START + new_count
diff          = target_tc_row - tc_row
print(f"Original T&C row: {tc_row}  |  New records: {new_count}  |  Target T&C row: {target_tc_row}  |  Diff: {diff:+d}")

# ─── 5. SAFE row adjustment – remove extra blank rows and preserve merges ─────
footer_merges = []

if diff != 0:
    # Collect and unmerge footer merges (at or below original tc_row)
    for m in list(ws.merged_cells.ranges):
        if m.min_row >= tc_row:
            footer_merges.append((m.min_col, m.min_row, m.max_col, m.max_row, m.coord))
            ws.unmerge_cells(m.coord)
        elif diff < 0 and m.min_row >= target_tc_row:
            # Unmerge any placeholder merges inside the rows being deleted
            ws.unmerge_cells(m.coord)

    if diff < 0:
        # Delete extra blank rows directly before T&C
        delete_count = -diff
        ws.delete_rows(target_tc_row, delete_count)
        print(f"Deleted {delete_count} extra blank row(s) starting at row {target_tc_row}")
    elif diff > 0:
        # Insert additional rows needed for data
        ws.insert_rows(tc_row, diff)
        print(f"Inserted {diff} row(s) at row {tc_row}")

    # Re-apply footer merges at their new shifted positions
    for (mc1, mr1, mc2, mr2, _) in footer_merges:
        new_coord = f"{get_column_letter(mc1)}{mr1 + diff}:{get_column_letter(mc2)}{mr2 + diff}"
        ws.merge_cells(new_coord)
        print(f"Shifted footer merge -> {new_coord}")
else:
    print("Data rows fit template perfectly. All merged cells preserved as-is.")

# ─── 6. Clear existing data area values ──────────────────────────────────────
for r in range(DATA_START, DATA_START + new_count):
    for c in range(1, MAX_COL + 1):
        ws.cell(row=r, column=c).value = None

# ─── 7. Write new data rows with copied styles ───────────────────────────────
for i, rec in enumerate(records):
    row_num = DATA_START + i
    alt     = i % 2

    # Apply template style to every cell in this row
    for c in range(1, MAX_COL + 1):
        apply_style(ws.cell(row=row_num, column=c), template_styles[alt][c])

    if template_heights[alt]:
        ws.row_dimensions[row_num].height = template_heights[alt]

    # Write values
    ws.cell(row=row_num, column=COL_SR      ).value = rec['sr']
    ws.cell(row=row_num, column=COL_STYLE   ).value = rec['style_no']
    ws.cell(row=row_num, column=COL_IMAGE   ).value = ''
    ws.cell(row=row_num, column=COL_DESC    ).value = rec['description']
    ws.cell(row=row_num, column=COL_METAL   ).value = rec['metal']
    
    # Metal Wt (g)
    m_cell = ws.cell(row=row_num, column=COL_METAL_WT)
    m_cell.value = rec['metal_wt']
    if isinstance(rec['metal_wt'], (int, float)):
        m_cell.number_format = '0.00'

    ws.cell(row=row_num, column=COL_DIA     ).value = rec['dia_qly']
    ws.cell(row=row_num, column=COL_CTTW    ).value = rec['cttw']
    ws.cell(row=row_num, column=COL_GEM     ).value = rec['gem_info']
    ws.cell(row=row_num, column=COL_CMT     ).value = ''

    price      = rec['unit_price']
    price_cell = ws.cell(row=row_num, column=COL_PRICE)
    if isinstance(price, float):
        price_cell.value         = price
        price_cell.number_format = '"$"#,##0.00'
    else:
        price_cell.value = price

    print(f"  Row {row_num:>3} -> SR#{rec['sr']:>2}  {rec['style_no']:<15} | Metal Wt: {str(rec['metal_wt']):<6} | CTTW: {rec['cttw']}")

# ─── Ensure Essential Merges Are Present ──────────────────────────────────
existing_coords = {m.coord for m in ws.merged_cells.ranges}

def ensure_merged(ws, coord):
    if coord not in existing_coords:
        try:
            ws.merge_cells(coord)
            existing_coords.add(coord)
            print(f"Ensured merge: {coord}")
        except Exception:
            pass

# Header merges
ensure_merged(ws, "A5:F5")
ensure_merged(ws, "A8:C8")
ensure_merged(ws, "G8:K8")
ensure_merged(ws, "A9:F9")
ensure_merged(ws, "G9:K9")
ensure_merged(ws, "A11:G11")   # METAL LOCK PRICES
ensure_merged(ws, "H11:K11")   # TARIFF NOTICE header
ensure_merged(ws, "A12:G12")   # Gold/Silver lock price details
ensure_merged(ws, "H12:K12")   # Tariff details

# Footer merges (find dynamic current T&C row)
cur_tc_row = None
for r in range(DATA_START, ws.max_row + 1):
    v = ws.cell(row=r, column=1).value
    if v and 'TERMS' in str(v).upper():
        cur_tc_row = r
        break

if cur_tc_row:
    ensure_merged(ws, f"A{cur_tc_row}:K{cur_tc_row}")
    ensure_merged(ws, f"A{cur_tc_row+1}:K{cur_tc_row+1}")
    ensure_merged(ws, f"A{cur_tc_row+3}:F{cur_tc_row+3}")
    ensure_merged(ws, f"H{cur_tc_row+3}:K{cur_tc_row+3}")

# ─── 8. Save Workbook ─────────────────────────────────────────────────────────
os.makedirs(os.path.dirname(RECAP_OUT), exist_ok=True)
try:
    wb.save(RECAP_OUT)
    print(f"\n[OK] Saved successfully -> {RECAP_OUT}")
except PermissionError:
    alt_out = RECAP_OUT.replace(".xlsx", "_new.xlsx")
    wb.save(alt_out)
    print(f"\n[NOTICE] '{os.path.basename(RECAP_OUT)}' is currently open in Excel.")
    print(f"Saved update to fallback file -> {alt_out}")
    print("Close the file in Excel if you want to overwrite the primary file.")

