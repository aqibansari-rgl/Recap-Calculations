"""
update_recap.py  —  Fixed + Code_Mapping version
Fixes the merged-cell row corruption bug by:
  1. Capturing template styles as deep copies BEFORE any row manipulation
  2. Only inserting the EXTRA rows needed (not delete-all + insert-all)
Also resolves diamond / LGD / stone quality codes to descriptive names
using Code_Mapping.xlsx (sheets: Dia, LGD, Stones).
"""

import csv, copy, os, re
from openpyxl.utils import range_boundaries, get_column_letter
import openpyxl
from datetime import date

# ─── Paths ────────────────────────────────────────────────────────────────────
MAPPING_FILE=r"C:\Users\ASUS\Documents\Manufacturing\3Marketing\Recap\files_BT\Code_Mapping.xlsx"
TOLERANCE_FILE = r"C:\Users\ASUS\Documents\Manufacturing\3Marketing\Recap\files_BT\Copy of Diamond Tolerance Ranges -Signet Vs Ren.xlsx"
RECAP_IN  =  r"C:\Users\ASUS\Documents\Manufacturing\3Marketing\Recap\files_BT\Zales Disney Bridal recap 29.09.xlsx"
CSV_FILE  =  r"C:\Users\ASUS\Documents\Manufacturing\3Marketing\Recap\outputs\Zales_Disney\structured_styles.csv"
RECAP_OUT =  r"C:\Users\ASUS\Documents\Manufacturing\3Marketing\Recap\outputs\Zales_Disney\Zales_Disney_Bridal_recap_UPDATED.xlsx"

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

# ─── Load Tolerance Ranges (Copy of Diamond Tolerance Ranges -Signet Vs Ren.xlsx)
def load_tolerance_ranges(filepath):
    wb_tol = openpyxl.load_workbook(filepath, read_only=True, data_only=True)
    ws_tol = wb_tol['Table 1-US TOLERANCE']
    ranges = []
    for r in range(4, ws_tol.max_row + 1):
        mn = ws_tol.cell(r, 1).value
        mx = ws_tol.cell(r, 2).value
        fr = ws_tol.cell(r, 6).value
        if mn is not None and mx is not None and fr is not None and str(fr).lower() != 'no fraction':
            fr_str = re.sub(r'\s+', ' ', str(fr).strip())
            ranges.append((float(mn), float(mx), fr_str))
    wb_tol.close()
    return ranges

print("Loading Tolerance Ranges...")
TOLERANCE_RANGES = load_tolerance_ranges(TOLERANCE_FILE)

def get_fraction(wt, ranges):
    if not wt or wt <= 0: return ''
    for mn, mx, fr in ranges:
        if mn <= wt <= mx:
            return fr
    return f"{wt:.2f}"

def load_recap_cttw(filepath):
    wb_rec = openpyxl.load_workbook(filepath, data_only=True)
    ws_rec = wb_rec.active
    cttw_map = {}
    for r in range(15, ws_rec.max_row + 1):
        st = ws_rec.cell(r, 2).value
        ct = ws_rec.cell(r, 8).value
        if st and ct:
            m = re.search(r'\d{6}', str(st))
            if m:
                cttw_map[m.group(0)] = str(ct).strip()
    wb_rec.close()
    return cttw_map

print("Loading existing Recap CTTW...")
RECAP_CTTW = load_recap_cttw(RECAP_IN)

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

    # Description: Metal + Dia1_Value + Dia1_Qly(desc) + Gem_Qly(desc) + Style_No
    parts = [fmt_metal_desc(metal_code)]
    if dia1_val:       parts.append(dia1_val)
    if dia1_qly_desc:  parts.append(dia1_qly_desc)
    if gem_qly_desc:   parts.append(gem_qly_desc)
    parts.append(sn)
    description = '  '.join(p for p in parts if p)

    try:    price = round(float(ny_sell_raw), 2)
    except: price = 'TBD' if not ny_sell_raw else ny_sell_raw

    records.append({
        'sr': sr, 'style_no': sn,
        'description': description,
        'metal'    : fmt_metal_cell(metal_code),
        'metal_wt' : metal_wt,
        'dia_qly'  : dia1_qly_desc,    # descriptive name for Diamond Quality col
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
print(f"Date → {today.strftime('%B, %d %Y')}")

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

# ─── 4. Find last actual data row (skip blanks before T&C) ───────────────────
last_data_row = tc_row - 1
while last_data_row >= DATA_START:
    if any(ws.cell(row=last_data_row, column=c).value for c in [1, 2]):
        break
    last_data_row -= 1

actual_count = last_data_row - DATA_START + 1
new_count    = len(records)
diff         = new_count - actual_count
print(f"Existing data rows: {actual_count}  |  New: {new_count}  |  Diff: {diff:+d}")

# ─── 5. SAFE row adjustment – manually handle merged cells ───────────────────
# openpyxl's insert_rows / delete_rows does NOT reliably shift merged ranges.
# We must: (a) unmerge everything at/after the insertion point,
#          (b) do the row op, then (c) re-merge with corrected row numbers.

insert_at = last_data_row + 1   # first row after current data (before blank/T&C)

# Collect all existing merge ranges
all_merges = [(range_boundaries(str(m)), str(m)) for m in ws.merged_cells.ranges]

# Split into: keep-as-is (entirely above insert_at) vs. must-shift
merges_above  = []  # (min_col, min_row, max_col, max_row)
merges_shift  = []  # will have rows bumped by diff

for (min_col, min_row, max_col, max_row), _ in all_merges:
    if diff > 0:
        if min_row >= insert_at:
            merges_shift.append((min_col, min_row, max_col, max_row))
        else:
            merges_above.append((min_col, min_row, max_col, max_row))
    elif diff < 0:
        delete_start = insert_at + diff
        delete_end   = insert_at - 1
        if min_row > delete_end:
            merges_shift.append((min_col, min_row, max_col, max_row))
        else:
            merges_above.append((min_col, min_row, max_col, max_row))

# Unmerge ALL cells (we will re-apply everything after the row operation)
for _, rng_str in all_merges:
    ws.unmerge_cells(rng_str)

if diff > 0:
    ws.insert_rows(insert_at, diff)
    print(f"Inserted {diff} rows at row {insert_at}")
elif diff < 0:
    ws.delete_rows(insert_at + diff, -diff)
    print(f"Deleted {-diff} rows starting at row {insert_at + diff}")

# Re-apply merges (shift the ones that needed shifting)
def _merge_str(mc1, mr1, mc2, mr2):
    return f"{get_column_letter(mc1)}{mr1}:{get_column_letter(mc2)}{mr2}"

for (mc1, mr1, mc2, mr2) in merges_above:
    ws.merge_cells(_merge_str(mc1, mr1, mc2, mr2))

for (mc1, mr1, mc2, mr2) in merges_shift:
    ws.merge_cells(_merge_str(mc1, mr1 + diff, mc2, mr2 + diff))

print(f"Re-applied {len(merges_above)} unchanged + {len(merges_shift)} shifted merge(s)")

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

    print(f"  Row {row_num:>3} → SR#{rec['sr']:>2}  {rec['style_no']:<15} | Metal Wt: {str(rec['metal_wt']):<6} | CTTW: {rec['cttw']}")


# ─── 8. Save ──────────────────────────────────────────────────────────────────
os.makedirs(os.path.dirname(RECAP_OUT), exist_ok=True)
wb.save(RECAP_OUT)
print(f"\n✅ Saved → {RECAP_OUT}")
