import zipfile
import xml.etree.ElementTree as ET
import openpyxl
import re

PRICING_FILE = r"C:\Users\ASUS\Documents\Manufacturing\3Marketing\Recap\files_BT\Zales Dinsey Bridal pricing dtd 29.09.xlsx"
RECAP_IN     = r"C:\Users\ASUS\Documents\Manufacturing\3Marketing\Recap\files_BT\Zales Disney Bridal recap 29.09.xlsx"

wb_pr = openpyxl.load_workbook(PRICING_FILE, data_only=True)
ws_pr = wb_pr.active

# 1. Parse all items from Pricing sheet
items = []
for r in range(4, ws_pr.max_row + 1):
    val_sr = ws_pr.cell(r, 1).value
    val_sn = ws_pr.cell(r, 2).value
    if val_sr is not None and isinstance(val_sr, (int, float)) and val_sn:
        items.append({
            'row': r,
            'sr': int(val_sr),
            'style_no': str(val_sn).strip()
        })

print(f"Total individual style items found in pricing sheet: {len(items)}")

# 2. Extract rightBrace shapes from drawing XML
with zipfile.ZipFile(PRICING_FILE, "r") as z:
    content = z.read("xl/drawings/drawing1.xml")
    root = ET.fromstring(content)
    ns = {
        "xdr": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
        "a": "http://schemas.openxmlformats.org/drawingml/2006/main"
    }
    
    braces = []
    for anchor in root:
        f = anchor.find("xdr:from", ns)
        t = anchor.find("xdr:to", ns)
        sp = anchor.find("xdr:sp", ns)
        if sp is not None:
            prstGeom = sp.find(".//a:prstGeom", ns)
            geom = prstGeom.attrib.get("prst", "") if prstGeom is not None else ""
            if geom == "rightBrace":
                f_r = int(f.find("xdr:row", ns).text) + 1  # 1-indexed Excel row
                t_r = int(t.find("xdr:row", ns).text) + 1
                braces.append((f_r, t_r))

# Sort braces by start row
braces.sort()

# Remove duplicate braces that cover the same range
unique_braces = []
for b in braces:
    if not unique_braces or abs(b[0] - unique_braces[-1][0]) > 3:
        unique_braces.append(b)

print(f"\nUnique rightBrace groupings found: {len(unique_braces)}\n")

for idx, (f_r, t_r) in enumerate(unique_braces, 1):
    # Find all items whose header row is within [f_r - 1, t_r]
    # (Allowing 1 row tolerance for anchor top)
    contained_items = [it for it in items if (f_r - 2) <= it['row'] <= t_r]
    
    # Find BS code in Column 5
    bs_code = None
    tot_dia_wt = None
    for r in range(f_r, t_r + 2):
        c5 = ws_pr.cell(r, 5).value
        if c5 and str(c5).strip().startswith("BS"):
            bs_code = str(c5).strip()
        c13 = ws_pr.cell(r, 13).value
        if c13 and "TOTAL DIA WT" in str(c13).upper():
            tot_dia_wt = ws_pr.cell(r, 14).value

    item_names = [f"Sr {it['sr']} ({it['style_no']})" for it in contained_items]
    print(f"Group #{idx}:")
    print(f"  Rows: {f_r} to {t_r}")
    print(f"  Group Product Code: {bs_code}")
    print(f"  Combined Total Dia Wt: {tot_dia_wt}")
    print(f"  Contains {len(contained_items)} items: {', '.join(item_names)}")
    print("-" * 60)
