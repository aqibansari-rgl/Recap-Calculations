import time

t0 = time.perf_counter()

# 1. Load Code_Mapping
import openpyxl, re
t_start = time.perf_counter()
from Calculation.update_recap_Fast import MAPPING_FILE, TOLERANCE_FILE, RECAP_IN, CSV_FILE, load_code_mapping, load_tolerance_ranges, load_recap_cttw

t1 = time.perf_counter()
print(f"Imports took: {t1 - t_start:.2f}s")

tm = time.perf_counter()
dia_map, lgd_map, stone_map = load_code_mapping(MAPPING_FILE)
print(f"Loading Code_Mapping.xlsx took: {time.perf_counter() - tm:.2f}s")

tt = time.perf_counter()
tol = load_tolerance_ranges(TOLERANCE_FILE)
print(f"Loading Tolerance Ranges took: {time.perf_counter() - tt:.2f}s")

tr = time.perf_counter()
rec_ct = load_recap_cttw(RECAP_IN)
print(f"Loading Recap CTTW took: {time.perf_counter() - tr:.2f}s")

tw = time.perf_counter()
wb = openpyxl.load_workbook(RECAP_IN)
print(f"Loading RECAP_IN workbook took: {time.perf_counter() - tw:.2f}s")


def restore_drawings(orig_path, target_path, out_path):
    orig_drawings = {}
    with zipfile.ZipFile(orig_path, 'r') as z_orig:
        for name in z_orig.namelist():
            if name.startswith('xl/media/') or name.startswith('xl/drawings/'):
                orig_drawings[name] = z_orig.read(name)
        orig_sheet1_rels = z_orig.read('xl/worksheets/_rels/sheet1.xml.rels').decode('utf-8')
        orig_content_types = z_orig.read('[Content_Types].xml').decode('utf-8')

    target_files = {}
    with zipfile.ZipFile(target_path, 'r') as z_tgt:
        for name in z_tgt.namelist():
            target_files[name] = z_tgt.read(name)

    # 1. Add drawings and media from original
    for k, v in orig_drawings.items():
        target_files[k] = v

    # 2. Ensure <drawing r:id="rId2"/> in sheet1.xml
    sheet1_xml = target_files['xl/worksheets/sheet1.xml'].decode('utf-8')
    if '<drawing' not in sheet1_xml:
        idx = sheet1_xml.rfind('</worksheet>')
        if idx != -1:
            sheet1_xml = sheet1_xml[:idx] + '<drawing r:id="rId2"/></worksheet>'
            target_files['xl/worksheets/sheet1.xml'] = sheet1_xml.encode('utf-8')

    # 3. Ensure sheet1.xml.rels has drawing1.xml relationship
    target_files['xl/worksheets/_rels/sheet1.xml.rels'] = orig_sheet1_rels.encode('utf-8')

    # 4. Ensure [Content_Types].xml has drawing1 and image extensions
    ct_xml = target_files['[Content_Types].xml'].decode('utf-8')
    idx = ct_xml.rfind('</Types>')
    additions = ''
    if 'PartName="/xl/drawings/drawing1.xml"' not in ct_xml:
        additions += '<Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>'
    for ext in ['png', 'jpeg', 'jpg']:
        if f'Extension="{ext}"' not in ct_xml:
            additions += f'<Default Extension="{ext}" ContentType="image/{ext}"/>'
    ct_xml = ct_xml[:idx] + additions + '</Types>'
    target_files['[Content_Types].xml'] = ct_xml.encode('utf-8')

    with zipfile.ZipFile(out_path, 'w', compression=zipfile.ZIP_DEFLATED) as z_out:
        for name, data in target_files.items():
            z_out.writestr(name, data)

    print(f"✅ Successfully created {out_path} with all images & drawings restored!")

restore_drawings(ORIG_FILE, UPD_FILE, FINAL_OUT)

with zipfile.ZipFile(FINAL_OUT, 'r') as z:
    media = [n for n in z.namelist() if 'media' in n or 'drawing' in n]
    print(f"Total media/drawing files in FINAL: {len(media)}")












