"""
batch_process_all.py — Batch Pricing & Recap Automation Pipeline
================================================================
Reads pricing sheets from 'files_BT', extracts structured styles with images,
and updates corresponding client Recap sheets in designated output folders.
"""

import os
import sys
import time
import argparse
import csv

# Set UTF-8 encoding for console output
if sys.stdout and hasattr(sys.stdout, 'reconfigure'):
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass

# Ensure both current dir and Calculation/ dir are in python path
CURR_DIR = os.path.dirname(os.path.abspath(__file__))
ROOT_DIR = os.path.abspath(os.path.join(CURR_DIR, "..")) if os.path.basename(CURR_DIR).lower() == "calculation" else CURR_DIR
CALC_DIR = os.path.join(ROOT_DIR, "Calculation")

for p in [ROOT_DIR, CALC_DIR]:
    if p not in sys.path:
        sys.path.insert(0, p)

try:
    from Calculation.read_file_pricing import convert_excel
    from Calculation.recap_main import update_recap, get_code_maps, get_tolerance_ranges
except ImportError:
    from read_file_pricing import convert_excel
    from recap_main import update_recap, get_code_maps, get_tolerance_ranges

# ─── Directories & Reference Files ─────────────────────────────────────────────
FILES_BT_DIR  = os.path.join(ROOT_DIR, "files_BT")
OUTPUTS_DIR   = os.path.join(ROOT_DIR, "outputs")
MAPPING_FILE  = os.path.join(FILES_BT_DIR, "Code_Mapping.xlsx")
TOLERANCE_CSV = os.path.join(FILES_BT_DIR, "tollerance.csv")

# ─── Verified Pricing -> Recap Mappings ────────────────────────────────────────
PAIRS = [
    {
        "id": 1,
        "name": "Belle's 35th Anniversary",
        "folder": "Belles_35th_Anniversary",
        "pricing_file": "Belle's 35th Anniversary Spring test pricing.xlsx",
        "recap_file": "Belle's 35th Anniversary Spring test recap.xlsx",
        "recap_out_name": "Belles_35th_Anniversary_recap_UPDATED.xlsx",
    },
    {
        "id": 2,
        "name": "Zales Morganite Bridal",
        "folder": "Zales_Morganite",
        "pricing_file": "Zales Morganite Bridal pricing.xlsx",
        "recap_file": "Zales Morganite Bridal Recap.xlsx",
        "recap_out_name": "Zales_Morganite_Bridal_recap_UPDATED.xlsx",
    },
    {
        "id": 3,
        "name": "Zales Disney Bridal",
        "folder": "Zales_Disney",
        "pricing_file": "Zales Dinsey Bridal pricing dtd 29.09.xlsx",
        "recap_file": "Zales Disney Bridal recap 29.09.xlsx",
        "recap_out_name": "Zales_Disney_Bridal_recap_UPDATED.xlsx",
    },
    {
        "id": 4,
        "name": "Zales Canada Gemstone CORE",
        "folder": "Zales_Canada_Gemstone_CORE",
        "pricing_file": "Zales Canada- Gemstone- CORE- 9-15 Meeting Pricing.xls",
        "recap_file": "Zales Canada- Gemstone-CORE- 9-15 Meeting Recap.xlsx",
        "recap_out_name": "Zales_Canada_Gemstone_CORE_recap_UPDATED.xlsx",
    },
    {
        "id": 5,
        "name": "Zales Canada Diamond CORE",
        "folder": "Zales_Canada_Diamond_CORE",
        "pricing_file": "Zales Canada- Diamond- PRICING-  9-15 Meeting- Christie Byrd.xlsx",
        "recap_file": "Diamond- Recap  9-15 Meeting-CORE.xlsx",
        "recap_out_name": "Zales_Canada_Diamond_CORE_recap_UPDATED.xlsx",
    },
]


def print_mapping_table():
    """Prints the registered mapping between pricing and recap files."""
    print("=" * 105)
    print("                      FILES_BT PRICING -> RECAP MAPPING MATRIX")
    print("=" * 105)
    print(f"{'#':<3} | {'Project Name':<28} | {'Pricing Sheet (Input)':<34} | {'Recap Template (Input)'}")
    print("-" * 105)
    for p in PAIRS:
        p_name = p['pricing_file']
        if len(p_name) > 34:
            p_name = p_name[:31] + "..."
        r_name = p['recap_file']
        if len(r_name) > 34:
            r_name = r_name[:31] + "..."
        print(f"{p['id']:<3} | {p['name']:<28} | {p_name:<34} | {r_name}")
    print("=" * 105)
    print("Ignored Reference Files: Code_Mapping.xlsx, tollerance.csv, Copy of Diamond Tolerance Ranges... etc.")
    print("=" * 105 + "\n")


def process_pair(pair, dia_map, lgd_map, stone_map, tolerance_ranges):
    """
    Executes conversion and recap update for a single pricing/recap pair.
    """
    p_id    = pair['id']
    name    = pair['name']
    folder  = pair['folder']
    
    pricing_path = os.path.join(FILES_BT_DIR, pair['pricing_file'])
    recap_path   = os.path.join(FILES_BT_DIR, pair['recap_file'])

    out_dir      = os.path.join(OUTPUTS_DIR, folder)
    csv_out      = os.path.join(out_dir, "structured_styles.csv")
    images_dir   = os.path.join(out_dir, "style_images")
    recap_out    = os.path.join(out_dir, pair['recap_out_name'])

    print("\n" + "#" * 90)
    print(f"  [{p_id}/{len(PAIRS)}] STARTING: {name.upper()}")
    print(f"  Pricing Input : {pair['pricing_file']}")
    print(f"  Recap Template: {pair['recap_file']}")
    print(f"  Output Folder : {out_dir}")
    print("#" * 90)

    # 1. Validation
    if not os.path.exists(pricing_path):
        print(f"[ERROR] Pricing file not found: {pricing_path}")
        return {"id": p_id, "name": name, "status": "FAILED", "reason": "Pricing file missing", "styles": 0}

    if not os.path.exists(recap_path):
        print(f"[ERROR] Recap file not found: {recap_path}")
        return {"id": p_id, "name": name, "status": "FAILED", "reason": "Recap template missing", "styles": 0}

    os.makedirs(out_dir, exist_ok=True)
    os.makedirs(images_dir, exist_ok=True)

    t_start = time.time()

    # 2. Step 1: Read pricing sheet via COM & generate structured_styles.csv
    print(f"\n--- Step 1: Parsing Pricing Sheet & Extracting Images ---")
    try:
        convert_excel(pricing_path, csv_out, images_dir)
    except Exception as e:
        print(f"[ERROR] Failed in convert_excel: {e}")
        return {"id": p_id, "name": name, "status": "FAILED", "reason": str(e), "styles": 0}

    if not os.path.exists(csv_out):
        return {"id": p_id, "name": name, "status": "FAILED", "reason": "CSV not generated", "styles": 0}

    # Count styles in CSV
    style_count = 0
    try:
        with open(csv_out, encoding='utf-8') as f:
            reader = csv.reader(f)
            header = next(reader, None)
            style_count = sum(1 for _ in reader)
    except Exception:
        pass

    # 3. Step 2: Populate Recap Workbook
    print(f"\n--- Step 2: Populating Recap Sheet ({style_count} styles) ---")
    try:
        final_recap_out = update_recap(
            recap_in=recap_path,
            csv_file=csv_out,
            recap_out=recap_out,
            dia_map=dia_map,
            lgd_map=lgd_map,
            stone_map=stone_map,
            tolerance_ranges=tolerance_ranges
        )
    except Exception as e:
        print(f"[ERROR] Failed in update_recap: {e}")
        return {"id": p_id, "name": name, "status": "FAILED", "reason": str(e), "styles": style_count}

    elapsed = round(time.time() - t_start, 1)
    print(f"\n>>> [SUCCESS] {name} completed in {elapsed}s.")
    print(f"    Recap saved: {final_recap_out}")

    return {
        "id": p_id,
        "name": name,
        "status": "SUCCESS",
        "output_recap": final_recap_out,
        "styles": style_count,
        "elapsed": elapsed
    }


def main():
    parser = argparse.ArgumentParser(description="Batch process jewelry pricing sheets into client recaps.")
    parser.add_argument("--list", action="store_true", help="Print mapped file matrix and exit.")
    parser.add_argument("--pair", type=int, choices=[1, 2, 3, 4, 5], help="Run only a specific pair by ID (1 to 5).")
    args = parser.parse_args()

    print_mapping_table()

    if args.list:
        return

    # Select target pairs
    if args.pair:
        target_pairs = [p for p in PAIRS if p['id'] == args.pair]
        print(f"Target selected: ID {args.pair} ({target_pairs[0]['name']})\n")
    else:
        target_pairs = PAIRS
        print(f"Running full batch pipeline for all {len(PAIRS)} pairs...\n")

    # Load reference lookups once for performance
    print("Loading shared reference datasets...")
    dia_map, lgd_map, stone_map = get_code_maps(MAPPING_FILE)
    tolerance_ranges = get_tolerance_ranges(TOLERANCE_CSV)
    print(f"Shared datasets loaded successfully.\n")

    overall_start = time.time()
    results = []

    for pair in target_pairs:
        res = process_pair(pair, dia_map, lgd_map, stone_map, tolerance_ranges)
        results.append(res)

    total_time = round(time.time() - overall_start, 1)

    # Print Final Summary Report
    print("\n" + "=" * 95)
    print("                              BATCH EXECUTION SUMMARY")
    print("=" * 95)
    print(f"{'#':<3} | {'Project Name':<28} | {'Status':<10} | {'Styles':<8} | {'Time (s)':<10} | Output")
    print("-" * 95)
    for r in results:
        status_str = r['status']
        styles_str = str(r.get('styles', 0))
        time_str   = str(r.get('elapsed', '-'))
        out_str    = os.path.basename(r.get('output_recap', r.get('reason', '')))
        print(f"{r['id']:<3} | {r['name']:<28} | {status_str:<10} | {styles_str:<8} | {time_str:<10} | {out_str}")
    print("=" * 95)
    print(f"Total time elapsed: {total_time}s")
    print("=" * 95 + "\n")


if __name__ == "__main__":
    main()
