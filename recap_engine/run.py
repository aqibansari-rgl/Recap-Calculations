"""
recap_engine.run
================
CLI Entrypoint for executing the OOP-rich Recap Engine across all pricing sheets.
Runs seamlessly on Linux servers, macOS, and Windows without MS Excel.
"""

import os
import sys
import argparse
from typing import List, Dict, Any

# Ensure UTF-8 output encoding
if sys.stdout and hasattr(sys.stdout, 'reconfigure'):
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass

# Ensure parent directory is in Python path for package execution
PACKAGE_DIR = os.path.dirname(os.path.abspath(__file__))
PROJECT_ROOT = os.path.abspath(os.path.join(PACKAGE_DIR, ".."))
if PROJECT_ROOT not in sys.path:
    sys.path.insert(0, PROJECT_ROOT)

from recap_engine.pipeline import RecapPipeline, PipelineResult

# ─── Registered Project Mappings ──────────────────────────────────────────────
BASE_TEMPLATE_PATH = os.path.join(PROJECT_ROOT, "static", "Recap-Template.xlsx")

PROJECT_PAIRS: List[Dict[str, Any]] = [
    {
        "id": 1,
        "name": "Belle's 35th Anniversary",
        "customer": "ZALES ENCHANTED BRIDAL",
        "folder": "Belles_35th_Anniversary",
        "pricing_file": "Belle's 35th Anniversary Spring test pricing.xlsx",
        "recap_out_name": "Belles_35th_Anniversary_recap_UPDATED.xlsx",
    },
    {
        "id": 2,
        "name": "Zales Morganite Bridal",
        "customer": "ZALES ENCHANTED BRIDAL",
        "folder": "Zales_Morganite",
        "pricing_file": "Zales Morganite Bridal pricing.xlsx",
        "recap_out_name": "Zales_Morganite_Bridal_recap_UPDATED.xlsx",
    },
    {
        "id": 3,
        "name": "Zales Disney Bridal",
        "customer": "ZALES ENCHANTED BRIDAL",
        "folder": "Zales_Disney",
        "pricing_file": "Zales Dinsey Bridal pricing dtd 29.09.xlsx",
        "recap_out_name": "Zales_Disney_Bridal_recap_UPDATED.xlsx",
    },
    {
        "id": 4,
        "name": "Zales Canada Gemstone CORE",
        "customer": "ZALES CANADA",
        "folder": "Zales_Canada_Gemstone_CORE",
        "pricing_file": "Zales Canada- Gemstone- CORE- 9-15 Meeting Pricing.xls",
        "recap_out_name": "Zales_Canada_Gemstone_CORE_recap_UPDATED.xlsx",
    },
    {
        "id": 5,
        "name": "Zales Canada Diamond CORE",
        "customer": "ZALES CANADA",
        "folder": "Zales_Canada_Diamond_CORE",
        "pricing_file": "Zales Canada- Diamond- PRICING-  9-15 Meeting- Christie Byrd.xlsx",
        "recap_out_name": "Zales_Canada_Diamond_CORE_recap_UPDATED.xlsx",
    },
]


def print_matrix(template_file: str = "static/Recap-Template.xlsx") -> None:
    """Displays the registered file mapping matrix."""
    print("=" * 105)
    print("                      RECAP ENGINE — PROJECT MAPPING MATRIX")
    print(f"               Base Template: {template_file}")
    print("=" * 105)
    print(f"{'#':<3} | {'Project Name':<28} | {'Pricing Sheet (Input)':<38} | {'Customer'}")
    print("-" * 105)
    for p in PROJECT_PAIRS:
        p_name = p['pricing_file']
        if len(p_name) > 38:
            p_name = p_name[:35] + "..."
        cust = p.get('customer', 'Default')
        print(f"{p['id']:<3} | {p['name']:<28} | {p_name:<38} | {cust}")
    print("=" * 105 + "\n")


def print_summary(results: List[PipelineResult]) -> None:
    """Prints a structured execution summary table."""
    print("\n" + "=" * 95)
    print("                          RECAP ENGINE — EXECUTION SUMMARY")
    print("=" * 95)
    print(f"{'#':<3} | {'Project Name':<28} | {'Status':<10} | {'Styles':<8} | {'Time (s)':<10} | Output")
    print("-" * 95)
    total_time = 0.0
    for r in results:
        total_time += r.elapsed_seconds
        out_name = os.path.basename(r.output_recap) if r.output_recap else (r.error_message or "")
        print(f"{r.pair_id:<3} | {r.name:<28} | {r.status:<10} | {str(r.styles_count):<8} | {str(r.elapsed_seconds):<10} | {out_name}")
    print("=" * 95)
    print(f"Total processing time: {round(total_time, 2)}s across {len(results)} project(s).")
    print("=" * 95 + "\n")


def main() -> None:
    parser = argparse.ArgumentParser(description="OOP Recap Engine Runner for Jewelry Pricing & Recaps.")
    parser.add_argument("--list", action="store_true", help="List registered project pairs and exit.")
    parser.add_argument("--pair", type=int, choices=[1, 2, 3, 4, 5], help="Run a specific project by ID (1 to 5).")
    parser.add_argument("--template", default=BASE_TEMPLATE_PATH, help="Path to base recap template .xlsx")
    args = parser.parse_args()

    print_matrix(os.path.basename(args.template))
    if args.list:
        return

    mapping_file = os.path.join(PROJECT_ROOT, "files_BT", "Code_Mapping.xlsx")
    tolerance_file = os.path.join(PROJECT_ROOT, "files_BT", "tollerance.csv")

    pipeline = RecapPipeline(
        mapping_file=mapping_file,
        tolerance_csv=tolerance_file
    )

    if args.pair:
        target = [p for p in PROJECT_PAIRS if p['id'] == args.pair]
        results = pipeline.process_all(target, PROJECT_ROOT, default_template=args.template)
    else:
        results = pipeline.process_all(PROJECT_PAIRS, PROJECT_ROOT, default_template=args.template)

    print_summary(results)


if __name__ == "__main__":
    main()
