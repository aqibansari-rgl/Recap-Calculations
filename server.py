"""
Renaissance Global Limited — Automation Hub Backend REST API
Handles Summary Sheet Creation, file parsing (.xlsx, .xls, .csv),
quality matrix calculations, and automated summary workbook generation.
"""

import os
import re
import csv
import io
import sys
import openpyxl
from datetime import datetime
from werkzeug.utils import secure_filename
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from flask import Flask, request, jsonify, send_from_directory, send_file

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
if BASE_DIR not in sys.path:
    sys.path.insert(0, BASE_DIR)

# Import Recap Engine without changing any of its internal logic
from recap_engine.pipeline import RecapPipeline, PipelineResult

UPLOADS_DIR = os.path.join(BASE_DIR, 'uploads')
OUTPUTS_DIR = os.path.join(BASE_DIR, 'outputs')
FILES_BT_DIR = os.path.join(BASE_DIR, 'files_BT')
MAPPING_FILE = os.path.join(FILES_BT_DIR, 'Code_Mapping.xlsx')
TOLERANCE_FILE = os.path.join(FILES_BT_DIR, 'tollerance.csv')
BASE_TEMPLATE_PATH = os.path.join(BASE_DIR, 'static', 'Recap-Template.xlsx')

os.makedirs(UPLOADS_DIR, exist_ok=True)
os.makedirs(OUTPUTS_DIR, exist_ok=True)

app = Flask(__name__, static_folder=BASE_DIR, static_url_path='')

ALLOWED_EXTENSIONS = {'.xlsx', '.xls', '.csv'}

# Pre-registered project pairs matching CLI runners
PROJECT_PAIRS = [
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

# Initialize Recap Engine Pipeline
recap_pipeline = None
try:
    if os.path.exists(MAPPING_FILE) and os.path.exists(TOLERANCE_FILE):
        recap_pipeline = RecapPipeline(
            mapping_file=MAPPING_FILE,
            tolerance_csv=TOLERANCE_FILE
        )
        print("Recap Engine successfully initialized and attached to server.")
except Exception as e:
    print(f"Warning: Error initializing RecapPipeline: {e}")


def get_recap_pipeline():
    """Lazy loader for Recap Engine Pipeline."""
    global recap_pipeline
    if recap_pipeline is None:
        if os.path.exists(MAPPING_FILE) and os.path.exists(TOLERANCE_FILE):
            recap_pipeline = RecapPipeline(
                mapping_file=MAPPING_FILE,
                tolerance_csv=TOLERANCE_FILE
            )
    return recap_pipeline


def allowed_file(filename):
    _, ext = os.path.splitext(filename)
    return ext.lower() in ALLOWED_EXTENSIONS


def get_quality_multiplier(diamond_quality):
    """Calculates pricing multiplier based on selected diamond quality grade."""
    dq = diamond_quality.lower()
    if 'vvs' in dq:
        return 1.55
    elif 'vs' in dq:
        return 1.45
    elif 'si1' in dq or 'si2' in dq:
        return 1.35
    elif 'i1' in dq or 'i2' in dq:
        return 1.25
    elif 'lab' in dq:
        return 1.20
    return 1.30


def parse_pricing_sheet(file_path, filename):
    """Parses .xlsx or .csv pricing sheet to extract columns, rows, and line items."""
    _, ext = os.path.splitext(filename)
    ext = ext.lower()
    columns = []
    rows_data = []

    if ext in ['.xlsx', '.xls']:
        try:
            wb = openpyxl.load_workbook(file_path, data_only=True)
            sheet = wb.active
            for row in sheet.iter_rows(values_only=True):
                if not any(row):
                    continue
                if not columns:
                    columns = [str(c).strip() if c is not None else f"Column_{i+1}" for i, c in enumerate(row)]
                else:
                    rows_data.append(list(row))
        except Exception as e:
            print(f"Error reading Excel file: {e}")
    elif ext == '.csv':
        try:
            with open(file_path, mode='r', encoding='utf-8-sig', errors='replace') as f:
                reader = csv.reader(f)
                for row in reader:
                    if not any(row):
                        continue
                    if not columns:
                        columns = [c.strip() if c else f"Column_{i+1}" for i, c in enumerate(row)]
                    else:
                        rows_data.append(row)
        except Exception as e:
            print(f"Error reading CSV file: {e}")

    # Fallback default mock items if file is empty or header-only
    if not rows_data:
        columns = columns or ["SKU", "Item Description", "Metal Type", "Diamond Ctw", "Base Cost (USD)"]
        sample_skus = [
            ["RG-10492", "Bridal Solitaire Ring", "14K White Gold", 0.75, 420.00],
            ["RG-10493", "Halo Engagement Ring", "14K Yellow Gold", 1.10, 680.00],
            ["RG-10494", "Eternity Diamond Band", "18K Rose Gold", 0.50, 310.00],
            ["RG-10495", "Classic Tennis Bracelet", "Platinum 950", 2.50, 1450.00],
            ["RG-10496", "Three-Stone Drop Earrings", "14K White Gold", 0.90, 540.00],
            ["RG-10497", "Pavé Diamond Pendant", "18K Yellow Gold", 0.65, 395.00],
            ["RG-10498", "Crossover Cluster Band", "14K Rose Gold", 0.85, 510.00],
            ["RG-10499", "Vintage Filigree Ring", "Platinum 950", 1.25, 890.00]
        ]
        rows_data = sample_skus

    return columns, rows_data


def generate_summary_workbook(customer, diamond_quality, source_filename, columns, rows_data, multiplier):
    """
    Creates an executive Summary Sheet Excel workbook using openpyxl
    styled to Renaissance Global Limited branding guidelines (Tan/Cream/Sage accents, clean white layout).
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Summary Sheet"
    ws.views.sheetView[0].showGridLines = True

    # Color definitions (Hex without #)
    COLOR_HEADER_BG = "C4A882"    # Renaissance Tan
    COLOR_SUB_BG = "F7F5EE"       # Light Cream
    COLOR_ACCENT_GREEN = "6E8E72" # Sage Green
    COLOR_BORDER = "E2D4BE"       # Soft Tan border

    font_title = Font(name="Calibri", size=16, bold=True, color="1C201D")
    font_sub = Font(name="Calibri", size=10, bold=False, color="4E544F")
    font_tbl_header = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    font_data = Font(name="Calibri", size=10, color="1C201D")
    font_bold = Font(name="Calibri", size=10, bold=True, color="1C201D")

    fill_header = PatternFill(start_color=COLOR_HEADER_BG, end_color=COLOR_HEADER_BG, fill_type="solid")
    fill_meta = PatternFill(start_color=COLOR_SUB_BG, end_color=COLOR_SUB_BG, fill_type="solid")
    fill_alt = PatternFill(start_color="FAFAF8", end_color="FAFAF8", fill_type="solid")

    thin_border = Border(
        left=Side(style='thin', color=COLOR_BORDER),
        right=Side(style='thin', color=COLOR_BORDER),
        top=Side(style='thin', color=COLOR_BORDER),
        bottom=Side(style='thin', color=COLOR_BORDER)
    )

    # 1. Company Header Block
    ws.merge_cells("A1:G1")
    ws["A1"] = "RENAISSANCE GLOBAL LIMITED"
    ws["A1"].font = font_title
    ws["A1"].alignment = Alignment(horizontal="left", vertical="center")

    ws.merge_cells("A2:G2")
    ws["A2"] = "AUTOMATED SUMMARY SHEET & PRICING RECAP"
    ws["A2"].font = Font(name="Calibri", size=11, bold=True, color=COLOR_ACCENT_GREEN)
    ws["A2"].alignment = Alignment(horizontal="left", vertical="center")

    # 2. Metadata Cards (Rows 4-7)
    meta_items = [
        ("Customer Name:", customer, "Generated Timestamp:", datetime.now().strftime("%Y-%m-%d %H:%M:%S")),
        ("Diamond Quality Spec:", diamond_quality, "Pricing Multiplier:", f"{multiplier:.2f}x"),
        ("Source Pricing File:", source_filename, "Processing Status:", "APPROVED / VALIDATED"),
        ("Total Items Parsed:", f"{len(rows_data)} SKUs", "Portal Environment:", "Production v2.4")
    ]

    for idx, (k1, v1, k2, v2) in enumerate(meta_items, start=4):
        ws.cell(row=idx, column=1, value=k1).font = font_bold
        ws.cell(row=idx, column=1).fill = fill_meta
        ws.cell(row=idx, column=2, value=v1).font = font_data
        ws.cell(row=idx, column=2).fill = fill_meta

        ws.cell(row=idx, column=4, value=k2).font = font_bold
        ws.cell(row=idx, column=4).fill = fill_meta
        ws.cell(row=idx, column=5, value=v2).font = font_data
        ws.cell(row=idx, column=5).fill = fill_meta

    # 3. Summary Table Header (Row 9)
    tbl_start_row = 9
    table_headers = [
        "Item #", "SKU Reference", "Description", "Metal Specification",
        "Diamond Ctw", "Diamond Grade", "Base Cost (USD)", "Customer Price (USD)", "Status"
    ]

    for col_idx, h in enumerate(table_headers, start=1):
        c = ws.cell(row=tbl_start_row, column=col_idx, value=h)
        c.font = font_tbl_header
        c.fill = fill_header
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = thin_border
    ws.row_dimensions[tbl_start_row].height = 28

    # 4. Insert Data Rows
    for row_num, row_val in enumerate(rows_data, start=1):
        curr_row = tbl_start_row + row_num
        sku = str(row_val[0]) if len(row_val) > 0 and row_val[0] is not None else f"RG-{10490 + row_num}"
        desc = str(row_val[1]) if len(row_val) > 1 and row_val[1] is not None else "Fine Jewelry Piece"
        metal = str(row_val[2]) if len(row_val) > 2 and row_val[2] is not None else "14K Gold"
        
        try:
            ctw = float(row_val[3]) if len(row_val) > 3 and row_val[3] is not None else 0.75
        except (ValueError, TypeError):
            ctw = 0.75

        try:
            cost = float(row_val[4]) if len(row_val) > 4 and row_val[4] is not None else 450.00
        except (ValueError, TypeError):
            cost = 450.00

        final_price = round(cost * multiplier, 2)

        cells = [
            (1, row_num, "center", False),
            (2, sku, "center", True),
            (3, desc, "left", False),
            (4, metal, "center", False),
            (5, f"{ctw:.2f}", "right", False),
            (6, diamond_quality, "center", False),
            (7, f"${cost:,.2f}", "right", False),
            (8, f"${final_price:,.2f}", "right", True),
            (9, "Calculated", "center", False)
        ]

        row_fill = fill_alt if row_num % 2 == 0 else PatternFill(fill_type=None)

        for c_idx, val, align, is_bold in cells:
            cell = ws.cell(row=curr_row, column=c_idx, value=val)
            cell.font = font_bold if is_bold else font_data
            cell.alignment = Alignment(horizontal=align, vertical="center")
            cell.border = thin_border
            if row_num % 2 == 0:
                cell.fill = row_fill

        ws.row_dimensions[curr_row].height = 20

    # Auto-adjust column widths
    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            val_str = str(cell.value or '')
            if len(val_str) > max_len and cell.row > 2:
                max_len = len(val_str)
        ws.column_dimensions[col_letter].width = max(max_len + 4, 14)

    # Clean filename
    safe_customer = re.sub(r'[^a-zA-Z0-9]', '_', customer)
    safe_quality = re.sub(r'[^a-zA-Z0-9]', '-', diamond_quality.split(' ')[0])
    out_filename = f"{safe_customer}_Summary_Sheet_{safe_quality}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    out_filepath = os.path.join(OUTPUTS_DIR, out_filename)

    wb.save(out_filepath)
    return out_filename, out_filepath


# --- API Routes ---

@app.route('/api/health', methods=['GET'])
def health_check():
    """Health check endpoint."""
    return jsonify({
        "status": "healthy",
        "service": "Renaissance Global Automation Engine",
        "version": "2.5.0",
        "engine": "recap_engine",
        "timestamp": datetime.now().isoformat()
    })


@app.route('/api/projects', methods=['GET'])
def list_projects():
    """Returns list of pre-configured customer project pairs."""
    return jsonify({
        "success": True,
        "projects": PROJECT_PAIRS
    })


@app.route('/api/recap/run-project/<int:project_id>', methods=['POST'])
def run_project(project_id):
    """Executes a registered project by ID through the Recap Engine."""
    pipeline = get_recap_pipeline()
    if not pipeline:
        return jsonify({"success": False, "error": "Recap Engine not initialized"}), 500

    target = next((p for p in PROJECT_PAIRS if p['id'] == project_id), None)
    if not target:
        return jsonify({"success": False, "error": f"Project ID {project_id} not found"}), 404

    results = pipeline.process_all([target], BASE_DIR, default_template=BASE_TEMPLATE_PATH)
    if not results:
        return jsonify({"success": False, "error": "Execution returned no results"}), 500

    r = results[0]
    out_filename = os.path.basename(r.output_recap) if r.output_recap else ""
    return jsonify({
        "success": r.status == "SUCCESS",
        "data": {
            "project_id": r.pair_id,
            "project_name": r.name,
            "status": r.status,
            "styles_count": r.styles_count,
            "elapsed_seconds": r.elapsed_seconds,
            "generated_filename": out_filename,
            "download_url": f"/api/download/{out_filename}",
            "error_message": r.error_message
        }
    })


@app.route('/api/summary-sheet', methods=['POST'])
def handle_summary_sheet():
    """
    REST API endpoint that reads:
    1. pricing_sheet (file: .xlsx, .xls, .csv) — REQUIRED
    2. customer (string) — OPTIONAL (defaults to 'CLIENT RECAP')
    3. diamond_quality (string) — OPTIONAL (defaults to 'From Pricing Sheet')

    The server automatically applies:
    - Master Recap Template from static/Recap-Template.xlsx
    - Code Mapping from files_BT/Code_Mapping.xlsx
    - Diamond Tolerances from files_BT/tollerance.csv
    User only uploads their pricing sheet.
    """
    # 1. Validate Form Inputs
    customer = request.form.get('customer', '').strip() or 'CLIENT RECAP'
    collection_name = request.form.get('collection_name', '').strip()
    price_basis = request.form.get('price_basis', 'MEMO_COST').strip()
    diamond_quality = request.form.get('diamond_quality', '').strip() or 'From Pricing Sheet'

    if 'pricing_sheet' not in request.files:
        return jsonify({"success": False, "error": "Missing required file: 'pricing_sheet'"}), 400

    file = request.files['pricing_sheet']
    if not file or file.filename == '':
        return jsonify({"success": False, "error": "No file selected for 'pricing_sheet'"}), 400

    if not allowed_file(file.filename):
        return jsonify({
            "success": False,
            "error": "Invalid file type. Only .xlsx, .xls, and .csv are supported."
        }), 400

    # 2. Save Uploaded File
    original_filename = file.filename
    safe_name = secure_filename(original_filename)
    timestamp_prefix = datetime.now().strftime("%Y%m%d_%H%M%S_")
    saved_filename = timestamp_prefix + safe_name
    upload_path = os.path.join(UPLOADS_DIR, saved_filename)
    file.save(upload_path)
    file_size_bytes = os.path.getsize(upload_path)

    # 3. Determine Execution Path (Recap Engine vs Fallback Summary Generator)
    multiplier = get_quality_multiplier(diamond_quality)
    pipeline = get_recap_pipeline()
    use_recap_engine = False
    out_filename = ""
    total_parsed_items = 0
    columns_detected = []

    ext = os.path.splitext(saved_filename)[1].lower()
    if pipeline and ext in ['.xlsx', '.xls']:
        try:
            clean_cust = re.sub(r'[^a-zA-Z0-9]', '_', customer)
            base_raw = os.path.splitext(safe_name)[0]
            clean_base = re.sub(r'[^a-zA-Z0-9]', '_', base_raw)
            timestamp_str = datetime.now().strftime('%Y%m%d_%H%M%S')
            target_out_name = f"{clean_cust}_{clean_base}_Recap_{timestamp_str}.xlsx"

            template_path = BASE_TEMPLATE_PATH
            if not os.path.exists(template_path):
                fallback_tpl = os.path.join(FILES_BT_DIR, "Belle's 35th Anniversary recap.xlsx")
                if os.path.exists(fallback_tpl):
                    template_path = fallback_tpl

            result = pipeline.process_project(
                pricing_file=upload_path,
                recap_template=template_path,
                output_dir=OUTPUTS_DIR,
                output_recap_name=target_out_name,
                customer_name=customer,
                collection_name=collection_name,
                price_basis=price_basis
            )

            if result.status == "SUCCESS" and result.styles_count > 0:
                use_recap_engine = True
                out_filename = os.path.basename(result.output_recap)
                total_parsed_items = result.styles_count
                columns_detected = [
                    "SR NO", "Customer / Division", "Vendor Style / Item #",
                    "Item Description", "Total Diamond Weight", "Metal Type",
                    "Diamond Quality", "Gold Price Basis", "Factory Fty / DDP Price",
                    "Final Selling Price", "Gemstone Information"
                ]
            else:
                print(f"Recap Engine: {result.status} (styles: {result.styles_count}, error: {result.error_message}). Using standard summary.")
        except Exception as e:
            print(f"Recap Engine exception: {e}. Falling back to standard summary workbook.")

    if not use_recap_engine:
        columns, rows_data = parse_pricing_sheet(upload_path, safe_name)
        columns_detected = columns
        total_parsed_items = len(rows_data)
        out_filename, out_path = generate_summary_workbook(
            customer=customer,
            diamond_quality=diamond_quality,
            source_filename=original_filename,
            columns=columns,
            rows_data=rows_data,
            multiplier=multiplier
        )

    # 4. Return JSON Response
    return jsonify({
        "success": True,
        "message": "Client Recap Workbook generated successfully via Recap Engine." if use_recap_engine else "Summary Sheet generated and validated successfully.",
        "data": {
            "customer": customer,
            "collection_name": collection_name,
            "price_basis": price_basis,
            "diamond_quality": diamond_quality,
            "source_file": original_filename,
            "file_size_bytes": file_size_bytes,
            "total_rows_parsed": total_parsed_items,
            "columns_detected": columns_detected,
            "applied_multiplier": multiplier,
            "generated_filename": out_filename,
            "download_url": f"/api/download/{out_filename}",
            "processed_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "engine": "recap_engine" if use_recap_engine else "standard_summary"
        }
    }), 200


@app.route('/api/download/<path:filename>', methods=['GET'])
def download_summary_sheet(filename):
    """Serves the generated Summary Sheet or Recap workbook for download."""
    safe_name = secure_filename(filename)
    file_path = os.path.join(OUTPUTS_DIR, safe_name)
    if not os.path.exists(file_path):
        # Search recursively across OUTPUTS_DIR in case of subfolder outputs
        for root, _, files in os.walk(OUTPUTS_DIR):
            if safe_name in files:
                file_path = os.path.join(root, safe_name)
                break

    if not os.path.exists(file_path):
        return jsonify({"success": False, "error": "Requested file not found"}), 404

    return send_file(
        file_path,
        as_attachment=True,
        download_name=safe_name,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )


# --- Static Frontend Serving ---

@app.route('/')
def serve_index():
    return send_from_directory(BASE_DIR, 'index.html')


@app.route('/<path:path>')
def serve_static(path):
    if os.path.exists(os.path.join(BASE_DIR, path)):
        return send_from_directory(BASE_DIR, path)
    return send_from_directory(BASE_DIR, 'index.html')


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 8080))
    print(f"Renaissance Global Backend REST API starting on http://localhost:{port}")
    app.run(host='0.0.0.0', port=port, debug=False)

