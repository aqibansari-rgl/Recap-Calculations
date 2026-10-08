# Recap Engine (`recap_engine/`)

A modular, Object-Oriented (OOP), cross-platform Python engine for parsing jewelry pricing sheets and generating client Recap presentation workbooks.

---

## Architecture Overview

```
recap_engine/
├── __init__.py      # Package exports and public API
├── models.py        # Domain models: StyleItem, RecapRow, MetalDetails, StoneDetails
├── parsers.py       # ExcelMatrixReader (pure Python, .xlsx & .xls) & PricingBlockParser
├── resolvers.py     # QualityCodeResolver (Code_Mapping.xlsx), ToleranceResolver, DescriptionHelper
├── builder.py       # RecapWorkbookBuilder (openpyxl formatting, merge preservation)
├── pipeline.py      # RecapPipeline (end-to-end transformation coordinator)
├── run.py           # CLI runner for Linux servers and local workstations
└── README.md        # Documentation
```

---

## Key OOP Concepts

1. **Domain Models (`models.py`)**:
   - `StyleItem`: Strongly-typed representation of a style, with physical attributes, stone details, and pricing logic.
   - `StyleItem.effective_unit_price`: Encapsulates business priority (`Box Set` -> `Memo Price` -> `NY Selling incl. Coop`).
   - `MetalDetails`: Encapsulates metal weights and formats display text (`14K\nYELLOW\nGOLD`).
   - `RecapRow`: Presentation-layer data transfer object.

2. **Decoupled Readers & Parsers (`parsers.py`)**:
   - `ExcelMatrixReader`: Zero MS Excel dependency. Reads `.xlsx` (via `openpyxl`) and `.xls` (via `xlrd`) directly in memory.
   - `PricingBlockParser`: Selects Occurrence 2 (the revised production quote) for duplicates and consolidates Bridal Set pairs.

3. **Domain Resolvers (`resolvers.py`)**:
   - `QualityCodeResolver`: Translates gemological and compound codes (`MLGDFVS2`, `RBWNDA`) using `Code_Mapping.xlsx`.
   - `ToleranceResolver`: Maps continuous carat weights into commercial fractions (`2 1/5 cttw`) using `tollerance.csv`.

4. **Workbook Builder (`builder.py`)**:
   - `RecapWorkbookBuilder`: Deep-copies template styles (fonts, fills, borders, alignments, number formats) from rows 15 & 16.
   - Dynamically deletes/inserts data rows while shifting footer merged cells (`TERMS & CONDITIONS`) and preserving merged headers.

5. **Pipeline Orchestrator (`pipeline.py`)**:
   - `RecapPipeline`: Manages shared reference datasets and coordinates the execution flow.

---

## How to Run

### Command Line Interface:
```bash
# Run all 5 projects
python recap_engine/run.py

# Run a specific project pair by ID (1 to 5)
python recap_engine/run.py --pair 1

# View project mapping matrix without executing
python recap_engine/run.py --list
```

### Programmatic Usage in Python:
```python
from recap_engine import RecapPipeline

pipeline = RecapPipeline(
    mapping_file="files_BT/Code_Mapping.xlsx",
    tolerance_csv="files_BT/tollerance.csv"
)

result = pipeline.process_project(
    pricing_file="files_BT/Zales Morganite Bridal pricing.xlsx",
    recap_template="files_BT/Zales Morganite Bridal Recap.xlsx",
    output_dir="outputs/Zales_Morganite",
    output_recap_name="Zales_Morganite_Bridal_recap_UPDATED.xlsx",
    project_name="Zales Morganite"
)

print(result.status, result.styles_count, result.output_recap)
```

---

## Linux Server Deployment

Requirements:
```bash
pip install openpyxl xlrd
```
No graphical interface or Microsoft Office required. Works 100% headless.
