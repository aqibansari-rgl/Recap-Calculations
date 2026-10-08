"""
recap_engine.parsers
====================
Direct Excel readers (cross-platform, zero MS Excel application) and pricing block
parsers extracting structured jewelry styles, memo prices, and bridal set consolidations.
"""

import os
import re
import csv
from typing import List, Dict, Any, Tuple, Optional
import openpyxl

from .models import StyleItem, MetalDetails, StoneDetails


class ExcelMatrixReader:
    """
    Reads spreadsheet files directly into a 2D matrix of raw values.
    Supports .xlsx via openpyxl and legacy .xls via xlrd.
    """

    @staticmethod
    def read(file_path: str) -> List[List[Any]]:
        ext = os.path.splitext(file_path)[1].lower()
        matrix: List[List[Any]] = []

        if ext == '.xls':
            import xlrd
            wb = xlrd.open_workbook(file_path)
            ws = wb.sheet_by_index(0)
            for r in range(ws.nrows):
                matrix.append([ws.cell_value(r, c) for c in range(ws.ncols)])
        else:
            wb = openpyxl.load_workbook(file_path, data_only=True, read_only=True)
            ws = wb.worksheets[0]
            for row in ws.iter_rows(values_only=True):
                matrix.append(list(row))
            wb.close()

        return matrix


class PricingBlockParser:
    """
    Parses raw spreadsheet rows into structured StyleItem domain models.
    Handles duplicate revised quote selection and Bridal Set box consolidations.
    """

    @staticmethod
    def _safe_str(row: List[Any], idx: int, default: str = "") -> str:
        try:
            val = row[idx]
            return str(val).strip() if val is not None else default
        except (IndexError, TypeError):
            return default

    @staticmethod
    def _safe_float(row: List[Any], idx: int, default: float = 0.0) -> float:
        try:
            val = row[idx]
            if val is None or val == "":
                return default
            return float(val)
        except (IndexError, TypeError, ValueError):
            return default

    @staticmethod
    def _get_6digit_key(text: Any) -> str:
        m = re.search(r'\d{6}', str(text or ''))
        return m.group(0) if m else str(text or '').strip()

    def parse_block(self, block: List[List[Any]]) -> StyleItem:
        """Parses a multi-row block representing a single style into a StyleItem."""
        r0 = block[0]
        r1 = block[1] if len(block) > 1 else []
        r2 = block[2] if len(block) > 2 else []

        # 1. Identity & Metal
        metal = MetalDetails(
            code=self._safe_str(r0, 3),
            gold_weight=self._safe_float(r1, 5),
            silver_weight=self._safe_float(r0, 5),
            gold_alloy=self._safe_str(r1, 3),
            total_value=self._safe_float(r0, 7)
        )

        # 2. Diamonds & Gemstones
        dia1 = StoneDetails(
            label="Dia1",
            shape_size=self._safe_str(r0, 8),
            range_code=self._safe_str(r0, 9),
            dimension_mm=self._safe_str(r0, 10),
            count=self._safe_float(r0, 11),
            quality_code=self._safe_str(r0, 12),
            carat_weight=self._safe_float(r0, 13),
            rate=self._safe_float(r0, 14),
            value=self._safe_float(r0, 15)
        )

        dia2 = StoneDetails(
            label="Dia2",
            shape_size=self._safe_str(r1, 8),
            range_code=self._safe_str(r1, 9),
            dimension_mm=self._safe_str(r1, 10),
            count=self._safe_float(r1, 11),
            quality_code=self._safe_str(r1, 12),
            carat_weight=self._safe_float(r1, 13),
            rate=self._safe_float(r1, 14),
            value=self._safe_float(r1, 15)
        )

        gem = StoneDetails(
            label=self._safe_str(r2, 2) or "Gem",
            shape_size=self._safe_str(r2, 8),
            range_code=self._safe_str(r2, 9),
            dimension_mm=self._safe_str(r2, 10),
            count=self._safe_float(r2, 11),
            quality_code=self._safe_str(r2, 12),
            carat_weight=self._safe_float(r2, 13),
            rate=self._safe_float(r2, 14),
            value=self._safe_float(r2, 15)
        )

        # 3. Pricing - Check MEMO COST / WITH TARIFF MEMO COST
        memo_price: Optional[float] = None
        for row in block:
            for c_idx, val in enumerate(row):
                if val and any(k in str(val).upper() for k in ['MEMO COST', 'WITH TARIFF MEMO COST']):
                    val_cand = None
                    if len(row) > 50 and isinstance(row[50], (int, float)):
                        val_cand = row[50]
                    elif c_idx + 1 < len(row) and isinstance(row[c_idx + 1], (int, float)):
                        val_cand = row[c_idx + 1]
                    elif len(row) > 49 and isinstance(row[49], (int, float)):
                        val_cand = row[49]
                    if val_cand is not None:
                        try:
                            memo_price = float(val_cand)
                            break
                        except (ValueError, TypeError):
                            pass
            if memo_price is not None:
                break

        # Check for character name (identified by prefix "disney-")
        char_name = ""
        for row in block:
            for cell in row:
                if cell:
                    c_str = str(cell).strip()
                    if c_str.lower().startswith('disney-'):
                        char_name = c_str[7:].strip().title()
                        break
            if char_name:
                break

        # Base NY Selling Incl Coop from row 0 (Col 51 index 50 or Col 50 index 49)
        base_coop = self._safe_float(r0, 50) or self._safe_float(r0, 49)

        item = StyleItem(
            style_no=self._safe_str(r0, 1),
            sr_no=self._safe_str(r0, 0),
            ref=self._safe_str(r1, 1).replace('Ref :', '').replace('Ref', '').strip(),
            size=self._safe_str(r0, 2),
            metal=metal,
            dia1=dia1,
            dia2=dia2,
            gem=gem,
            memo_price=memo_price,
            ny_sell_coop=base_coop if base_coop > 0 else None,
            components_info=self._safe_str(r2, 29),
            character_name=char_name
        )
        return item

    def parse_file(self, file_path: str) -> List[StyleItem]:
        """
        Parses all blocks in a pricing file, picks the revised duplicate occurrence,
        and consolidates Bridal Set pairs.
        """
        matrix = ExcelMatrixReader.read(file_path)
        if len(matrix) < 3:
            return []

        # Split into style row blocks by numeric Sr No
        raw_blocks: List[Tuple[List[List[Any]], int]] = []
        current_block: List[List[Any]] = []
        block_start = -1

        for i, row in enumerate(matrix):
            if i < 2:
                continue
            col_a = row[0] if row else None
            is_new_sr = False
            if col_a is not None and col_a != "":
                if isinstance(col_a, (int, float)) and col_a > 0:
                    is_new_sr = True
                elif str(col_a).strip().replace('.0', '').isdigit():
                    is_new_sr = True

            if is_new_sr:
                if current_block:
                    raw_blocks.append((current_block, block_start))
                current_block = [row]
                block_start = i + 1
            elif current_block:
                current_block.append(row)

        if current_block:
            raw_blocks.append((current_block, block_start))

        # First pass: Group raw blocks by 6-digit style key into occurrences
        occurrences: Dict[str, List[Any]] = {}
        ordered_keys: List[str] = []

        idx = 0
        while idx < len(raw_blocks):
            blk, start_row = raw_blocks[idx]
            sn = str(blk[0][1] or '').strip() if blk and len(blk[0]) > 1 else ''
            key = self._get_6digit_key(sn)

            # Check if this is a Bridal Set pair (Col E starts with BS)
            bs_codes = [r[4] for r in blk if len(r) > 4 and r[4] and str(r[4]).strip().startswith('BS')]
            if bs_codes and (idx + 1) < len(raw_blocks):
                bs_name = str(bs_codes[0]).strip()
                bs_key = self._get_6digit_key(bs_name)
                next_blk, next_start = raw_blocks[idx + 1]

                if bs_key not in occurrences:
                    occurrences[bs_key] = []
                    ordered_keys.append(bs_key)
                occurrences[bs_key].append((True, bs_name, blk, next_blk, start_row))
                idx += 2
            else:
                if key not in occurrences:
                    occurrences[key] = []
                    ordered_keys.append(key)
                occurrences[key].append((False, sn, blk, None, start_row))
                idx += 1

        # Second pass: Select Occurrence 2 (revised quote with actual production weights)
        final_items: List[StyleItem] = []

        for key in ordered_keys:
            occ_list = occurrences[key]
            # Pick index 1 (the first duplicate row) if multiple exist, otherwise index 0
            chosen = occ_list[1] if len(occ_list) >= 2 else occ_list[0]
            is_bs, style_name, blk, next_blk, start_row = chosen

            if is_bs and next_blk:
                re_item = self.parse_block(blk)
                rb_item = self.parse_block(next_blk)

                # Look for BOX SET price in next_blk
                box_prices = [r[50] for r in next_blk if len(r) > 50 and r[48] and 'BOX SET' in str(r[48]).upper()]
                box_price = float(box_prices[0]) if box_prices and isinstance(box_prices[0], (int, float)) else None

                # Look for TOTAL DIA WT in next_blk
                tot_dia_wts = [r[13] for r in next_blk if len(r) > 13 and r[12] and 'TOTAL DIA WT' in str(r[12]).upper()]
                tot_dia_wt = float(tot_dia_wts[0]) if tot_dia_wts and isinstance(tot_dia_wts[0], (int, float)) else None

                # Create consolidated Bridal Set
                re_item.style_no = style_name
                re_item.is_bridal_set = True
                re_item.box_price = box_price
                re_item.components_info = f"Bridal Set ({blk[0][1]} + {next_blk[0][1]})"
                re_item.character_name = re_item.character_name or rb_item.character_name

                # Sum metal weights
                re_item.metal.gold_weight = round(re_item.metal.gold_weight + rb_item.metal.gold_weight, 2)
                re_item.metal.silver_weight = round(re_item.metal.silver_weight + rb_item.metal.silver_weight, 2)

                if tot_dia_wt:
                    re_item.dia1.carat_weight = round(tot_dia_wt, 2)

                final_items.append(re_item)
            else:
                final_items.append(self.parse_block(blk))

        return final_items

    @staticmethod
    def export_to_csv(items: List[StyleItem], output_csv: str) -> None:
        """Exports parsed StyleItem models to structured_styles.csv for review."""
        os.makedirs(os.path.dirname(output_csv), exist_ok=True)
        headers = [
            'Sr_No', 'Style_No', 'Character_Name', 'Ref', 'Size', 'Metal', 'Max_Wt_GOLD', 'Max_Wt_SIL',
            'Dia1_Size', 'Dia1_Count', 'Dia1_Qly', 'Dia1_Wt', 'Dia1_Value',
            'Dia2_Size', 'Dia2_Count', 'Dia2_Qly', 'Dia2_Wt',
            'Gem_Label', 'Gem_Count', 'Gem_Qly', 'Gem_Wt',
            'Memo_Price', 'NY_Sell_Incl_COOP', 'Effective_Price', 'Components', 'Image_File'
        ]

        try:
            target_path = output_csv
            f = open(target_path, 'w', newline='', encoding='utf-8')
        except PermissionError:
            target_path = output_csv.replace(".csv", "_new.csv")
            try:
                f = open(target_path, 'w', newline='', encoding='utf-8')
            except PermissionError:
                return

        with f:
            writer = csv.DictWriter(f, fieldnames=headers)
            writer.writeheader()
            for item in items:
                writer.writerow({
                    'Sr_No': item.sr_no,
                    'Style_No': item.style_no,
                    'Character_Name': item.character_name,
                    'Ref': item.ref,
                    'Size': item.size,
                    'Metal': item.metal.code,
                    'Max_Wt_GOLD': str(item.metal.gold_weight) if item.metal.gold_weight > 0 else '',
                    'Max_Wt_SIL': str(item.metal.silver_weight) if item.metal.silver_weight > 0 else '',
                    'Dia1_Size': item.dia1.shape_size,
                    'Dia1_Count': str(item.dia1.count) if item.dia1.count > 0 else '',
                    'Dia1_Qly': item.dia1.quality_code,
                    'Dia1_Wt': str(item.dia1.carat_weight) if item.dia1.carat_weight > 0 else '',
                    'Dia1_Value': str(item.dia1.value) if item.dia1.value > 0 else '',
                    'Dia2_Size': item.dia2.shape_size,
                    'Dia2_Count': str(item.dia2.count) if item.dia2.count > 0 else '',
                    'Dia2_Qly': item.dia2.quality_code,
                    'Dia2_Wt': str(item.dia2.carat_weight) if item.dia2.carat_weight > 0 else '',
                    'Gem_Label': item.gem.label,
                    'Gem_Count': str(item.gem.count) if item.gem.count > 0 else '',
                    'Gem_Qly': item.gem.quality_code,
                    'Gem_Wt': str(item.gem.carat_weight) if item.gem.carat_weight > 0 else '',
                    'Memo_Price': str(item.memo_price) if item.memo_price is not None else '',
                    'NY_Sell_Incl_COOP': str(item.ny_sell_coop) if item.ny_sell_coop is not None else '',
                    'Effective_Price': str(item.effective_unit_price),
                    'Components': item.components_info,
                    'Image_File': item.image_file,
                })
