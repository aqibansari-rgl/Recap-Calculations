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

    def __init__(self, stone_codes: Optional[Any] = None):
        self.stone_codes = {str(c).strip().upper() for c in stone_codes} if stone_codes else set()
        self.gold_lock_rate = 4250.0
        self.silver_lock_rate = 65.0

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

    def _is_gemstone_stone(self, shape: str, qly: str, mm: str, label: str = "") -> bool:
        s = str(shape or '').upper()
        q = str(qly or '').upper()
        m = str(mm or '').upper()
        lbl = str(label or '').upper()
        if any(q.startswith(d) for d in ['LGD', 'RLGD', 'MLGD', 'OWLG', 'CULG', 'PRLG', 'WNDA', 'NDA', 'BDA']):
            return False
        if self.stone_codes and q in self.stone_codes:
            return True
        if s.startswith('STN') or s.startswith('GEM'):
            return True
        if any(gem in m or gem in lbl or gem in s for gem in ['MORGANITE', 'SAPH', 'SAPPHIRE', 'TANZANITE', 'AMETHYST', 'TOPAZ', 'RUBY', 'EMERALD', 'OPAL']):
            return True
        if any(q.startswith(g) for g in ['OVCMG', 'RDXSW', 'RDXS', 'OVTZ', 'TZ', 'CMG']):
            return True
        return False

    def parse_block(self, block: List[List[Any]]) -> StyleItem:
        """Parses a multi-row block representing a single style into a StyleItem."""
        r0 = block[0]
        r1 = block[1] if len(block) > 1 else []
        r2 = block[2] if len(block) > 2 else []

        # 1. Identity & Metal
        gw = self._safe_float(r0, 5) or self._safe_float(r1, 5)
        sw = self._safe_float(r0, 5) if str(r0[3] or '').strip().upper() in ('SIL', 'STGSIL', 'SS', '925') else 0.0
        if sw > 0:
            gw = 0.0

        metal = MetalDetails(
            code=self._safe_str(r0, 3),
            gold_weight=gw,
            silver_weight=sw,
            gold_alloy=self._safe_str(r1, 3),
            total_value=self._safe_float(r0, 7)
        )

        # 2. Dynamic Stone Scanning across all rows in the block
        diamonds: List[StoneDetails] = []
        gemstones: List[StoneDetails] = []

        for r in block:
            sh = self._safe_str(r, 8)
            rg = self._safe_str(r, 9)
            mm = self._safe_str(r, 10)
            cnt = self._safe_float(r, 11)
            qly = self._safe_str(r, 12)
            wt = self._safe_float(r, 13)
            rt = self._safe_float(r, 14)
            val = self._safe_float(r, 15)

            # Skip summary lines or rows without real stone definition
            if 'TOTAL' in sh.upper() or 'TOTAL' in qly.upper():
                continue
            if not (qly or (sh and not sh.upper().startswith('WITH DIA CUT')) or cnt > 0):
                continue

            stn = StoneDetails(
                shape_size=sh, range_code=rg, dimension_mm=mm,
                count=cnt, quality_code=qly, carat_weight=wt,
                rate=rt, value=val
            )
            if self._is_gemstone_stone(sh, qly, mm, self._safe_str(r, 2)):
                stn.label = self._safe_str(r, 2) or "Gem"
                gemstones.append(stn)
            else:
                stn.label = f"Dia{len(diamonds)+1}"
                diamonds.append(stn)

        dia1 = diamonds[0] if len(diamonds) > 0 else StoneDetails()
        dia2 = StoneDetails()
        if len(diamonds) == 2:
            dia2 = diamonds[1]
        elif len(diamonds) > 2:
            dia2 = diamonds[1]
            extra_wt = sum(d.carat_weight for d in diamonds[2:])
            extra_cnt = sum(d.count for d in diamonds[2:])
            dia2.carat_weight = round(dia2.carat_weight + extra_wt, 4)
            dia2.count += extra_cnt

        gem = gemstones[0] if gemstones else StoneDetails()
        if len(gemstones) > 1:
            mm_list = [g.dimension_mm for g in gemstones if g.dimension_mm]
            if mm_list:
                gem.dimension_mm = " & ".join(dict.fromkeys(mm_list))

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

        # Extract dynamic Metal Lock rates from pricing sheet
        gold_rate = 0.0
        silver_rate = 0.0
        for row in matrix[:25]:
            if len(row) > 4:
                code_s = str(row[3] or '').strip().upper()
                r_val = self._safe_float(row, 4)
                if r_val > 0:
                    if any(g in code_s for g in ['G14', 'G10', 'G18', 'GOLD']):
                        gold_rate = gold_rate or r_val
                    elif any(s in code_s for s in ['SIL', 'STGSIL', 'SS', '925']):
                        silver_rate = silver_rate or r_val
        self.gold_lock_rate = gold_rate or 4250.0
        self.silver_lock_rate = silver_rate or 65.0

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
                    re_item.dia2.carat_weight = 0.0

                if not re_item.gem.quality_code and rb_item.gem.quality_code:
                    re_item.gem = rb_item.gem

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
