import win32com.client as win32
import os
import csv
import time

def parse_block(block, header1, header2):
    """
    Parses a block of rows belonging to a single Style entry into a flat dict.
    Block structure (0-indexed from block start):
      Row 0: Main row (Sr, Style No, Size, Metal, Rate, Max Wt, Met Val, Total Met Val, Dia rows, Set type, charges, pricing...)
      Row 1: Second metal row (Ref, G10Y metal, 2nd diamond row)
      Row 2: Tanzanite / Gemstone row + pricing type label
      Row 3: Internal/Spec/Production label + Dia Approval + Tone
      Row 4: WaxWt + BackPlate + Surface Finish
      Row 5: ConvWt + ConvWt value + Exclusive
      Row 6: SKU#
    """
    record = {}

    def safe(row, idx, default=''):
        try:
            v = row[idx]
            return str(v).strip() if v is not None else default
        except (IndexError, TypeError):
            return default

    def safenum(row, idx, default=''):
        try:
            v = row[idx]
            if v is None: return default
            if isinstance(v, float) and v == int(v): return str(int(v))
            return str(v)
        except (IndexError, TypeError):
            return default

    if not block:
        return record

    r0 = block[0]  # Main data row
    r1 = block[1] if len(block) > 1 else []
    r2 = block[2] if len(block) > 2 else []
    r3 = block[3] if len(block) > 3 else []
    r4 = block[4] if len(block) > 4 else []
    r5 = block[5] if len(block) > 5 else []
    r6 = block[6] if len(block) > 6 else []

    # ---- Identity ----
    record['Sr_No']       = safenum(r0, 0)
    record['Style_No']    = safe(r0, 1)
    record['Ref']         = safe(r1, 1).replace('Ref :', '').strip() if safe(r1, 1).startswith('Ref') else ''
    record['Size']        = safe(r0, 2)
    record['Metal']       = safe(r0, 3)
    record['Gold_Alloy']  = safe(r1, 3)  # e.g., G10Y

    # ---- Metal Pricing ----
    record['Rate_SIL']      = safenum(r0, 4)
    record['Max_Wt_SIL']    = safenum(r0, 5)
    record['Met_Val_SIL']   = safenum(r0, 6)
    record['Rate_GOLD']     = safenum(r1, 4)
    record['Max_Wt_GOLD']   = safenum(r1, 5)
    record['Met_Val_GOLD']  = safenum(r1, 6)
    record['Total_Metal_Value'] = safenum(r0, 7)

    # ---- Diamond 1 (Row 0) ----
    record['Dia1_Size']   = safe(r0, 8)
    record['Dia1_Range']  = safe(r0, 9)
    record['Dia1_MM']     = safe(r0, 10)
    record['Dia1_Count']  = safenum(r0, 11)
    record['Dia1_Qly']    = safe(r0, 12)
    record['Dia1_Wt']     = safenum(r0, 13)
    record['Dia1_Rate']   = safenum(r0, 14)
    record['Dia1_Value']  = safenum(r0, 15)
    record['Dia1_Total_Val'] = safenum(r0, 16)

    # ---- Diamond 2 (Row 1) ----
    record['Dia2_Size']   = safe(r1, 8)
    record['Dia2_Range']  = safe(r1, 9)
    record['Dia2_MM']     = safe(r1, 10)
    record['Dia2_Count']  = safenum(r1, 11)
    record['Dia2_Qly']    = safe(r1, 12)
    record['Dia2_Wt']     = safenum(r1, 13)
    record['Dia2_Rate']   = safenum(r1, 14)
    record['Dia2_Value']  = safenum(r1, 15)

    # ---- Gemstone / Tanzanite (Row 2) ----
    record['Gem_Label']   = safe(r2, 2)   # e.g., "Tanzanite"
    record['Gem_Size']    = safe(r2, 8)   # e.g., "SOVL 7x5"
    record['Gem_Cut']     = safe(r2, 9)   # e.g., "SO6"
    record['Gem_MM']      = safe(r2, 10)  # e.g., "7.0 X x5"
    record['Gem_Count']   = safenum(r2, 11)
    record['Gem_Qly']     = safe(r2, 12)  # e.g., "OVTZ2"
    record['Gem_Wt']      = safenum(r2, 13)
    record['Gem_Rate']    = safenum(r2, 14)
    record['Gem_Value']   = safenum(r2, 15)

    # ---- Pricing Type (Row 2 col4 / Row 3 col4) ----
    pricing_type_r2 = safe(r2, 4)
    pricing_type_r3 = safe(r3, 4)
    record['Pricing_Type'] = pricing_type_r2 if pricing_type_r2 else pricing_type_r3

    # ---- Set Type & Charges (Row 0, cols 24-33) ----
    record['Set_Type']    = safe(r0, 24)   # e.g., DRPRWF
    record['Set_Stones']  = safenum(r0, 25)
    record['Chg_Stn']     = safenum(r0, 26)
    record['Chg']         = safenum(r0, 27)
    record['Tot_Chg']     = safenum(r0, 28)
    record['CFP']         = safenum(r0, 29)
    record['RH']          = safenum(r0, 30)
    record['Other_Charges'] = safenum(r0, 31)
    record['Plating']     = safenum(r0, 32)
    record['Tag_Packing'] = safenum(r0, 33)

    # ---- Set / Component info (Row 0 col25 or r1 col29) ----
    record['Comp_Info']   = safe(r1, 29)   # e.g., "Comp+LPU"
    record['Components']  = safe(r2, 29)   # e.g., "Components: 2"

    # ---- Pricing Summary (Row 0, cols 41-50) ----
    record['FOB_India']   = safenum(r0, 42)
    record['NY_Cost']     = safenum(r0, 43)
    record['NY_Cost_Landed'] = safenum(r0, 44)
    record['NY_With_CentStn'] = safenum(r0, 46)
    record['NY_Selling']  = safenum(r0, 48)
    record['NY_Sell_Incl_COOP'] = safenum(r0, 49)
    record['Cust_Margin'] = safenum(r0, 50)
    record['Retail']      = safenum(r0, 51)

    # ---- Mark Up / Duty ----
    record['Gold_Factor'] = safenum(r0, 52)
    record['Mark_Up']     = safenum(r0, 53)
    record['Duty']        = safenum(r0, 54)
    record['NY_Mark_Up']  = safenum(r0, 55)
    record['India_Margin'] = safenum(r0, 57)
    record['USA_Margin']  = safenum(r0, 58)

    # ---- Quotation / List ----
    record['Quotation_Sr'] = safenum(r0, 59)
    record['Quotation_Code'] = safe(r0, 60)
    record['List_Version'] = safe(r0, 61)
    record['Quality_Comb'] = safe(r0, 62)

    # ---- Style Physical Details (Row 0, cols 65-72) ----
    record['Style_Qty']   = safenum(r0, 64)
    record['CPU']         = safenum(r0, 65)
    record['Est_Wt']      = safenum(r0, 66)
    record['Length_MM']   = safenum(r0, 67)
    record['Width_MM']    = safenum(r0, 68)
    record['Depth_MM']    = safenum(r0, 69)
    record['Labour_Cost'] = safenum(r0, 70)
    record['Lab_Earned']  = safenum(r0, 71)

    # ---- WaxWt / ConvWt / Flags ----
    record['WaxWt']       = safenum(r4, 5)
    record['ConvWt']      = safenum(r5, 5)
    record['Dia_Approval'] = safe(r3, 16).replace('Dia. Approval:', '').strip()
    record['BackPlate']   = safe(r4, 16).replace('BackPlate:', '').strip()
    record['Exclusive']   = safe(r5, 16).replace('Exclusive:', '').strip()
    record['Tone']        = safe(r3, 29).replace('Tone:', '').strip()
    record['Surface_Finish'] = safe(r4, 29).replace('Surface Finish:', '').strip()
    record['Surface_Area_MM'] = safenum(r5, 30)

    # ---- Metal Loss / Fnd Details (Row 0, cols 75-82) ----
    record['Fnd_Gold_Rate']    = safenum(r0, 75)
    record['Fnd_Silver_Rate']  = safenum(r0, 76)
    record['Fnd_Brass_Rate']   = safenum(r0, 77)
    record['Fnd_Platinum_Rate']= safenum(r0, 78)
    record['Fnd_Steel_Rate']   = safenum(r0, 79)
    record['Metal_Loss_Pct']   = safenum(r0, 80)

    # ---- Remarks (Row 1, col 75) ----
    record['Remarks']     = safe(r1, 75)

    # ---- SKU ----
    record['SKU'] = ''
    for row in block:
        for c_idx, val in enumerate(row):
            if val == 'SKU#' and c_idx + 1 < len(row):
                record['SKU'] = str(row[c_idx + 1]).strip() if row[c_idx + 1] else ''
                break

    return record


def convert_excel(file_path, output_csv, images_dir):
    file_path = os.path.abspath(file_path)
    os.makedirs(images_dir, exist_ok=True)

    print("Starting Excel in background...")
    excel = win32.Dispatch('Excel.Application')
    excel.Visible = False
    excel.DisplayAlerts = False
    wb = None

    try:
        wb = excel.Workbooks.Open(file_path)
        ws = wb.Worksheets(1)
        print("Reading all cells...")
        used_range = ws.UsedRange.Value

        header1 = list(used_range[0])
        header2 = list(used_range[1])

        records = []
        current_block = []
        block_start_row = -1  # 1-based row in sheet for image lookup

        raw_blocks = []
        for i, row in enumerate(used_range):
            if i < 2:
                continue  # skip headers
            row = list(row)
            col_a = row[0] if row else None

            if col_a is not None and isinstance(col_a, (int, float)):
                if current_block:
                    raw_blocks.append((current_block, block_start_row))
                current_block = [row]
                block_start_row = i + 1  # 1-based
            elif current_block:
                current_block.append(row)

        if current_block:
            raw_blocks.append((current_block, block_start_row))

        import re
        import copy

        def get_digits(s):
            m = re.search(r'\d{6}', str(s or ''))
            return m.group(0) if m else str(s or '').strip()

        seen_digits = set()

        # Consolidate Bridal Sets (RE + BS Code + RB -> single BS record) and deduplicate
        idx = 0
        while idx < len(raw_blocks):
            blk, start_row = raw_blocks[idx]
            re_sn = str(blk[0][1] or '').strip() if blk and len(blk[0]) > 1 else ''
            re_dig = get_digits(re_sn)

            bs_codes = [r[4] for r in blk if len(r) > 4 and r[4] and str(r[4]).strip().startswith('BS')]

            if bs_codes and (idx + 1) < len(raw_blocks):
                bs_code = str(bs_codes[0]).strip()
                bs_dig = get_digits(bs_code)

                if bs_dig in seen_digits or re_dig in seen_digits:
                    idx += 2
                    continue

                next_blk, _ = raw_blocks[idx + 1]
                rb_sn = str(next_blk[0][1] or '').strip() if next_blk and len(next_blk[0]) > 1 else ''
                rb_dig = get_digits(rb_sn)

                seen_digits.add(bs_dig)
                seen_digits.add(re_dig)
                seen_digits.add(rb_dig)

                re_rec = parse_block(blk, header1, header2)
                rb_rec = parse_block(next_blk, header1, header2)

                # Look for BOX SET price in next_blk
                box_prices = [r[50] for r in next_blk if len(r) > 50 and r[48] and 'BOX SET' in str(r[48]).upper()]
                box_price = box_prices[0] if box_prices else None

                # Look for TOTAL DIA WT in next_blk
                tot_dia_wts = [r[13] for r in next_blk if len(r) > 13 and r[12] and 'TOTAL DIA WT' in str(r[12]).upper()]
                tot_dia_wt = tot_dia_wts[0] if tot_dia_wts else None

                # Build consolidated Bridal Set record
                bs_rec = copy.deepcopy(re_rec)
                bs_rec['Style_No'] = bs_code
                bs_rec['Components'] = f"Bridal Set ({re_rec.get('Style_No')} + {rb_rec.get('Style_No')})"

                # Sum metal weights
                try:
                    sil_wt = float(re_rec.get('Max_Wt_SIL') or 0) + float(rb_rec.get('Max_Wt_SIL') or 0)
                    bs_rec['Max_Wt_SIL'] = str(round(sil_wt, 2)) if sil_wt > 0 else ''
                except:
                    pass
                try:
                    gold_wt = float(re_rec.get('Max_Wt_GOLD') or 0) + float(rb_rec.get('Max_Wt_GOLD') or 0)
                    bs_rec['Max_Wt_GOLD'] = str(round(gold_wt, 2)) if gold_wt > 0 else ''
                except:
                    pass

                if tot_dia_wt is not None:
                    try:
                        bs_rec['Dia1_Wt'] = str(round(float(tot_dia_wt), 2))
                    except:
                        bs_rec['Dia1_Wt'] = str(tot_dia_wt)

                if box_price is not None:
                    try:
                        bs_rec['NY_Sell_Incl_COOP'] = str(round(float(box_price), 2))
                    except:
                        bs_rec['NY_Sell_Incl_COOP'] = str(box_price)

                records.append((bs_rec, start_row))
                idx += 2  # Consumed both RE and RB together
            else:
                if re_dig in seen_digits:
                    idx += 1
                    continue

                seen_digits.add(re_dig)
                records.append((parse_block(blk, header1, header2), start_row))
                idx += 1

        print(f"\nParsed {len(records)} clean unique consolidated style entries.")

        if records:
            keys = list(records[0][0].keys())
            keys.append('Image_File')

            # --- Extract Images ---
            print("\nExtracting images...")
            # Map: row_number (1-based) -> image filename
            row_to_img = {}
            try:
                shapes = ws.Shapes
                print(f"Found {shapes.Count} shapes.")
                for i in range(1, shapes.Count + 1):
                    shape = shapes(i)
                    if shape.Type == 13:  # msoPicture
                        img_row = shape.TopLeftCell.Row
                        row_to_img[img_row] = img_row  # store row for later naming
            except Exception as e:
                print(f"Warning: Could not iterate shapes: {e}")

            # Now save images and name them by Style No
            try:
                from PIL import ImageGrab
                for i in range(1, ws.Shapes.Count + 1):
                    shape = ws.Shapes(i)
                    if shape.Type == 13:
                        img_row = shape.TopLeftCell.Row
                        # Find which style this belongs to
                        style_no = "Unknown"
                        for rec, start_row in records:
                            if abs(start_row - img_row) <= 10:
                                style_no = rec.get('Style_No', 'Unknown')
                                break

                        safe_name = "".join(c for c in style_no if c.isalnum() or c in (' ', '_', '-')).strip()
                        img_file = f"{safe_name}_R{img_row}.png"
                        img_path = os.path.join(images_dir, img_file)

                        shape.Copy()
                        time.sleep(0.6)
                        img = ImageGrab.grabclipboard()
                        if img:
                            img.save(img_path, 'PNG')
                            print(f"  Saved: {img_file}")
                            row_to_img[img_row] = img_file
                        else:
                            print(f"  Clipboard empty for row {img_row}")
            except ImportError:
                print("Pillow not installed - skipping image save. Run: pip install pillow")

            # Match images to records
            for rec, start_row in records:
                matched = ''
                for img_row, img_file in row_to_img.items():
                    if abs(start_row - img_row) <= 10:
                        matched = img_file if isinstance(img_file, str) else ''
                        break
                rec['Image_File'] = matched

            # Write CSV
            print(f"\nWriting to {output_csv}...")
            with open(output_csv, 'w', newline='', encoding='utf-8') as f:
                writer = csv.DictWriter(f, fieldnames=keys)
                writer.writeheader()
                for rec, _ in records:
                    writer.writerow(rec)
            print("Done!")

    finally:
        if wb: wb.Close(False)
        excel.Quit()


if __name__ == "__main__":
    convert_excel(
        "C:/Users/ASUS/Documents/Manufacturing/3Marketing/Recap/files_BT/Zales Morganite Bridal pricing.xlsx",
        "outputs/Zales_Morganite/structured_styles.csv",
        "outputs/Zales_Morganite/style_images"
    )
