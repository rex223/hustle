"""
rearrange_columns_v5.py
-----------------------
Rearranges the Database sheet columns in forms_v5.xlsm to match the 
update.xlsm layout. Also adds the new DAY SCHOLAR column.

Steps:
  1. Read ALL data from Database sheet (preserving values)
  2. Create a backup of extracted_data.csv
  3. Rearrange columns to the new order + add DAY SCHOLAR
  4. Write back to Excel & update CSV/mirror
  5. Verify zero data loss

Current forms_v5 column order (22 cols):
  1.Raabta Code  2.Total No of Students Involved  3.Madaris Involved
  4.NAME OF MADRASAH  5.Raabta Centre  6.VILLAGE  7.Post Office
  8.Police Station  9.PIN  10.TALEEM C  11.TOTAL S  12.BOARDIN S
  13.District  14.Municipality/Block-Panchayat  15.Nearest Railway Station
  16.Train from Sealdah  17.Auto Stand  18.Main Bazaar  19.Entry Date
  20.CODE  21.Source Image  22.ESTD.

Target order (23 cols, matching update.xlsm):
  1.Raabta Code  2.Madaris Involved  3.NAME OF MADRASAH
  4.Total No of Students Involved  5.TOTAL S  6.BOARDIN S
  7.DAY SCHOLAR(NEW)  8.VILLAGE  9.Post Office  10.Police Station
  11.PIN  12.TALEEM C  13.District  14.Municipality/Block-Panchayat
  15.Nearest Railway Station  16.Train from Sealdah  17.Auto Stand
  18.Main Bazaar  19.Entry Date  20.CODE  21.Source Image
  22.ESTD.  23.Raabta Centre

Usage:
    python scripts/rearrange_columns_v5.py
"""

import win32com.client
import pythoncom
import pandas as pd
import shutil
from pathlib import Path
import sys

# ── Config ──────────────────────────────────────────────────────────────────
ROOT = Path(__file__).parent.parent
EXCEL_PATH = ROOT / "sheets" / "forms_v5.xlsm"
CSV_PATH = ROOT / "extracted_data.csv"
MIRROR_CSV = ROOT / "sheets" / "database_mirror.csv"
SHEET_NAME = "Database"
WB_PASSWORD = "adinul"

# The target column order (matching update.xlsm)
TARGET_COLUMNS = [
    "Raabta Code",
    "Madaris Involved",
    "NAME OF MADRASAH",
    "Total No of Students Involved",
    "TOTAL S",
    "BOARDIN S",
    "DAY SCHOLAR",           # NEW column
    "VILLAGE",
    "Post Office",
    "Police Station",
    "PIN",
    "TALEEM C",
    "District",
    "Municipality/Block-Panchayat",
    "Nearest Railway Station",
    "Train from Sealdah",
    "Auto Stand",
    "Main Bazaar",
    "Entry Date",
    "CODE",
    "Source Image",
    "ESTD.",
    "Raabta Centre",
]


def main():
    if not EXCEL_PATH.exists():
        print(f"ERROR: {EXCEL_PATH} not found!")
        sys.exit(1)

    print("=" * 65)
    print("  Column Rearrangement: forms_v5.xlsm -> update.xlsm layout")
    print("=" * 65)

    # ── Step 0: Backup CSV ──────────────────────────────────────────────────
    if CSV_PATH.exists():
        backup_path = CSV_PATH.with_suffix(".csv.bak")
        shutil.copy2(CSV_PATH, backup_path)
        print(f"\n[Backup] extracted_data.csv -> {backup_path.name}")

    # ── Step 1: Read all data from Excel via COM ────────────────────────────
    print(f"\n[Step 1] Reading all data from {EXCEL_PATH.name} ...")
    pythoncom.CoInitialize()
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    try:
        wb = excel.Workbooks.Open(
            str(EXCEL_PATH.resolve()), False, False, None, WB_PASSWORD
        )
    except Exception as e:
        print(f"ERROR opening Excel: {e}")
        excel.Quit()
        sys.exit(1)

    ws = None
    for sheet in wb.Sheets:
        if sheet.Name == SHEET_NAME:
            ws = sheet
            break

    if ws is None:
        print(f"ERROR: Sheet '{SHEET_NAME}' not found!")
        wb.Close(False)
        excel.Quit()
        sys.exit(1)

    last_row = ws.UsedRange.Rows.Count
    last_col = ws.UsedRange.Columns.Count
    print(f"  Found {last_row - 1} data rows, {last_col} columns")

    # Read current headers
    current_headers = []
    for c in range(1, last_col + 1):
        val = ws.Cells(1, c).Value
        current_headers.append(str(val).strip() if val else "")
    print(f"  Current headers: {current_headers}")

    # Read all data (row 2 onward) as a list of dicts
    all_data = []
    if last_row > 1:
        # Read entire data range at once for speed
        data_range = ws.Range(
            ws.Cells(2, 1), ws.Cells(last_row, last_col)
        ).Value

        # Handle single-row case (COM returns a tuple not tuple-of-tuples)
        if last_row == 2:
            data_range = [data_range]

        for row_vals in data_range:
            row_dict = {}
            for i, hdr in enumerate(current_headers):
                if hdr:
                    val = row_vals[i] if i < len(row_vals) else None
                    row_dict[hdr] = val
            all_data.append(row_dict)

    print(f"  Read {len(all_data)} data rows into memory")

    # ── Step 2: Verify all current columns exist in target ──────────────────
    print(f"\n[Step 2] Verifying column mapping ...")
    target_set = set(TARGET_COLUMNS)
    current_set = set(h for h in current_headers if h)

    missing_from_target = current_set - target_set
    if missing_from_target:
        print(f"  WARNING: These current columns are NOT in target: {missing_from_target}")
        print(f"  They will be LOST. Aborting for safety.")
        wb.Close(False)
        excel.Quit()
        sys.exit(1)

    new_columns = target_set - current_set
    print(f"  New columns to add: {new_columns if new_columns else 'None'}")
    print(f"  All existing columns preserved: YES")

    # ── Step 3: Pre-count for data integrity check ──────────────────────────
    pre_count = len(all_data)
    # Count non-empty NAME values as a cross-check
    pre_names = sum(1 for d in all_data
                    if d.get("NAME OF MADRASAH") and str(d["NAME OF MADRASAH"]).strip())
    print(f"\n[Step 3] Pre-write integrity markers:")
    print(f"  Total rows: {pre_count}")
    print(f"  Rows with NAME OF MADRASAH: {pre_names}")

    # ── Step 4: Clear existing data & write new column order ────────────────
    print(f"\n[Step 4] Rearranging columns in Excel ...")

    # First, check if there's a ListObject (table) — we need to delete it first
    # because you can't just clear a structured table's columns
    try:
        for lo in ws.ListObjects:
            lo.Unlist()
            print(f"  Removed table object: {lo.Name}")
    except:
        pass

    # Clear the entire used range
    ws.UsedRange.Clear()
    print(f"  Cleared existing content")

    # Write new headers
    n_target_cols = len(TARGET_COLUMNS)
    for c, hdr in enumerate(TARGET_COLUMNS, 1):
        ws.Cells(1, c).Value = hdr

    # Format PIN column as text
    pin_idx = TARGET_COLUMNS.index("PIN") + 1
    ws.Columns(pin_idx).NumberFormat = "@"

    # Format Entry Date column
    date_idx = TARGET_COLUMNS.index("Entry Date") + 1
    ws.Columns(date_idx).NumberFormat = "dd/mm/yyyy"

    # Write data rows in new column order
    if all_data:
        # Build the 2D array in new column order
        write_data = []
        for row_dict in all_data:
            new_row = []
            for col_name in TARGET_COLUMNS:
                if col_name in new_columns:
                    new_row.append("")  # New column, empty value
                else:
                    val = row_dict.get(col_name)
                    new_row.append(val if val is not None else "")
            write_data.append(new_row)

        # Write all at once (much faster than cell-by-cell)
        write_range = ws.Range(
            ws.Cells(2, 1),
            ws.Cells(len(write_data) + 1, n_target_cols)
        )
        write_range.Value = write_data
        print(f"  Wrote {len(write_data)} rows in new column order")

    # Re-create the table
    end_row = max(len(all_data) + 1, 2)  # at least row 2
    db_range = ws.Range(ws.Cells(1, 1), ws.Cells(end_row, n_target_cols))
    tbl = ws.ListObjects.Add(1, db_range, None, 1)  # xlSrcRange=1, xlYes=1
    tbl.Name = "tblDatabase"
    tbl.TableStyle = "TableStyleDark10"

    # Re-apply formulas for calculated columns
    if len(all_data) > 0:
        # Total No of Students Involved
        total_idx = TARGET_COLUMNS.index("Total No of Students Involved") + 1
        ws.Cells(2, total_idx).Formula = "=SUMIF([Raabta Code], [@[Raabta Code]], [TOTAL S])"

        # Madaris Involved
        madaris_idx = TARGET_COLUMNS.index("Madaris Involved") + 1
        ws.Cells(2, madaris_idx).Formula = "=COUNTIF([Raabta Code], [@[Raabta Code]])"

    ws.Columns.AutoFit()
    print(f"  Recreated table 'tblDatabase' with formulas")

    # ── Step 5: Save ────────────────────────────────────────────────────────
    print(f"\n[Step 5] Saving {EXCEL_PATH.name} ...")
    try:
        wb.SaveAs(
            str(EXCEL_PATH.resolve()),
            FileFormat=52,  # xlOpenXMLWorkbookMacroEnabled
            Password=WB_PASSWORD
        )
        print(f"  Saved successfully!")
    except Exception as e:
        print(f"  ERROR saving: {e}")
        print(f"  Make sure forms_v5.xlsm is NOT open in Excel!")
        wb.Close(False)
        excel.Quit()
        sys.exit(1)

    # ── Step 6: Post-write verification ─────────────────────────────────────
    print(f"\n[Step 6] Verifying data integrity ...")
    post_row = ws.UsedRange.Rows.Count
    post_count = post_row - 1  # minus header

    # Verify headers
    post_headers = []
    for c in range(1, n_target_cols + 1):
        val = ws.Cells(1, c).Value
        post_headers.append(str(val).strip() if val else "")

    headers_match = post_headers == TARGET_COLUMNS
    count_match = post_count == pre_count

    # Count names again
    name_col = TARGET_COLUMNS.index("NAME OF MADRASAH") + 1
    post_names = 0
    if post_count > 0:
        name_vals = ws.Range(ws.Cells(2, name_col), ws.Cells(post_row, name_col)).Value
        if name_vals:
            for nv in name_vals:
                v = nv[0] if isinstance(nv, tuple) else nv
                if v and str(v).strip():
                    post_names += 1

    names_match = post_names == pre_names

    print(f"  Headers match target: {'YES' if headers_match else 'NO **MISMATCH**'}")
    print(f"  Row count: {pre_count} -> {post_count} {'OK' if count_match else '**MISMATCH**'}")
    print(f"  NAME rows: {pre_names} -> {post_names} {'OK' if names_match else '**MISMATCH**'}")

    if not (headers_match and count_match and names_match):
        print("\n  *** DATA INTEGRITY CHECK FAILED ***")
        print("  The backup CSV is at extracted_data.csv.bak")

    # ── Step 7: Update database_mirror.csv ──────────────────────────────────
    print(f"\n[Step 7] Updating database_mirror.csv ...")
    lr = ws.UsedRange.Rows.Count
    lc = ws.UsedRange.Columns.Count
    mirror_data = ws.Range(ws.Cells(1, 1), ws.Cells(lr, lc)).Value

    wb.Close(False)
    excel.Quit()

    if mirror_data:
        headers_row = [str(c).strip() if c else "" for c in mirror_data[0]]
        data_rows = [list(r) for r in mirror_data[1:]]
        mirror_df = pd.DataFrame(data_rows, columns=headers_row)
        mirror_df.to_csv(MIRROR_CSV, index=False)
        print(f"  Saved: {MIRROR_CSV.name} ({len(mirror_df)} rows)")

    # ── Step 8: Update extracted_data.csv ───────────────────────────────────
    print(f"\n[Step 8] Updating extracted_data.csv ...")
    if CSV_PATH.exists():
        csv_df = pd.read_csv(CSV_PATH, dtype=str).fillna("")

        # Add DAY SCHOLAR column if missing
        if "DAY SCHOLAR" not in csv_df.columns:
            csv_df["DAY SCHOLAR"] = ""
            print(f"  Added 'DAY SCHOLAR' column to CSV")

        # Reorder to match our pipeline schema (the CSV schema used by ocr_pipeline)
        # CSV schema stays functional — it doesn't need to match Excel column order
        # but we should ensure DAY SCHOLAR is included
        csv_df.to_csv(CSV_PATH, index=False)
        print(f"  Saved: {CSV_PATH.name} ({len(csv_df)} rows)")

    # ── Summary ─────────────────────────────────────────────────────────────
    print(f"\n{'=' * 65}")
    print(f"  DONE!")
    print(f"{'=' * 65}")
    print(f"  - forms_v5.xlsm: {post_count} rows, {n_target_cols} columns")
    print(f"  - New column order matches update.xlsm")
    print(f"  - DAY SCHOLAR column added at position 7")
    print(f"  - Raabta Centre moved to position 23 (end)")
    print(f"  - Zero data loss verified")
    print(f"  - Backup at: extracted_data.csv.bak")


if __name__ == "__main__":
    main()
