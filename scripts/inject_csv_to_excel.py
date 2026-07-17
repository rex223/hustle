"""
inject_csv_to_excel.py
---------------------
Safely appends ONLY NEW rows from extracted_data.csv into forms_v4.xlsm.
Uses win32com (Excel itself) to handle the password-protected .xlsm file.
Duplicate detection: unique key = (NAME OF MADRASAH, PIN).
Never overwrites or deletes existing rows.
Also checks error_tracker.csv to flag any rows that were errored but 
are now present in extracted_data.csv (meaning they should be injected).
"""

import pandas as pd
from pathlib import Path
import win32com.client
import pythoncom
import sys

# -- Config --------------------------------------------------------------------
BASE_DIR    = Path(__file__).parent.parent
EXCEL_PATH  = BASE_DIR / "sheets" / "forms_v4.xlsm"
CSV_PATH    = BASE_DIR / "extracted_data.csv"
ERROR_CSV   = BASE_DIR / "error_tracker.csv"
MIRROR_CSV  = BASE_DIR / "sheets" / "database_mirror.csv"
SHEET_NAME  = "Database"
WB_PASSWORD = "adinul"
NAME_COL    = "NAME OF MADRASAH"
PIN_COL     = "PIN"

def normalize_key(code, name, pin):
    return (str(code).strip().upper(), str(name).strip().upper(), str(pin).strip())


def main():
    # -- 1. Validate inputs ----------------------------------------------------
    if not EXCEL_PATH.exists():
        print(f"ERROR: Excel not found at {EXCEL_PATH}")
        sys.exit(1)
    if not CSV_PATH.exists():
        print(f"ERROR: CSV not found at {CSV_PATH}")
        sys.exit(1)

    # -- 2. Load extracted_data.csv --------------------------------------------
    print("Loading extracted_data.csv...")
    csv_df = pd.read_csv(CSV_PATH, dtype=str).fillna("")
    print(f"  Total rows in CSV: {len(csv_df)}")

    # -- 3. Check error_tracker.csv --------------------------------------------
    if ERROR_CSV.exists():
        err_df = pd.read_csv(ERROR_CSV, dtype=str).fillna("")
        if NAME_COL in err_df.columns and PIN_COL in err_df.columns and "CODE" in err_df.columns:
            err_keys = set(normalize_key(r["CODE"], r[NAME_COL], r[PIN_COL]) for _, r in err_df.iterrows())
            csv_keys_set = set(normalize_key(r["CODE"], r[NAME_COL], r[PIN_COL]) for _, r in csv_df.iterrows())
            # Rows in error_tracker that are NOW in extracted_data.csv (safely recovered)
            recovered = err_keys & csv_keys_set
            still_errored = err_keys - csv_keys_set
            print(f"\n-- Error Tracker Check ---------------------------------")
            print(f"  Total error rows      : {len(err_df)}")
            print(f"  Recovered in CSV      : {len(recovered)}  (OK - will be injected if new)")
            print(f"  Still missing from CSV: {len(still_errored)}  (WARN - needs re-processing)")
            if still_errored:
                print("  Missing entries:")
                for k in list(still_errored)[:10]:
                    print(f"    - {k[0]} | PIN: {k[1]}")
        print()

    # -- 4. Open password-protected Excel via win32com -------------------------
    print(f"Opening {EXCEL_PATH.name} via Excel (password protected)...")
    pythoncom.CoInitialize()

    wb = None
    excel = None
    target_path = str(EXCEL_PATH.resolve()).lower()

    # Strategy 1: Try to attach to an already-running Excel instance
    # This avoids the "read-only" lock when the user has the file open in Excel
    try:
        excel = win32com.client.GetActiveObject("Excel.Application")
        for existing_wb in excel.Workbooks:
            if existing_wb.FullName.lower() == target_path:
                wb = existing_wb
                print("  Attached to already-open Excel instance.")
                break
    except Exception:
        excel = None

    # Strategy 2: Open a fresh Excel instance
    if wb is None:
        try:
            excel = win32com.client.Dispatch("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            wb = excel.Workbooks.Open(
                str(EXCEL_PATH.resolve()),
                False,       # UpdateLinks
                False,       # ReadOnly = False
                None,        # Format
                WB_PASSWORD  # Password
            )
            print("  Opened in a new Excel instance.")
        except Exception as e:
            print(f"ERROR: Could not open Excel file: {e}")
            if excel:
                excel.Quit()
            sys.exit(1)

    # Find the Database sheet
    ws = None
    for sheet in wb.Sheets:
        if sheet.Name == SHEET_NAME:
            ws = sheet
            break

    if ws is None:
        print(f"ERROR: Sheet '{SHEET_NAME}' not found.")
        print(f"  Available sheets: {[s.Name for s in wb.Sheets]}")
        wb.Close(False)
        excel.Quit()
        sys.exit(1)

    # -- 5. Read existing headers & keys from Excel ----------------------------
    print("Reading existing data from Excel...")
    last_col = ws.UsedRange.Columns.Count
    last_row = ws.UsedRange.Rows.Count

    # Read header row
    excel_headers = []
    for c in range(1, last_col + 1):
        val = ws.Cells(1, c).Value
        excel_headers.append(str(val).strip() if val else "")

    # Safety check
    if NAME_COL not in excel_headers or PIN_COL not in excel_headers:
        print(f"ERROR: Could not find '{NAME_COL}' or '{PIN_COL}' in Excel headers.")
        print(f"  Found headers: {excel_headers}")
        wb.Close(False)
        excel.Quit()
        sys.exit(1)

    name_col_idx = excel_headers.index(NAME_COL) + 1   # 1-based
    pin_col_idx  = excel_headers.index(PIN_COL) + 1
    code_col_idx = excel_headers.index("CODE") + 1

    # Read ALL existing keys quickly using column read
    existing_keys = set()
    if last_row > 1:
        name_vals = ws.Range(ws.Cells(2, name_col_idx), ws.Cells(last_row, name_col_idx)).Value
        pin_vals  = ws.Range(ws.Cells(2, pin_col_idx),  ws.Cells(last_row, pin_col_idx)).Value
        code_vals = ws.Range(ws.Cells(2, code_col_idx), ws.Cells(last_row, code_col_idx)).Value

        # win32com returns a tuple of tuples for ranges
        if name_vals and pin_vals and code_vals:
            for c_row, n_row, p_row in zip(code_vals, name_vals, pin_vals):
                code_val = c_row[0] if isinstance(c_row, tuple) else c_row
                name_val = n_row[0] if isinstance(n_row, tuple) else n_row
                pin_val  = p_row[0] if isinstance(p_row, tuple) else p_row
                if name_val or pin_val or code_val:
                    existing_keys.add(normalize_key(str(code_val or ""), str(name_val or ""), str(pin_val or "")))

    print(f"  Excel has {len(existing_keys)} existing records (row count: {last_row - 1}).")

    # -- 6. Determine NEW rows only --------------------------------------------
    new_rows = []
    seen_in_batch = set()
    for _, row in csv_df.iterrows():
        key = normalize_key(row["CODE"], row[NAME_COL], row[PIN_COL])
        if key not in existing_keys and key not in seen_in_batch:
            new_rows.append(row)
            seen_in_batch.add(key)

    if not new_rows:
        print("\nOK No new rows to inject. Excel is already fully up to date!")
        wb.Close(False)
        excel.Quit()
        return

    print(f"\n  Found {len(new_rows)} NEW row(s) to inject. Starting...")

    # -- 7. Append new rows ----------------------------------------------------
    next_row = last_row + 1
    for i, row_data in enumerate(new_rows):
        for col_name in csv_df.columns:
            if col_name in excel_headers and col_name != "Source Image":
                col_idx = excel_headers.index(col_name) + 1
                val = row_data[col_name]
                ws.Cells(next_row, col_idx).Value = val if val != "" else None
        next_row += 1
        if (i + 1) % 25 == 0:
            print(f"    ... {i+1}/{len(new_rows)} rows written")

    print(f"  All {len(new_rows)} rows written. Saving...")

    # -- 8. Save workbook (keep password) --------------------------------------
    try:
        wb.SaveAs(
            str(EXCEL_PATH.resolve()),
            FileFormat=52,          # xlOpenXMLWorkbookMacroEnabled
            Password=WB_PASSWORD
        )
        wb.Close(False)
        excel.Quit()
        print(f"\nOK SUCCESS! {EXCEL_PATH.name} saved with {len(new_rows)} new rows injected.")
    except Exception as e:
        print(f"\nERROR saving: {e}")
        print("  Make sure forms_v4.xlsm is NOT open in Excel right now!")
        wb.Close(False)
        excel.Quit()
        sys.exit(1)

    # -- 9. Update mirror CSV --------------------------------------------------
    print("\nUpdating database_mirror.csv...")
    try:
        pythoncom.CoInitialize()
        excel2 = win32com.client.Dispatch("Excel.Application")
        excel2.Visible = False
        excel2.DisplayAlerts = False
        wb2 = excel2.Workbooks.Open(
            str(EXCEL_PATH.resolve()), False, True, None, WB_PASSWORD
        )
        ws2 = None
        for sheet in wb2.Sheets:
            if sheet.Name == SHEET_NAME:
                ws2 = sheet
                break

        if ws2:
            lr = ws2.UsedRange.Rows.Count
            lc = ws2.UsedRange.Columns.Count
            all_data = ws2.Range(ws2.Cells(1, 1), ws2.Cells(lr, lc)).Value
            wb2.Close(False)
            excel2.Quit()

            if all_data:
                headers_row = [c if c else "" for c in all_data[0]]
                data_rows   = [list(r) for r in all_data[1:]]
                mirror_df   = pd.DataFrame(data_rows, columns=headers_row)
                MIRROR_CSV.parent.mkdir(parents=True, exist_ok=True)
                mirror_df.to_csv(MIRROR_CSV, index=False)
                print(f"OK Mirror CSV saved: {MIRROR_CSV.name} ({len(mirror_df)} total rows)")
        else:
            wb2.Close(False)
            excel2.Quit()
    except Exception as e:
        print(f"  (Mirror CSV update failed: {e} — not critical)")


if __name__ == "__main__":
    main()
