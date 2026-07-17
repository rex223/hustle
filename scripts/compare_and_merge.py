"""
compare_and_merge.py
─────────────────────
Compares 'previous' sheet (sheets/previous.xlsm or previous.xlsx) 
against current forms_v4.xlsm and injects ONLY the missing rows.
Unique key: (NAME OF MADRASAH, PIN)
"""

import sys
import win32com.client
import pythoncom
import pandas as pd
from pathlib import Path

BASE_DIR    = Path(r"D:\CLLG\hustle")
PREV_PATH   = BASE_DIR / "sheets" / "previous.xlsm"   # Change extension if .xlsx
CURR_PATH   = BASE_DIR / "sheets" / "forms_v4.xlsm"
MIRROR_CSV  = BASE_DIR / "sheets" / "database_mirror.csv"
WB_PASSWORD = "adinul"
SHEET_NAME  = "Database"
NAME_COL    = "NAME OF MADRASAH"
PIN_COL     = "PIN"

def normalize_key(name, pin):
    return (str(name).strip().upper(), str(pin).strip())

def read_sheet_via_com(path, password=None, sheet_name=SHEET_NAME):
    """Read a sheet into a list of dicts using win32com."""
    pythoncom.CoInitialize()
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    wb = excel.Workbooks.Open(str(path.resolve()), False, True, None, password or "")
    
    ws = None
    for sheet in wb.Sheets:
        if sheet.Name == sheet_name:
            ws = sheet
            break
    
    if ws is None:
        # Try first sheet
        ws = wb.Sheets(1)
        print(f"  Sheet '{sheet_name}' not found, using: {ws.Name}")
    
    lr = ws.UsedRange.Rows.Count
    lc = ws.UsedRange.Columns.Count
    all_data = ws.Range(ws.Cells(1, 1), ws.Cells(lr, lc)).Value
    wb.Close(False)
    excel.Quit()
    
    if not all_data:
        return [], []
    headers = [str(c).strip() if c else "" for c in all_data[0]]
    rows = [list(r) for r in all_data[1:]]
    return headers, rows

def main():
    # --- Check files ---
    if not CURR_PATH.exists():
        print(f"ERROR: forms_v4.xlsm not found at {CURR_PATH}")
        sys.exit(1)
    
    # Try both .xlsm and .xlsx for previous
    prev_path = PREV_PATH
    if not prev_path.exists():
        prev_path = BASE_DIR / "sheets" / "previous.xlsx"
    if not prev_path.exists():
        print(f"ERROR: Previous file not found. Place it at:")
        print(f"  {PREV_PATH}  (or as previous.xlsx)")
        sys.exit(1)
    
    print(f"Reading previous sheet: {prev_path.name}...")
    prev_headers, prev_rows = read_sheet_via_com(prev_path, WB_PASSWORD)
    print(f"  Rows in previous: {len(prev_rows)}")
    
    # Build prev DataFrame
    prev_df = pd.DataFrame(prev_rows, columns=prev_headers)
    prev_df = prev_df.fillna("")
    
    # Check if required columns exist
    if NAME_COL not in prev_headers:
        # Try to find a close match
        print(f"  Headers found: {prev_headers[:10]}")
        print(f"ERROR: '{NAME_COL}' column not found in previous sheet.")
        sys.exit(1)
    if PIN_COL not in prev_headers:
        print(f"ERROR: '{PIN_COL}' column not found in previous sheet.")
        sys.exit(1)
    
    # --- Read current Excel ---
    print(f"\nReading current forms_v4.xlsm...")
    curr_headers, curr_rows = read_sheet_via_com(CURR_PATH, WB_PASSWORD)
    print(f"  Rows in current : {len(curr_rows)}")
    
    # Build existing keys
    existing_keys = set()
    for row in curr_rows:
        row_dict = dict(zip(curr_headers, row))
        n = str(row_dict.get(NAME_COL, "") or "")
        p = str(row_dict.get(PIN_COL, "") or "")
        if n or p:
            existing_keys.add(normalize_key(n, p))
    
    # --- Find missing rows from previous ---
    missing_rows = []
    seen_in_batch = set()
    for _, row in prev_df.iterrows():
        n = str(row.get(NAME_COL, "") or "")
        p = str(row.get(PIN_COL, "") or "")
        key = normalize_key(n, p)
        if key not in existing_keys and key not in seen_in_batch and (n or p):
            missing_rows.append(row)
            seen_in_batch.add(key)
    
    print(f"\n-- Comparison Result --")
    print(f"  Previous sheet rows : {len(prev_rows)}")
    print(f"  Current Excel rows  : {len(curr_rows)}")
    print(f"  Missing from current: {len(missing_rows)}")
    
    if not missing_rows:
        print("\nOK: All rows from previous sheet are already in forms_v4.xlsm!")
        return
    
    print(f"\n  Sample of missing rows:")
    for r in missing_rows[:10]:
        print(f"    {r.get(NAME_COL,'')} | PIN: {r.get(PIN_COL,'')} | District: {r.get('District','')}")
    
    print(f"\nInjecting {len(missing_rows)} missing rows into forms_v4.xlsm...")
    
    # --- Open current for writing ---
    pythoncom.CoInitialize()
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    wb = excel.Workbooks.Open(str(CURR_PATH.resolve()), False, False, None, WB_PASSWORD)
    ws = None
    for sheet in wb.Sheets:
        if sheet.Name == SHEET_NAME:
            ws = sheet
            break
    if ws is None:
        ws = wb.Sheets(1)
    
    next_row = ws.UsedRange.Rows.Count + 1
    
    for i, row_data in enumerate(missing_rows):
        for col_name in prev_headers:
            if col_name and col_name in curr_headers:
                col_idx = curr_headers.index(col_name) + 1
                val = row_data.get(col_name, "")
                ws.Cells(next_row, col_idx).Value = val if str(val).strip() != "" else None
        next_row += 1
        if (i + 1) % 25 == 0:
            print(f"  ... {i+1}/{len(missing_rows)} written")
    
    print(f"  Saving...")
    try:
        wb.SaveAs(str(CURR_PATH.resolve()), FileFormat=52, Password=WB_PASSWORD)
        wb.Close(False)
        excel.Quit()
        print(f"\nSUCCESS! Injected {len(missing_rows)} rows from previous sheet into forms_v4.xlsm")
    except Exception as e:
        print(f"ERROR saving: {e}")
        print("  Make sure forms_v4.xlsm is NOT open in Excel!")
        wb.Close(False)
        excel.Quit()
        sys.exit(1)
    
    # --- Update mirror CSV ---
    print("\nUpdating database_mirror.csv...")
    curr_df = pd.DataFrame(curr_rows, columns=curr_headers)
    miss_df = pd.DataFrame(missing_rows).reindex(columns=curr_headers)
    combined = pd.concat([curr_df, miss_df], ignore_index=True)
    MIRROR_CSV.parent.mkdir(parents=True, exist_ok=True)
    combined.to_csv(MIRROR_CSV, index=False)
    print(f"Mirror CSV: {len(combined)} total rows -> {MIRROR_CSV.name}")

if __name__ == "__main__":
    main()
