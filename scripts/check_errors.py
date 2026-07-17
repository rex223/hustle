import pandas as pd
from pathlib import Path
import warnings

# Suppress openpyxl warnings
warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

EXCEL_PATH = Path(__file__).parent.parent / "sheets" / "forms_v4.xlsm"
ERROR_CSV_PATH = Path(__file__).parent.parent / "error_tracker.csv"

def check_errors():
    if not EXCEL_PATH.exists():
        print("Excel file not found!")
        return
        
    if not ERROR_CSV_PATH.exists():
        print("No error_tracker.csv found!")
        return
        
    print("Loading Excel Database...")
    try:
        live_df = pd.read_excel(EXCEL_PATH, sheet_name="Database", engine="openpyxl", dtype=str).fillna("")
        live_keys = set(zip(live_df['NAME OF MADRASAH'].str.strip().str.upper(), live_df['PIN'].str.strip()))
    except Exception as e:
        print(f"Error reading Excel: {e}")
        return

    print("Loading error tracker...")
    try:
        err_df = pd.read_csv(ERROR_CSV_PATH, dtype=str).fillna("")
    except Exception as e:
        print(f"Error reading error tracker: {e}")
        return
        
    if err_df.empty:
        print("Error tracker is empty. Nothing to check.")
        return

    overlap_count = 0
    print("\n--- Cross-check Results ---")
    for idx, row in err_df.iterrows():
        name = str(row.get('NAME OF MADRASAH', '')).strip().upper()
        pin = str(row.get('PIN', '')).strip()
        if (name, pin) in live_keys:
            overlap_count += 1
            print(f"Match found in Excel: {name} (PIN: {pin}) - Originally Rejected Reason: {row.get('Error Reason', '')}")
            
    if overlap_count == 0:
        print("Great! None of the rows in the error tracker are currently in the Excel Database.")
    else:
        print(f"\nTotal matches found: {overlap_count}")

if __name__ == "__main__":
    check_errors()
