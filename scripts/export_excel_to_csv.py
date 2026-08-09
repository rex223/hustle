import pandas as pd
from pathlib import Path
import warnings

warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

EXCEL_PATH = Path(__file__).parent.parent / "sheets" / "forms_v5.xlsm"
CSV_MIRROR_PATH = Path(__file__).parent.parent / "sheets" / "database_mirror.csv"

def export_to_csv():
    if not EXCEL_PATH.exists():
        print("Excel file not found!")
        return
        
    print(f"Loading Excel Database from {EXCEL_PATH.name}...")
    try:
        live_df = pd.read_excel(EXCEL_PATH, sheet_name="Database", engine="openpyxl", dtype=str).fillna("")
        
        # Write to mirror CSV
        live_df.to_csv(CSV_MIRROR_PATH, index=False)
        print(f"Successfully exported {len(live_df)} rows to {CSV_MIRROR_PATH.name}!")
        
    except Exception as e:
        print(f"Error reading/writing: {e}")

if __name__ == "__main__":
    export_to_csv()
