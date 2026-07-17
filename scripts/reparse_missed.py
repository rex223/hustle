"""
reparse_missed.py
─────────────────
Identifies images that produced genuinely missed/errored rows in error_tracker.csv,
copies them back to images/ folder and triggers OCR re-parsing.

Skips rows that were correctly flagged as duplicates of pre-existing data.
Only re-processes:
  1. Rows rejected as "Duplicate within this image scan" (false positive bug)
  2. Rows rejected for "Invalid PIN format"
"""

import pandas as pd
import shutil
from pathlib import Path

BASE_DIR    = Path(__file__).parent.parent
ERROR_CSV   = BASE_DIR / "error_tracker.csv"
DONE_DIR    = BASE_DIR / "images" / "done"
IMAGES_DIR  = BASE_DIR / "images"

# Error reasons that indicate genuinely missed data (not correctly rejected duplicates)
REPARSE_REASONS = [
    "Invalid PIN format",
    "Duplicate Madrasah (Name within this image scan)",  # false-positive bug
]

def main():
    if not ERROR_CSV.exists():
        print("No error_tracker.csv found.")
        return

    df = pd.read_csv(ERROR_CSV, dtype=str).fillna("")
    print(f"Total error rows: {len(df)}")

    # Filter to only the rows that need re-parsing
    needs_reparse = df[df["Error Reason"].str.startswith(tuple(REPARSE_REASONS))]
    print(f"Rows needing re-parse: {len(needs_reparse)}")

    if needs_reparse.empty:
        print("Nothing to re-parse.")
        return

    # Get unique source images
    images_to_reprocess = needs_reparse["Source Image"].unique()
    print(f"\nImages to move back for re-processing:")

    moved = []
    for img_name in images_to_reprocess:
        src = DONE_DIR / img_name
        dst = IMAGES_DIR / img_name
        if src.exists():
            shutil.copy2(str(src), str(dst))
            print(f"  Copied: {img_name}")
            moved.append(img_name)
        elif dst.exists():
            print(f"  Already in images/: {img_name}")
            moved.append(img_name)
        else:
            print(f"  NOT FOUND in done/ or images/: {img_name}")

    print(f"\n{len(moved)} image(s) ready in images/ folder.")
    print(f"\nRows that will be re-extracted:")
    for _, row in needs_reparse.iterrows():
        print(f"  [{row['Source Image']}] {row['NAME OF MADRASAH']} | PIN: {row['PIN']} | Reason: {row['Error Reason']}")

    print(f"\nRun: .\\run_phase1_ocr.bat   to re-parse these images.")
    print(f"Then run: python scripts\\inject_csv_to_excel.py   to inject new rows.")

if __name__ == "__main__":
    main()
