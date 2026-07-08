"""
ocr_pipeline.py
---------------
Reads all images from the images/ folder, sends each to Gemini Vision API
to extract madrasa table data, maps it to our DB schema, and merges new rows
into extracted_data.csv. Processed images are moved to images/done/.

Usage:
    python scripts/ocr_pipeline.py
"""

import os
import base64
import json
import shutil
import re
import pandas as pd
import requests
from pathlib import Path
from dotenv import load_dotenv

# ── Config ──────────────────────────────────────────────────────────────────
load_dotenv(dotenv_path=Path(__file__).parent.parent / ".env")
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY", "")
IMAGES_DIR = Path(__file__).parent.parent / "images"
DONE_DIR = IMAGES_DIR / "done"
CSV_PATH = Path(__file__).parent.parent / "extracted_data.csv"
ERROR_CSV_PATH = Path(__file__).parent.parent / "error_tracker.csv"

# Model options:
#   gemini-1.5-flash  — fast, very accurate, established huge daily quotas
#   gemini-2.5-pro    — slower, most accurate (use for tricky/blurry images)
GEMINI_MODEL = "gemini-2.5-flash"

GEMINI_URL = f"https://generativelanguage.googleapis.com/v1beta/models/{GEMINI_MODEL}:generateContent?key={GEMINI_API_KEY}"

# The canonical column order we always write to
SCHEMA = [
    'Entry Date', 'CODE', 'NAME OF MADRASAH', 'Raabta Centre',
    'ESTD.', 'VILLAGE', 'Post Office', 'Police Station', 'PIN',
    'TALEEM C', 'TOTAL S', 'BOARDIN S', 'District',
    'Municipality/Block-Panchayat',
    'Nearest Railway Station', 'Train from Sealdah', 'Auto Stand', 'Main Bazaar',
    'Source Image'
]

MANDATORY_FIELDS = ['CODE', 'NAME OF MADRASAH', 'PIN', 'District', 'VILLAGE']

# Maps common OCR variations -> canonical column name
COLUMN_ALIASES = {
    'code': 'CODE',
    'name of madrasah': 'NAME OF MADRASAH',
    'name': 'NAME OF MADRASAH',
    'madrasah name': 'NAME OF MADRASAH',
    'estd': 'ESTD.',
    'estd.': 'ESTD.',
    'established': 'ESTD.',
    'village': 'VILLAGE',
    'post office': 'Post Office',
    'po': 'Post Office',
    'p.o.': 'Post Office',
    'police station': 'Police Station',
    'ps': 'Police Station',
    'p.s.': 'Police Station',
    'pin': 'PIN',
    'pincode': 'PIN',
    'pin code': 'PIN',
    'taleem c': 'TALEEM C',
    'taleem': 'TALEEM C',
    'total s': 'TOTAL S',
    'total students': 'TOTAL S',
    'boardin s': 'BOARDIN S',
    'boarding': 'BOARDIN S',
    'boarding students': 'BOARDIN S',
    'district': 'District',
    'raabta centre': 'Raabta Centre',
    'raabta center': 'Raabta Centre',
}

PROMPT = """
You are a data extraction assistant. The image contains a table of madrasa (Islamic school) records.
Extract ALL rows from the table as a JSON array. Each element should be a flat object with these exact keys
where data is available (use empty string "" if a cell is empty or not visible):

"CODE", "NAME OF MADRASAH", "ESTD.", "VILLAGE", "Post Office", "Police Station",
"PIN", "TALEEM C", "TOTAL S", "BOARDIN S", "District"

Rules:
- If Post Office and Police Station are in one cell separated by "-", split them correctly.
- Do NOT include header rows in the output.
- If the table spans multiple lines for one record, combine them into one object.
- Return ONLY valid JSON array, nothing else, no markdown, no explanation.

Example output:
[
  {"CODE": "S24PGS 453", "NAME OF MADRASAH": "DARUL ULOOM ASRARIA", "ESTD.": "2010",
   "VILLAGE": "FAKIRPARA", "Post Office": "SANTOSH PUR", "Police Station": "MAHESHTALA",
   "PIN": "700142", "TALEEM C": "HIFZ", "TOTAL S": "389", "BOARDIN S": "345", "District": "South 24 Parganas"}
]
"""


def image_to_base64(img_path: Path) -> tuple[str, str]:
    """Returns (base64_data, mime_type)"""
    ext = img_path.suffix.lower()
    mime = {'.jpg': 'image/jpeg', '.jpeg': 'image/jpeg',
            '.png': 'image/png', '.webp': 'image/webp'}.get(ext, 'image/jpeg')
    with open(img_path, 'rb') as f:
        return base64.b64encode(f.read()).decode(), mime


def call_gemini(img_path: Path) -> list[dict]:
    """Send image to Gemini Vision and return extracted rows."""
    b64, mime = image_to_base64(img_path)
    payload = {
        "contents": [{
            "parts": [
                {"text": PROMPT},
                {"inline_data": {"mime_type": mime, "data": b64}}
            ]
        }],
        "generationConfig": {"temperature": 0, "maxOutputTokens": 8192}
    }
    import time
    for attempt in range(5):
        resp = requests.post(GEMINI_URL, json=payload, timeout=60)
        if resp.status_code == 429:
            wait = 15 * (attempt + 1)
            print(f"    [429 Rate Limit hit!] Gemini is rate-limiting us. Sleeping for {wait}s before retrying...")
            time.sleep(wait)
            continue
        resp.raise_for_status()
        break
    else:
        resp.raise_for_status()
    text = resp.json()['candidates'][0]['content']['parts'][0]['text'].strip()

    # Strip markdown fences if Gemini wraps in ```json
    text = re.sub(r'^```[a-z]*\n?', '', text, flags=re.MULTILINE)
    text = re.sub(r'```$', '', text, flags=re.MULTILINE).strip()

    try:
        return json.loads(text)
    except json.JSONDecodeError:
        import ast
        # Replace common JSON-only tokens with Python equivalents before eval
        text_fixed = text.replace('null', 'None').replace('true', 'True').replace('false', 'False')
        return ast.literal_eval(text_fixed)


def map_to_schema(raw_rows: list[dict], img_name: str) -> pd.DataFrame:
    """Map raw extracted dicts to canonical schema columns."""
    mapped = []
    for row in raw_rows:
        new_row = {col: "" for col in SCHEMA}
        for raw_key, val in row.items():
            canonical = COLUMN_ALIASES.get(raw_key.strip().lower())
            if canonical:
                new_row[canonical] = str(val).strip() if val else ""
        new_row['Source Image'] = img_name
        mapped.append(new_row)
    return pd.DataFrame(mapped, columns=SCHEMA)


def get_existing_keys() -> set:
    if CSV_PATH.exists():
        existing = pd.read_csv(CSV_PATH, dtype=str).fillna("")
        return set(zip(existing['NAME OF MADRASAH'].str.strip().str.upper(), existing['PIN'].str.strip()))
    return set()


def validate_and_split(new_df: pd.DataFrame, existing_keys: set) -> tuple[pd.DataFrame, pd.DataFrame]:
    valid_rows = []
    error_rows = []
    seen_in_batch = set()
    
    import datetime
    today_str = datetime.date.today().strftime('%d/%m/%Y')
    
    for idx, row in new_df.iterrows():
        errors = []
        code = str(row.get('CODE', '')).strip().upper()
        name = str(row.get('NAME OF MADRASAH', '')).strip().upper()
        pin = str(row.get('PIN', '')).strip()
        
        # Always set Entry Date (even for rejected rows)
        if not str(row.get('Entry Date', '')).strip():
            row['Entry Date'] = today_str
        
        # 1. Missing Mandatory Fields
        missing = [f for f in MANDATORY_FIELDS if not str(row.get(f, '')).strip()]
        if missing:
            errors.append(f"Missing mandatory: {', '.join(missing)}")
            
        # 2. Duplicate Check (NAME + PIN)
        key = (name, pin)
        if name and key in existing_keys:
            errors.append("Duplicate Madrasah (Name + PIN already in database)")
        elif name and key in seen_in_batch:
            errors.append("Duplicate Madrasah (Name + PIN within this image scan)")
        else:
            if name: seen_in_batch.add(key)
            
        # 3. Illegal Format: PIN
        if pin and (not pin.isdigit() or len(pin) != 6):
            errors.append(f"Invalid PIN format: '{pin}' (must be 6 digits)")
            
        # 4. Truncated Row check (>50% empty fields)
        empty_count = sum(1 for col in SCHEMA if not str(row.get(col, '')).strip())
        if empty_count > len(SCHEMA) / 2:
            errors.append("Row looks truncated (>50% empty fields)")
            
        if errors:
            err_row = row.copy()
            err_row['Error Reason'] = " | ".join(errors)
            error_rows.append(err_row)
        else:
            valid_rows.append(row)
            
    df_valid = pd.DataFrame(valid_rows, columns=SCHEMA) if valid_rows else pd.DataFrame(columns=SCHEMA)
    err_cols = SCHEMA + ['Error Reason']
    df_err = pd.DataFrame(error_rows, columns=err_cols) if error_rows else pd.DataFrame(columns=err_cols)
    
    return df_valid, df_err


def merge_valid_rows(df_valid: pd.DataFrame):
    if df_valid.empty: return 0
    if CSV_PATH.exists():
        existing = pd.read_csv(CSV_PATH, dtype=str).fillna("")
        for col in SCHEMA:
            if col not in existing.columns:
                existing[col] = ""
        existing = existing[SCHEMA]
    else:
        existing = pd.DataFrame(columns=SCHEMA)
        
    merged = pd.concat([existing, df_valid], ignore_index=True)
    merged.to_csv(CSV_PATH, index=False)
    return len(df_valid)


def log_errors(df_err: pd.DataFrame):
    if df_err.empty: return 0
    err_cols = SCHEMA + ['Error Reason']
    if ERROR_CSV_PATH.exists():
        existing_err = pd.read_csv(ERROR_CSV_PATH, dtype=str).fillna("")
        for col in err_cols:
            if col not in existing_err.columns:
                existing_err[col] = ""
        existing_err = existing_err[err_cols]
    else:
        existing_err = pd.DataFrame(columns=err_cols)
        
    merged = pd.concat([existing_err, df_err], ignore_index=True)
    merged.to_csv(ERROR_CSV_PATH, index=False)
    return len(df_err)


def main():
    if not GEMINI_API_KEY:
        print("ERROR: GEMINI_API_KEY not found in .env file.")
        return

    DONE_DIR.mkdir(parents=True, exist_ok=True)

    image_files = [
        f for f in IMAGES_DIR.iterdir()
        if f.is_file() and f.suffix.lower() in {'.jpg', '.jpeg', '.png', '.webp'}
    ]

    if not image_files:
        print(f"No images found in {IMAGES_DIR}. Drop images there and re-run.")
        return

    print(f"Found {len(image_files)} image(s) to process.\n")
    
    existing_keys = get_existing_keys()
    total_added = 0
    total_errors = 0

    for img in image_files:
        print(f"Processing: {img.name} ...")
        try:
            raw_rows = call_gemini(img)
            df = map_to_schema(raw_rows, img.name)
            
            df_valid, df_err = validate_and_split(df, existing_keys)
            
            added = merge_valid_rows(df_valid)
            errors = log_errors(df_err)
            
            if added > 0:
                new_keys = set(zip(df_valid['NAME OF MADRASAH'].str.strip().str.upper(), df_valid['PIN'].str.strip()))
                existing_keys.update(new_keys)

            print(f"  Gemini extracted {len(raw_rows)} row(s).")
            print(f"  Added {added} valid row(s).")
            if errors > 0:
                print(f"  Rejected {errors} row(s) to error_tracker.csv")
                
            shutil.move(str(img), DONE_DIR / img.name)
            print(f"  Moved to images/done/")
            total_added += added
            total_errors += errors
            
            # Prevent hitting the 15 RPM rate limit (wait 5 seconds between images)
            import time
            time.sleep(5)
        except Exception as e:
            print(f"  ERROR processing {img.name}: {e}")

    print(f"\nDone. Total valid rows added: {total_added}")
    if total_errors > 0:
        print(f"WARNING: {total_errors} row(s) were rejected. Check error_tracker.csv!")


if __name__ == "__main__":
    main()
