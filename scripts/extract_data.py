import os
import pandas as pd
import io
from unstructured.partition.image import partition_image
import pytesseract

# Set tesseract path for windows
tesseract_path = r"D:\CLLG\hustle\.venv\Tesseract-OCR\tesseract.exe"
if os.path.exists(tesseract_path):
    pytesseract.pytesseract.tesseract_cmd = tesseract_path

def process_image(img_path):
    print(f"Processing {img_path}...")
    try:
        # hi_res strategy is usually needed for table extraction
        elements = partition_image(filename=img_path, strategy="hi_res", infer_table_structure=True)
        tables = [el for el in elements if el.category == "Table"]
        
        dfs = []
        for i, table in enumerate(tables):
            html_content = getattr(table.metadata, 'text_as_html', None)
            if html_content:
                try:
                    df_list = pd.read_html(io.StringIO(html_content))
                    if df_list:
                        dfs.append(df_list[0])
                except Exception as e:
                    print(f"Error parsing HTML table from element {i}: {e}")
        
        if dfs:
            return pd.concat(dfs, ignore_index=True)
        else:
            print(f"No tables found in {img_path}")
            return pd.DataFrame()
    except Exception as e:
        print(f"Failed to process {img_path}: {e}")
        return pd.DataFrame()

def clean_and_transform(df):
    if df.empty:
        return df
    
    # Standardize column names to upper case and strip whitespace
    df.columns = [str(c).upper().strip() for c in df.columns]
    
    # If the first row contains actual headers, we might need to promote it
    if "CODE" not in df.columns and len(df) > 0:
        row0 = [str(x).upper().strip() for x in df.iloc[0]]
        if "CODE" in row0 or "NAME OF MADRASAH" in row0:
            df.columns = row0
            df = df[1:].reset_index(drop=True)
            
    # Drop unwanted columns
    exclude_keywords = ["EXPENDITURE", "WORDS", "HEAD MAULANA"]
    cols_to_drop = []
    for col in df.columns:
        for kw in exclude_keywords:
            if kw in col:
                cols_to_drop.append(col)
                break
    
    if cols_to_drop:
        print(f"Dropping columns: {cols_to_drop}")
        df = df.drop(columns=cols_to_drop, errors='ignore')
        
    # Split P.O. - P.S into Post Office and Police Station
    po_ps_col = None
    for col in df.columns:
        if "P.O" in col or "P.S" in col:
            po_ps_col = col
            break
            
    if po_ps_col:
        print(f"Splitting column: {po_ps_col}")
        split_df = df[po_ps_col].astype(str).str.split(r'\s*-\s*', n=1, expand=True)
        if split_df.shape[1] == 2:
            df["Post Office"] = split_df[0]
            df["Police Station"] = split_df[1]
        else:
            df["Post Office"] = df[po_ps_col]
            df["Police Station"] = ""
        df = df.drop(columns=[po_ps_col])
        
    # Add new empty columns
    df["Entry Date"] = ""
    df["Municipality/Block-Panchayat"] = ""
    df["Raabta Centre"] = ""
    
    return df

def main():
    images = ["1.jpeg", "2.jpeg"]
    all_dfs = []
    
    for img in images:
        if os.path.exists(img):
            df = process_image(img)
            if not df.empty:
                df = clean_and_transform(df)
                all_dfs.append(df)
        else:
            print(f"Image not found: {img}")
            
    if all_dfs:
        final_df = pd.concat(all_dfs, ignore_index=True)
        final_df.to_csv("extracted_data.csv", index=False)
        print("Data extraction complete. Saved to extracted_data.csv.")
    else:
        print("No data extracted.")

if __name__ == "__main__":
    main()
