import pandas as pd

try:
    df_forms = pd.read_excel("forms.xlsm", engine="xlrd", sheet_name="Forms")
    print("Forms columns:", df_forms.columns.tolist())
    
    df_db = pd.read_excel("forms.xlsm", engine="xlrd", sheet_name="Database")
    print("Database columns:", df_db.columns.tolist())
except Exception as e:
    print(f"Error: {e}")
