import win32com.client as win32
import os

def inspect_with_com():
    filepath = os.path.abspath("forms.xlsm")
    print(f"Opening {filepath}...")
    
    excel = win32.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    
    try:
        wb = excel.Workbooks.Open(filepath)
        for i in range(1, wb.Sheets.Count + 1):
            ws = wb.Sheets(i)
            print(f"Sheet: {ws.Name}")
            for row in range(1, 6):
                row_vals = []
                for col in range(1, 15):
                    val = ws.Cells(row, col).Value
                    if val is not None:
                        row_vals.append(str(val))
                if row_vals:
                    print(f" Row {row}: {row_vals}")
        wb.Close(False)
    except Exception as e:
        print(f"COM Error: {e}")
    finally:
        excel.Quit()

inspect_with_com()
