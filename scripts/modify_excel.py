import win32com.client as win32
import os

def modify_excel():
    filepath_in = os.path.abspath("forms.xls")
    filepath_out = os.path.abspath("forms_updated.xlsm")
    
    excel = win32.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    
    try:
        wb = excel.Workbooks.Open(filepath_in)
        
        # We will modify "Forms" and "Database"
        for sheet_name in ["Forms", "Database"]:
            ws = wb.Sheets(sheet_name)
            
            # Find the header row (assume row 4 for Forms, row 1 for Database based on standard practice or we can search)
            # Let's search for "CODE" or "NAME"
            header_row = 1
            for r in range(1, 10):
                if ws.Cells(r, 1).Value and str(ws.Cells(r, 1).Value).upper().strip() == "CODE":
                    header_row = r
                    break
                    
            # 1. Delete Expenditure, Expenditure in words, Head Maulana
            # We iterate backwards to safely delete columns
            for c in range(20, 0, -1):
                val = ws.Cells(header_row, c).Value
                if val:
                    val_str = str(val).upper()
                    if "EXPENDITURE" in val_str or "WORDS" in val_str or "HEAD MAULANA" in val_str:
                        ws.Columns(c).Delete()
            
            # 2. Split P.O. - P.S into Post Office and Police Station
            # First, find where P.O. - P.S is
            po_col = -1
            for c in range(1, 20):
                val = ws.Cells(header_row, c).Value
                if val and "P.O" in str(val).upper():
                    po_col = c
                    break
            
            if po_col != -1:
                ws.Cells(header_row, po_col).Value = "Post Office"
                # Insert a new column after Post Office for Police Station
                ws.Columns(po_col + 1).Insert()
                ws.Cells(header_row, po_col + 1).Value = "Police Station"
                
            # 3. Add Entry Date, Municipality/Block-Panchayat, Raabta Centre
            # Find last column
            last_col = 1
            for c in range(1, 30):
                if not ws.Cells(header_row, c).Value:
                    last_col = c
                    break
                    
            ws.Cells(header_row, last_col).Value = "Entry Date"
            ws.Cells(header_row, last_col + 1).Value = "Municipality/Block-Panchayat"
            ws.Cells(header_row, last_col + 2).Value = "Raabta Centre"
            
            # 4. Color the Municipality and Raabta Centre columns
            # Apply color to the column header and cells below
            ws.Cells(header_row, last_col + 1).Interior.Color = 10092543 # Light Yellow
            ws.Cells(header_row, last_col + 2).Interior.Color = 10092543
            
        # Save as a modern .xlsm file (FileFormat 52)
        wb.SaveAs(filepath_out, FileFormat=52)
        wb.Close(True)
        print(f"Successfully saved as {filepath_out}")
        
    except Exception as e:
        print(f"COM Error: {e}")
    finally:
        excel.Quit()

if __name__ == "__main__":
    modify_excel()
