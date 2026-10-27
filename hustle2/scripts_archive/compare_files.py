"""
compare_files.py
Compares numerical values and new inputs from 8TH AUG to 1ST AUG file.
"""
import win32com.client
import os

FILE_1ST = os.path.abspath('hustle2/RENTAL DATASHEET 1ST AUG.xlsm')
FILE_8TH = os.path.abspath('hustle2/RENTAL DATASHEET 8TH AUG.xlsm')

def compare_sheets():
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    
    changes = []
    
    try:
        wb1 = excel.Workbooks.Open(FILE_1ST, ReadOnly=True)
        wb8 = excel.Workbooks.Open(FILE_8TH, ReadOnly=True)
        
        # We will check common data sheets
        sheets_to_check = ['RENT DATA', 'PAYMENT SHEET', 'PURCHASE SHEET', 'TODAYS EXPENDITURE', 'DATABASE_EXPENDITURE']
        
        for sheet_name in sheets_to_check:
            try:
                ws1 = wb1.Sheets(sheet_name)
                ws8 = wb8.Sheets(sheet_name)
            except:
                print(f"Skipping {sheet_name}, not found in both files.")
                continue
                
            # Get used range bounds from 8th Aug
            last_row = ws8.Cells.SpecialCells(11).Row # xlCellTypeLastCell
            last_col = ws8.Cells.SpecialCells(11).Column
            
            # To be efficient, grab values in bulk
            vals1 = ws1.Range(ws1.Cells(1, 1), ws1.Cells(last_row, last_col)).Value
            vals8 = ws8.Range(ws8.Cells(1, 1), ws8.Cells(last_row, last_col)).Value
            
            if not vals8:
                continue
                
            for r in range(last_row):
                for c in range(last_col):
                    v8 = vals8[r][c] if r < len(vals8) and c < len(vals8[r]) else None
                    v1 = vals1[r][c] if vals1 and r < len(vals1) and c < len(vals1[r]) else None
                    
                    if v8 is not None and v8 != '' and v8 != v1:
                        # Only numerical or date/text inputs we care about
                        # Filter out formulas or styling differences by only checking values
                        if type(v8) in (int, float) or (type(v8) is str and str(v8).strip() != ''):
                            col_letter = ws8.Cells(r+1, c+1).Address.split('$')[1]
                            changes.append({
                                'sheet': sheet_name,
                                'cell': f"{col_letter}{r+1}",
                                'old': v1,
                                'new': v8
                            })
                            
        wb1.Close(False)
        wb8.Close(False)
        excel.Quit()
        
        if not changes:
            print("No numerical data or text changes found.")
        else:
            print(f"Found {len(changes)} changes:")
            for ch in changes[:100]:  # print up to 100 changes
                old_val = ch['old'] if ch['old'] is not None else '<Empty>'
                print(f"Sheet: {ch['sheet']:<20} Cell: {ch['cell']:<5} | 1st Aug: {str(old_val):<15} -> 8th Aug: {str(ch['new'])}")
                
    except Exception as e:
        import traceback
        traceback.print_exc()
        try:
            wb1.Close(False)
            wb8.Close(False)
        except:
            pass
        excel.Quit()

if __name__ == '__main__':
    compare_sheets()
