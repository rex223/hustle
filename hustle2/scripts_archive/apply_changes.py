"""
apply_changes.py
Applies the known numerical changes from the 8th Aug file to the 1st Aug file.
Run with Excel CLOSED.
"""
import win32com.client
import os

FILE_1ST = os.path.abspath('hustle2/RENTAL DATASHEET 1ST AUG.xlsm')
FILE_8TH = os.path.abspath('hustle2/RENTAL DATASHEET 8TH AUG.xlsm')

def apply():
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    
    try:
        wb1 = excel.Workbooks.Open(FILE_1ST)
        wb8 = excel.Workbooks.Open(FILE_8TH, ReadOnly=True)
        
        changes = [
            ('RENT DATA', 'M2'),
            ('RENT DATA', 'N2'),
            ('PURCHASE SHEET', 'C18'),
            ('PURCHASE SHEET', 'E18'),
            ('PURCHASE SHEET', 'G18'),
            ('PURCHASE SHEET', 'C22'),
            ('PURCHASE SHEET', 'E22'),
            ('PURCHASE SHEET', 'F22'),
            ('PURCHASE SHEET', 'G22'),
        ]
        
        for sheet_name, cell in changes:
            ws1 = wb1.Sheets(sheet_name)
            ws8 = wb8.Sheets(sheet_name)
            
            val = ws8.Range(cell).Value
            ws1.Range(cell).Value = val
            print(f"Updated {sheet_name} {cell} to {val}")
            
        wb1.Save()
        print("SUCCESS: All changes applied.")
        
    except Exception as e:
        import traceback
        traceback.print_exc()
        print(f"Error: {e}")
    finally:
        try:
            wb1.Close(False)
            wb8.Close(False)
        except:
            pass
        excel.Quit()

if __name__ == '__main__':
    apply()
