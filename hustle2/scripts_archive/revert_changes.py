"""
revert_changes.py
Reverts the numerical changes made from the 8th Aug file to the 1st Aug file.
Run with Excel CLOSED.
"""
import win32com.client
import os

FILE_1ST = os.path.abspath('hustle2/RENTAL DATASHEET 1ST AUG.xlsm')

def revert():
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    
    try:
        wb1 = excel.Workbooks.Open(FILE_1ST)
        
        reversions = [
            ('RENT DATA', 'M2', 46000.0),
            ('RENT DATA', 'N2', 200600.0),
            ('PURCHASE SHEET', 'C18', None),
            ('PURCHASE SHEET', 'E18', 0.0),
            ('PURCHASE SHEET', 'G18', None),
            ('PURCHASE SHEET', 'C22', None),
            ('PURCHASE SHEET', 'E22', 0.0),
            ('PURCHASE SHEET', 'F22', None),
            ('PURCHASE SHEET', 'G22', None),
        ]
        
        for sheet_name, cell, val in reversions:
            ws1 = wb1.Sheets(sheet_name)
            
            if val is None:
                ws1.Range(cell).ClearContents()
            else:
                ws1.Range(cell).Value = val
                
            print(f"Reverted {sheet_name} {cell} to {val}")
            
        wb1.Save()
        print("SUCCESS: All changes reverted.")
        
    except Exception as e:
        import traceback
        traceback.print_exc()
        print(f"Error: {e}")
    finally:
        try:
            wb1.Close(False)
        except:
            pass
        excel.Quit()

if __name__ == '__main__':
    revert()
