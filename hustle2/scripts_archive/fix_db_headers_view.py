"""
fix_db_headers_view.py
Fixes the scroll position and freeze panes on DATABASE_EXPENDITURE so row 1 (the headers) is visible.
Run with Excel CLOSED.
"""
import win32com.client
import os

WORKBOOK_PATH = os.path.abspath('hustle2/RENTAL DATASHEET 1ST AUG.xlsm')

def run():
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    try:
        wb = excel.Workbooks.Open(WORKBOOK_PATH)
        ws_db = wb.Sheets("DATABASE_EXPENDITURE")
        
        ws_db.Activate()
        
        # 1. Unfreeze panes
        excel.ActiveWindow.FreezePanes = False
        
        # 2. Scroll to the very top left
        excel.ActiveWindow.ScrollRow = 1
        excel.ActiveWindow.ScrollColumn = 1
        
        # 3. Double check row 1 height just in case it was hidden
        if ws_db.Rows("1:1").RowHeight == 0:
            ws_db.Rows("1:1").RowHeight = 22
            
        # 4. Select row 2 to freeze row 1 above it
        ws_db.Cells(2, 1).Select()
        excel.ActiveWindow.FreezePanes = True

        wb.Save()
        print("SUCCESS: DATABASE_EXPENDITURE scroll and freeze panes fixed.")

    except Exception as e:
        import traceback
        traceback.print_exc()
        print(f"Error: {e}")
    finally:
        try:
            wb.Close(False)
        except Exception:
            pass
        excel.Quit()

if __name__ == '__main__':
    run()
