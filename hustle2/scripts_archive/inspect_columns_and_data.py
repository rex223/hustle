import win32com.client
import os

xl = win32com.client.Dispatch('Excel.Application')
xl.DisplayAlerts = False
xl.Visible = False

wb = xl.Workbooks.Open(os.path.abspath('RENTAL DATASHEET 19 OCT.xlsm'), ReadOnly=True)
wsToday = wb.Sheets('TODAYS_EXPENDITURE')
wsDB = wb.Sheets('database_expenditure')

print("=== TODAYS_EXPENDITURE Headers (Row 1) ===")
today_headers = []
for c in range(1, 50):
    val = wsToday.Cells(1, c).Value
    today_headers.append((c, val))
    if val is not None:
        print(f"Col {c:2d}: '{val}'")

print("\n=== database_expenditure Headers (Row 1) ===")
db_headers = []
for c in range(1, 50):
    val = wsDB.Cells(1, c).Value
    db_headers.append((c, val))
    if val is not None:
        print(f"Col {c:2d}: '{val}'")

print("\n=== Compare TODAYS_EXPENDITURE vs database_expenditure Row 1 ===")
max_cols = max(wsToday.UsedRange.Columns.Count, wsDB.UsedRange.Columns.Count, 45)
for c in range(1, max_cols + 1):
    h_today = wsToday.Cells(1, c).Value
    h_db = wsDB.Cells(1, c).Value
    if h_today or h_db:
        print(f"Col {c:2d} | Today: {str(h_today):20s} | DB: {str(h_db):20s}")

wb.Close(False)
xl.Quit()
