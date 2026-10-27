import win32com.client
import os

xl = win32com.client.Dispatch('Excel.Application')
xl.DisplayAlerts = False
xl.Visible = False

wb = xl.Workbooks.Open(os.path.abspath('RENTAL DATASHEET 19 OCT.xlsm'), ReadOnly=True)
wsToday = wb.Sheets('TODAYS_EXPENDITURE')
wsDB = wb.Sheets('database_expenditure')

print("=== TODAYS_EXPENDITURE Cols 40 to 45 ===")
for r in range(1, 35):
    row_vals = [(c, wsToday.Cells(r, c).Value) for c in range(40, 46) if wsToday.Cells(r, c).Value is not None]
    if row_vals:
        print(f"Row {r:2d}: {row_vals}")

print("\n=== database_expenditure Cols 40 to 45 (July: 2-34) ===")
for r in range(2, 35):
    row_vals = [(c, wsDB.Cells(r, c).Value) for c in range(40, 46) if wsDB.Cells(r, c).Value is not None]
    if row_vals:
        print(f"Row {r:2d}: {row_vals}")

print("\n=== database_expenditure Cols 40 to 45 (August: 34-65) ===")
for r in range(34, 66):
    row_vals = [(c, wsDB.Cells(r, c).Value) for c in range(40, 46) if wsDB.Cells(r, c).Value is not None]
    if row_vals:
        print(f"Row {r:2d}: {row_vals}")

wb.Close(False)
xl.Quit()
