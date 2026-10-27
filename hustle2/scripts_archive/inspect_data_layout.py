import win32com.client
import os

xl = win32com.client.Dispatch('Excel.Application')
xl.DisplayAlerts = False
xl.Visible = False

wb = xl.Workbooks.Open(os.path.abspath('RENTAL DATASHEET 19 OCT.xlsm'), ReadOnly=True)
wsToday = wb.Sheets('TODAYS_EXPENDITURE')
wsDB = wb.Sheets('database_expenditure')

print("=== ALL NON-EMPTY CELLS IN TODAYS_EXPENDITURE ROWS 1 to 15 ===")
for r in range(1, 16):
    row_vals = [(c, wsToday.Cells(r, c).Value, wsToday.Cells(r, c).NumberFormat) 
                for c in range(1, 46) if wsToday.Cells(r, c).Value is not None]
    if row_vals:
        print(f"Row {r:2d}: {row_vals}")

print("\n=== ALL NON-EMPTY CELLS IN TODAYS_EXPENDITURE ROWS 32 to 45 ===")
for r in range(32, 46):
    row_vals = [(c, wsToday.Cells(r, c).Value) 
                for c in range(1, 46) if wsToday.Cells(r, c).Value is not None]
    if row_vals:
        print(f"Row {r:2d}: {row_vals}")

print("\n=== ALL NON-EMPTY CELLS IN database_expenditure ROW 1 to 10 ===")
for r in range(1, 11):
    row_vals = [(c, wsDB.Cells(r, c).Value, wsDB.Cells(r, c).NumberFormat) 
                for c in range(1, 46) if wsDB.Cells(r, c).Value is not None]
    if row_vals:
        print(f"Row {r:2d}: {row_vals}")

wb.Close(False)
xl.Quit()
