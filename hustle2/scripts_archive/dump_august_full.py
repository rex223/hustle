import win32com.client
import os

xl = win32com.client.Dispatch('Excel.Application')
xl.DisplayAlerts = False
xl.Visible = False

wb = xl.Workbooks.Open(os.path.abspath('RENTAL DATASHEET 19 OCT.xlsm'), ReadOnly=True)
wsDB = wb.Sheets('database_expenditure')

print("=== ALL AUGUST ROWS (34 to 65) IN database_expenditure ===")
for r in range(34, 66):
    non_empty = []
    for c in range(1, 46):
        v = wsDB.Cells(r, c).Value
        fmt = wsDB.Cells(r, c).NumberFormat
        if v is not None:
            non_empty.append((c, v, fmt))
    if non_empty:
        print(f"Row {r:2d} ({wsDB.Cells(r, 1).Value}): {non_empty}\n")

wb.Close(False)
xl.Quit()
