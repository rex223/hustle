import win32com.client
import os

xl = win32com.client.Dispatch('Excel.Application')
xl.DisplayAlerts = False
xl.Visible = False

wb = xl.Workbooks.Open(os.path.abspath('RENTAL DATASHEET 19 OCT.xlsm'), ReadOnly=True)
ws = wb.Sheets('TODAYS_EXPENDITURE')
print("TODAYS_EXPENDITURE Rows 30 to 45:")
for r in range(30, min(45, ws.UsedRange.Rows.Count + 1)):
    row_vals = [ws.Cells(r, c).Value for c in range(1, 10)]
    print(f"Row {r:2d}: {row_vals}")

wb.Close(False)
xl.Quit()
