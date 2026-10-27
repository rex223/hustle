import win32com.client
import os

xl = win32com.client.Dispatch('Excel.Application')
xl.DisplayAlerts = False
xl.Visible = False

wb = xl.Workbooks.Open(os.path.abspath('RENTAL DATASHEET 19 OCT.xlsm'), ReadOnly=True)
ws = wb.Sheets('TODAYS_EXPENDITURE')

data = []
for r in range(3, 33):
    row_vals = [ws.Cells(r, c).Value for c in range(1, 6)]
    data.append((r, row_vals))

wb.Close(False)
xl.Quit()

for r, vals in data:
    print(f"Row {r:2d}: {vals}")
