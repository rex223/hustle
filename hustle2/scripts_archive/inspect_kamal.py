import win32com.client
import os

xl = win32com.client.Dispatch('Excel.Application')
xl.DisplayAlerts = False
xl.Visible = False

wb = xl.Workbooks.Open(os.path.abspath('RENTAL DATASHEET 19 OCT.xlsm'), ReadOnly=True)
if 'KAMAL' in [s.Name for s in wb.Sheets]:
    ws = wb.Sheets('KAMAL')
    print("KAMAL Sheet Rows 1 to 10:")
    for r in range(1, min(10, ws.UsedRange.Rows.Count + 1)):
        row_vals = [ws.Cells(r, c).Value for c in range(1, 9)]
        print(f"Row {r:2d}: {row_vals}")
wb.Close(False)
xl.Quit()
