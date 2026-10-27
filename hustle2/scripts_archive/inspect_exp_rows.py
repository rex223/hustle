import win32com.client
import os

xl = win32com.client.Dispatch('Excel.Application')
xl.DisplayAlerts = False
xl.Visible = False

wb = xl.Workbooks.Open(os.path.abspath('RENTAL DATASHEET 19 OCT.xlsm'), ReadOnly=True)
ws = wb.Sheets('TODAYS_EXPENDITURE')
print(f"TODAYS_EXPENDITURE rows (total {ws.UsedRange.Rows.Count}):")
for r in range(1, min(40, ws.UsedRange.Rows.Count + 1)):
    val_a = ws.Cells(r, 1).Value
    val_b = ws.Cells(r, 2).Value
    val_c = ws.Cells(r, 3).Value
    val_d = ws.Cells(r, 4).Value
    val_e = ws.Cells(r, 5).Value
    if any(v is not None for v in [val_a, val_b, val_c, val_d, val_e]):
        print(f"Row {r:2d} | Col A: {str(val_a):20} | Col B: {str(val_b):10} | Col C: {str(val_c):10} | Col D: {str(val_d):10} | Col E: {str(val_e):10}")

wb.Close(False)
xl.Quit()
