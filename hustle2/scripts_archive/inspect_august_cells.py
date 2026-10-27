import win32com.client
import os

xl = win32com.client.Dispatch('Excel.Application')
xl.DisplayAlerts = False
xl.Visible = False

wb = xl.Workbooks.Open(os.path.abspath('RENTAL DATASHEET 19 OCT.xlsm'), ReadOnly=True)
wsDB = wb.Sheets('database_expenditure')

print("=== ALL CELLS IN AUGUST (Rows 34 to 65) IN database_expenditure ===")
for r in range(34, 66):
    for c in range(1, 46):
        v = wsDB.Cells(r, c).Value
        if v is not None:
            fmt = wsDB.Cells(r, c).NumberFormat
            print(f"R{r:2d} C{c:2d} [{wsDB.Cells(1, c).Value}]: val={v} | type={type(v).__name__} | fmt={fmt}")

wb.Close(False)
xl.Quit()
