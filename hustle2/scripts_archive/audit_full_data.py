import win32com.client, os

xl = win32com.client.Dispatch('Excel.Application')
xl.DisplayAlerts = False
xl.Visible = False
wb = xl.Workbooks.Open(os.path.abspath('RENTAL DATASHEET 19 OCT.xlsm'), ReadOnly=True)

wsDB = wb.Sheets('database_expenditure')
wsToday = wb.Sheets('TODAYS_EXPENDITURE')

print('=== TODAY Row 1 ===')
for c in range(1, 45):
    v = wsToday.Cells(1, c).Value
    if v: print(f'Col {c:2d}: {v}')

print('\n=== DB Row 1 ===')
for c in range(1, 45):
    v = wsDB.Cells(1, c).Value
    if v: print(f'Col {c:2d}: {v}')

print('\n=== DB JULY Rows (Rows 2 to 33) non-empty cells ===')
for r in range(2, 34):
    row_data = [(c, wsDB.Cells(r, c).Value, wsDB.Cells(r, c).NumberFormat) for c in range(1, 45) if wsDB.Cells(r, c).Value is not None]
    if row_data:
        print(f'R{r:2d} ({wsDB.Cells(r, 1).Value}): {row_data}')

print('\n=== DB AUGUST Rows (Rows 34 to 65) non-empty cells ===')
for r in range(34, 66):
    row_data = [(c, wsDB.Cells(r, c).Value, wsDB.Cells(r, c).NumberFormat) for c in range(1, 45) if wsDB.Cells(r, c).Value is not None]
    if row_data:
        print(f'R{r:2d} ({wsDB.Cells(r, 1).Value}): {row_data}')

wb.Close(False)
xl.Quit()
