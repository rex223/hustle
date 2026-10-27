import win32com.client, os

xl = win32com.client.Dispatch('Excel.Application')
xl.DisplayAlerts = False
xl.Visible = False
wb = xl.Workbooks.Open(os.path.abspath('RENTAL DATASHEET 19 OCT.xlsm'), ReadOnly=True)

wsRent = wb.Sheets('todays_rent_data')
print('=== todays_rent_data Rows 1 to 35 ===')
for r in range(1, min(36, wsRent.UsedRange.Rows.Count + 1)):
    grp = wsRent.Cells(r, 6).Value
    name = wsRent.Cells(r, 2).Value
    amt = wsRent.Cells(r, 4).Value
    paid = wsRent.Cells(r, 9).Value
    due = wsRent.Cells(r, 10).Value
    print(f'Row {r:2d} | Grp: {str(grp):15} | Name: {str(name):15} | Amt: {str(amt):8} | Paid: {str(paid):8} | Due: {str(due):8}')

wsExp = wb.Sheets('TODAYS_EXPENDITURE')
print('\n=== TODAYS_EXPENDITURE Headers ===')
for c in range(1, wsExp.UsedRange.Columns.Count + 1):
    h1 = wsExp.Cells(1, c).Value
    h2 = wsExp.Cells(2, c).Value
    if h1 or h2:
        col_letter = chr(64 + c) if c <= 26 else chr(64 + (c-1)//26) + chr(65 + (c-1)%26)
        print(f'Col {c:2d} ({col_letter:2s}): R1="{h1}" | R2="{h2}"')

wb.Close(False)
xl.Quit()
