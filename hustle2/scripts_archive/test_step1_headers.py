import win32com.client
import os
import datetime

def run_fix():
    xl = win32com.client.Dispatch('Excel.Application')
    xl.DisplayAlerts = False
    xl.Visible = False

    wb_path = os.path.abspath('RENTAL DATASHEET 19 OCT.xlsm')
    wb = xl.Workbooks.Open(wb_path)

    wsDB = wb.Sheets('database_expenditure')
    wsToday = wb.Sheets('TODAYS_EXPENDITURE')

    # Standard 44 Column Headers
    headers = [
        "NAME", "SALARIES", "TYPE", "DATE", "PAID STATUS",
        "DATE", "ADV", "AMOUNT", "PAID",
        "DATE", "CESC / URP", "AMOUNT", "PAID",
        "DATE", "CESC / RESD", "AMOUNT", "PAID",
        "DATE", "LABOUR/URP", "AMOUNT", "PAID",
        "DATE", "DIGRUM / URP", "AMOUNT", "PAID",
        "DATE", "ANIMAL FEED", "AMOUNT", "PAID",
        "DATE", "FUEL/GAS", "AMOUNT", "PAID",
        "DATE", "MEDICAL", "AMOUNT", "PAID",
        "DATE", "LOCATION", "REPAIR", "AMOUNT", "PAID",
        "RELIGION", "AMOUNT"
    ]

    print("Step 1: Updating Row 1 Headers on both sheets...")
    for c, h in enumerate(headers, 1):
        wsDB.Cells(1, c).Value = h
        wsToday.Cells(1, c).Value = h

    # Clear excess old header cells beyond col 44
    for c in range(45, 55):
        wsDB.Cells(1, c).Value = None
        wsToday.Cells(1, c).Value = None

    print("Headers updated successfully!")
    wb.Save()
    wb.Close(False)
    xl.Quit()

if __name__ == '__main__':
    run_fix()
