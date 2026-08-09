import win32com.client
import os
import traceback

excel = win32com.client.Dispatch('Excel.Application')
excel.Visible = False
excel.DisplayAlerts = False
file_path = os.path.abspath('RENTAL DATASHEET 1ST AUG.xlsm')

try:
    wb = excel.Workbooks.Open(file_path)
    ws = wb.Sheets('PURCHASE SHEET DAILY')
    
    tbl = ws.ListObjects('DailyPurchaseLedger')
    col_a = tbl.ListColumns(1).DataBodyRange
    
    print('Deleting FC')
    col_a.FormatConditions.Delete()
    print('Adding FC')
    fc = col_a.FormatConditions.Add(2, Formula1="=$A2=$A1")
    print('Setting NumberFormat')
    fc.NumberFormat = ';;;'
    fc.StopIfTrue = False
    
    print('SUCCESS')
except Exception as e:
    traceback.print_exc()
finally:
    try:
        wb.Close(False)
    except:
        pass
    excel.Quit()
