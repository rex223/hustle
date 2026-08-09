import win32com.client
import os
import traceback

excel = win32com.client.Dispatch('Excel.Application')
excel.Visible = False
excel.DisplayAlerts = False
file_path = os.path.abspath('RENTAL DATASHEET 1ST AUG.xlsm')

try:
    wb = excel.Workbooks.Open(file_path)
    
    # We need to update the VBA code in ThisWorkbook
    workbook_module = wb.VBProject.VBComponents("ThisWorkbook").CodeModule
    
    lines = workbook_module.Lines(1, workbook_module.CountOfLines)
    
    # Replace the ignoreSheets strings to include PURCHASE SHEET DAILY
    lines = lines.replace(
        'ignoreSheets = "|PAYMENT SHEET|EXPENDITURE|PURCHASE|SHEET17|"',
        'ignoreSheets = "|PAYMENT SHEET|EXPENDITURE|PURCHASE|PURCHASE SHEET DAILY|SHEET17|"'
    )
    
    lines = lines.replace(
        'ignoreSheets = "|RENT DATA|PAYMENT SHEET|EXPENDITURE|PURCHASE|SHEET17|"',
        'ignoreSheets = "|RENT DATA|PAYMENT SHEET|EXPENDITURE|PURCHASE|PURCHASE SHEET DAILY|SHEET17|"'
    )
    
    # Clear and rewrite
    workbook_module.DeleteLines(1, workbook_module.CountOfLines)
    workbook_module.AddFromString(lines)
    
    # Run the summary update again to clear the bad data
    excel.Application.Run("ThisWorkbook.UpdateGroupSummary")
    
    wb.Save()
    print("SUCCESS")
except Exception as e:
    traceback.print_exc()
finally:
    try:
        wb.Close(False)
    except:
        pass
    excel.Quit()
