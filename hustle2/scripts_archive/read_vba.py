import win32com.client
import os

excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = False
wb = excel.Workbooks.Open(os.path.abspath('RENTAL DATASHEET 1ST AUG.xlsm'))

try:
    m = wb.VBProject.VBComponents("ThisWorkbook").CodeModule
    lines = m.CountOfLines
    if lines > 0:
        print(m.Lines(1, lines))
    else:
        print("No lines in ThisWorkbook.")
finally:
    wb.Close(False)
    excel.Quit()
