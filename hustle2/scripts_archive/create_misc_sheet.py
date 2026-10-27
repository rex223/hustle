import win32com.client
import os
import traceback

xl = win32com.client.Dispatch('Excel.Application')
xl.DisplayAlerts = False
wb = xl.Workbooks.Open(os.path.abspath('RENTAL DATASHEET 3 SEP.xlsm'))

try:
    # Create the new sheet at the end of the workbook
    new_sh = None
    for sh in wb.Sheets:
        if sh.Name.upper() == 'MISCELLANEOUS':
            new_sh = sh
            break
            
    if new_sh is None:
        new_sh = wb.Sheets.Add(After=wb.Sheets(wb.Sheets.Count))
        new_sh.Name = 'MISCELLANEOUS'
    
    new_sh.Cells.Clear()
    
    # Set headers
    headers = ["DATE", "PAID TO", "DESCRIPTION", "AMOUNT"]
    for i, h in enumerate(headers):
        new_sh.Cells(1, i + 1).Value = h
        
    # Add a blank row to start
    new_sh.Cells(2, 1).Value = ""
        
    # Format as a table
    table_range = new_sh.Range("A1:D2")
    # ListObjects is a property of Worksheet!
    new_sh.ListObjects.Add(1, table_range, 1, 1).Name = "MiscPaymentsTable"
    
    # Adjust column widths
    new_sh.Columns("A:A").ColumnWidth = 15
    new_sh.Columns("B:B").ColumnWidth = 25
    new_sh.Columns("C:C").ColumnWidth = 35
    new_sh.Columns("D:D").ColumnWidth = 15
    
    # Format Date column
    new_sh.Columns("A:A").NumberFormat = "dd-mmm-yyyy"
    # Format Amount column
    new_sh.Columns("D:D").NumberFormat = "#,##0"
    
    wb.Save()
    print('SUCCESS: Created MISCELLANEOUS sheet')
except Exception as e:
    traceback.print_exc()
finally:
    wb.Close(False)
    xl.Quit()
