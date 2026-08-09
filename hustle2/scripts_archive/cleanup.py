import win32com.client
import os

excel = win32com.client.Dispatch('Excel.Application')
excel.Visible = False
excel.DisplayAlerts = False
file_path = os.path.abspath('RENTAL DATASHEET 1ST AUG.xlsm')

try:
    wb = excel.Workbooks.Open(file_path)
    
    # 1. Rename Sheets
    ws_db = None
    ws_entry = None
    for s in wb.Sheets:
        if s.Name.upper() == "PURCHASE SHEET DAILY":
            ws_db = s
        elif s.Name.upper() == "TODAYS PURCHASES":
            ws_entry = s
            
    if ws_db:
        ws_db.Name = "DATABASE_PURCHASE"
    if ws_entry:
        ws_entry.Name = "PURCHASE SHEET"
        
    # 2. Update Date Formats
    if ws_db:
        ws_db.Columns("A").NumberFormat = "dd/mm/yyyy"
    if ws_entry:
        ws_entry.Range("C1").NumberFormat = "dd/mm/yyyy"
        
    # 3. Update VBA in ThisWorkbook
    workbook_module = wb.VBProject.VBComponents("ThisWorkbook").CodeModule
    if workbook_module.CountOfLines > 0:
        lines = workbook_module.Lines(1, workbook_module.CountOfLines)
        lines = lines.replace("|PURCHASE SHEET DAILY|", "|DATABASE_PURCHASE|PURCHASE SHEET|")
        workbook_module.DeleteLines(1, workbook_module.CountOfLines)
        workbook_module.AddFromString(lines)
        
    # 4. Update VBA in ModDailyPurchases
    for comp in wb.VBProject.VBComponents:
        if comp.Name == "ModDailyPurchases":
            lines = comp.CodeModule.Lines(1, comp.CodeModule.CountOfLines)
            lines = lines.replace('"TODAYS PURCHASES"', '"PURCHASE SHEET"')
            lines = lines.replace('"PURCHASE SHEET DAILY"', '"DATABASE_PURCHASE"')
            comp.CodeModule.DeleteLines(1, comp.CodeModule.CountOfLines)
            comp.CodeModule.AddFromString(lines)
            break
            
    # 5. Update VBA in the PURCHASE SHEET Worksheet module
    if ws_entry:
        ws_module = wb.VBProject.VBComponents(ws_entry.CodeName).CodeModule
        if ws_module.CountOfLines > 0:
            lines = ws_module.Lines(1, ws_module.CountOfLines)
            lines = lines.replace('"PURCHASE SHEET DAILY"', '"DATABASE_PURCHASE"')
            ws_module.DeleteLines(1, ws_module.CountOfLines)
            ws_module.AddFromString(lines)

    wb.Save()
    print("SUCCESS")
except Exception as e:
    print(f"Error: {e}")
finally:
    try:
        wb.Close(False)
    except:
        pass
    excel.Quit()
