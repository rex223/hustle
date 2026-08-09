import win32com.client
import os

excel = win32com.client.Dispatch('Excel.Application')
excel.Visible = False
excel.DisplayAlerts = False
file_path = os.path.abspath('RENTAL DATASHEET 1ST AUG.xlsm')

vba_macro = '''
Public Sub SaveDailyPurchases()
    Dim wsEntry As Worksheet
    Dim wsDB As Worksheet
    Dim tblEntry As ListObject
    Dim tblDB As ListObject
    
    On Error Resume Next
    Set wsEntry = ThisWorkbook.Sheets("TODAYS PURCHASES")
    Set wsDB = ThisWorkbook.Sheets("PURCHASE SHEET DAILY")
    Set tblEntry = wsEntry.ListObjects("EntryFormTable")
    Set tblDB = wsDB.ListObjects("DailyPurchaseLedger")
    On Error GoTo 0
    
    If wsEntry Is Nothing Or wsDB Is Nothing Then
        MsgBox "Error: Sheets not found!", vbCritical
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    Dim i As Long
    Dim qtyKg As Variant, qtyPcs As Variant, trsp As Variant
    Dim r As ListRow
    
    Dim itemsSaved As Long
    itemsSaved = 0
    
    For i = 1 To tblEntry.DataBodyRange.Rows.Count
        qtyKg = tblEntry.DataBodyRange.Cells(i, 3).Value
        qtyPcs = tblEntry.DataBodyRange.Cells(i, 4).Value
        trsp = tblEntry.DataBodyRange.Cells(i, 6).Value
        
        ' If they entered ANY quantity or transport cost, save it!
        If (IsNumeric(qtyKg) And qtyKg > 0) Or (IsNumeric(qtyPcs) And qtyPcs > 0) Or (IsNumeric(trsp) And trsp > 0) Then
            Set r = tblDB.ListRows.Add
            r.Range(1, 1).Value = Date ' Today's Date
            r.Range(1, 2).Value = tblEntry.DataBodyRange.Cells(i, 1).Value ' GROUP
            r.Range(1, 3).Value = tblEntry.DataBodyRange.Cells(i, 2).Value ' ITEM
            r.Range(1, 4).Value = qtyKg ' QTY KG
            r.Range(1, 5).Value = qtyPcs ' QTY PCS
            r.Range(1, 6).Value = tblEntry.DataBodyRange.Cells(i, 5).Value ' RATE
            r.Range(1, 7).Value = trsp ' TRSP
            r.Range(1, 8).Value = tblEntry.DataBodyRange.Cells(i, 7).Value ' AMOUNT
            
            ' Clear the entry form fields for this row
            tblEntry.DataBodyRange.Cells(i, 3).ClearContents
            tblEntry.DataBodyRange.Cells(i, 4).ClearContents
            tblEntry.DataBodyRange.Cells(i, 6).ClearContents
            
            itemsSaved = itemsSaved + 1
        End If
    Next i
    
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    
    If itemsSaved > 0 Then
        MsgBox itemsSaved & " items successfully saved to the ledger!", vbInformation, "Success"
    Else
        MsgBox "No quantities entered. Nothing to save.", vbExclamation, "No Data"
    End If
End Sub
'''

try:
    wb = excel.Workbooks.Open(file_path)
    
    # 1. Inject Macro into a Standard Module
    mod_name = "ModDailyPurchases"
    module_exists = False
    for comp in wb.VBProject.VBComponents:
        if comp.Name == mod_name:
            module_exists = True
            comp.CodeModule.DeleteLines(1, comp.CodeModule.CountOfLines)
            comp.CodeModule.AddFromString(vba_macro)
            break
            
    if not module_exists:
        new_module = wb.VBProject.VBComponents.Add(1) # 1 = vbext_ct_StdModule
        new_module.Name = mod_name
        new_module.CodeModule.AddFromString(vba_macro)

    # 2. Extract Unique Items from PURCHASE SHEET DAILY
    ws_db = wb.Sheets('PURCHASE SHEET DAILY')
    tbl_db = ws_db.ListObjects('DailyPurchaseLedger')
    
    unique_items = {} # key: ITEM NAME, value: (GROUP, RATE)
    for i in range(1, tbl_db.DataBodyRange.Rows.Count + 1):
        item = tbl_db.DataBodyRange.Cells(i, 3).Value
        group = tbl_db.DataBodyRange.Cells(i, 2).Value
        rate = tbl_db.DataBodyRange.Cells(i, 6).Value
        
        if item and str(item).strip() != "":
            item = str(item).strip()
            # Keep the most recent rate/group we see
            unique_items[item] = (group if group else "", rate if rate else 0)

    # 3. Create or reset TODAYS PURCHASES sheet
    sheet_name = "TODAYS PURCHASES"
    ws_entry = None
    for s in wb.Sheets:
        if s.Name.upper() == sheet_name:
            ws_entry = s
            break
            
    if ws_entry:
        ws_entry.Cells.Clear()
        for shape in ws_entry.Shapes:
            shape.Delete()
        for tbl in ws_entry.ListObjects:
            tbl.Delete()
    else:
        ws_entry = wb.Sheets.Add(Before=wb.Sheets(1))
        ws_entry.Name = sheet_name
        
    # 4. Setup Headers
    ws_entry.Range("A3:G3").Value = ["GROUP / FOR", "ITEM NAME", "QTY (KG)", "QTY (PCS)", "RATE", "TRSP", "TOTAL AMOUNT"]
    
    # 5. Populate Unique Items
    row_idx = 4
    for item, (group, rate) in unique_items.items():
        ws_entry.Cells(row_idx, 1).Value = group
        ws_entry.Cells(row_idx, 2).Value = item
        ws_entry.Cells(row_idx, 5).Value = rate
        # Set formula for TOTAL AMOUNT (Column G)
        # =IF(SUM(C4:D4, F4)>0, (C4+D4)*E4+F4, "")
        formula = f"=IF(SUM(C{row_idx}:D{row_idx}, F{row_idx})>0, (C{row_idx}+D{row_idx})*E{row_idx}+F{row_idx}, \"\")"
        ws_entry.Cells(row_idx, 7).Formula = formula
        row_idx += 1
        
    # 6. Format as Table
    table_range = f"A3:G{row_idx - 1 if row_idx > 4 else 4}"
    tbl_entry = ws_entry.ListObjects.Add(1, ws_entry.Range(table_range), 1, 1)
    tbl_entry.Name = "EntryFormTable"
    tbl_entry.TableStyle = "TableStyleMedium9"
    
    # Widths and alignment
    ws_entry.Columns("A:B").ColumnWidth = 20
    ws_entry.Columns("C:G").ColumnWidth = 15
    ws_entry.Rows("1:2").RowHeight = 30
    
    # Lock columns so user knows where to type (color QTY and TRSP columns slightly lighter)
    # Actually, let's just color the input columns yellow/green to highlight them
    # tbl_entry.ListColumns("QTY (KG)").DataBodyRange.Interior.Color = 13434879 # Light Yellow
    # tbl_entry.ListColumns("QTY (PCS)").DataBodyRange.Interior.Color = 13434879
    # tbl_entry.ListColumns("TRSP").DataBodyRange.Interior.Color = 13434879
    
    # 7. Create the Button (Shape)
    btn = ws_entry.Shapes.AddShape(1, 10, 10, 200, 40) # msoShapeRectangle
    btn.TextFrame.Characters().Text = "💾 SAVE TO LEDGER"
    btn.TextFrame.Characters().Font.Bold = True
    btn.TextFrame.Characters().Font.Size = 14
    btn.TextFrame.Characters().Font.Color = 16777215 # White
    btn.Fill.ForeColor.RGB = 5287936 # Green
    btn.TextFrame.HorizontalAlignment = -4108 # xlHAlignCenter
    btn.TextFrame.VerticalAlignment = -4108 # xlVAlignCenter
    btn.OnAction = "SaveDailyPurchases"
    
    # Gridlines off
    excel.ActiveWindow.DisplayGridlines = False

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
