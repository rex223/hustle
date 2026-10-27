import win32com.client
import os
import traceback

excel = win32com.client.Dispatch('Excel.Application')
excel.Visible = False
excel.DisplayAlerts = False
file_path = os.path.abspath('RENTAL DATASHEET 12TH AUG.xlsm')

vba_mod_purchases = '''Public Sub SaveDailyPurchases()
    Dim wsEntry As Worksheet
    Dim wsDB As Worksheet
    Dim tblEntry As ListObject
    Dim tblDB As ListObject
    Dim dateCell As Range
    
    On Error Resume Next
    Set wsEntry = ThisWorkbook.Sheets("PURCHASE SHEET")
    Set wsDB = ThisWorkbook.Sheets("DATABASE_PURCHASE")
    Set tblEntry = wsEntry.ListObjects("EntryFormTable")
    Set tblDB = wsDB.ListObjects("DailyPurchaseLedger")
    Set dateCell = wsEntry.Range("C1")
    On Error GoTo 0
    
    If wsEntry Is Nothing Or wsDB Is Nothing Then
        MsgBox "Error: Sheets not found!", vbCritical
        Exit Sub
    End If
    
    If Not IsDate(dateCell.Value) Then
        MsgBox "Please select a valid date at the top before saving!", vbExclamation
        Exit Sub
    End If
    
    Dim saveDate As Date
    saveDate = CDate(dateCell.Value)
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    Dim i As Long
    If Not tblDB.DataBodyRange Is Nothing Then
        For i = tblDB.DataBodyRange.Rows.Count To 1 Step -1
            If IsDate(tblDB.DataBodyRange.Cells(i, 1).Value) Then
                If CDate(tblDB.DataBodyRange.Cells(i, 1).Value) = saveDate Then
                    tblDB.ListRows(i).Delete
                End If
            End If
        Next i
    End If
    
    Dim qtyKg As Variant, qtyPcs As Variant, trsp As Variant
    Dim r As ListRow
    Dim itemsSaved As Long
    itemsSaved = 0
    
    For i = 1 To tblEntry.DataBodyRange.Rows.Count
        qtyKg = tblEntry.DataBodyRange.Cells(i, 3).Value
        qtyPcs = tblEntry.DataBodyRange.Cells(i, 4).Value
        trsp = tblEntry.DataBodyRange.Cells(i, 6).Value
        
        If (IsNumeric(qtyKg) And qtyKg > 0) Or (IsNumeric(qtyPcs) And qtyPcs > 0) Or (IsNumeric(trsp) And trsp > 0) Then
            Set r = tblDB.ListRows.Add
            r.Range(1, 1).Value = saveDate
            r.Range(1, 2).Value = tblEntry.DataBodyRange.Cells(i, 1).Value
            r.Range(1, 3).Value = tblEntry.DataBodyRange.Cells(i, 2).Value
            r.Range(1, 4).Value = qtyKg
            r.Range(1, 5).Value = qtyPcs
            r.Range(1, 6).Value = tblEntry.DataBodyRange.Cells(i, 5).Value
            r.Range(1, 7).Value = trsp
            r.Range(1, 8).Value = tblEntry.DataBodyRange.Cells(i, 7).Value
            r.Range(1, 9).Value = tblEntry.DataBodyRange.Cells(i, 8).Value ' PAID COLUMN
            itemsSaved = itemsSaved + 1
        End If
    Next i
    
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    
    MsgBox itemsSaved & " items successfully saved for " & Format(saveDate, "dd-mmm-yyyy") & "!", vbInformation, "Success"
End Sub
'''

vba_mod_expenditure_startnew = '''Sub StartNewMonth()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("todays_expenditure")
    
    Dim newMonth As String
    newMonth = InputBox("Enter the marker for the NEW month (e.g., AUGUST_2026):", "New Month")
    If newMonth = "" Then Exit Sub
    If Not UCase(Trim(newMonth)) Like "*_20##" Then
        MsgBox "Marker must end with an underscore and year (e.g., _2026)", vbCritical
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    ws.Range("A2").Value = UCase(Trim(newMonth))
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    
    If lastRow > 2 Then
        On Error Resume Next
        Dim c As Long
        Dim header As String
        For c = 4 To 100
            header = UCase(Trim(ws.Cells(1, c).Value))
            If header = "" Then header = UCase(Trim(ws.Cells(2, c).Value))
            If header = "" Then header = UCase(Trim(ws.Cells(3, c).Value))
            
            If header = "DATE" Or header = "MONTH" Or header = "AMOUNT" Or header = "PAID" Then
                ws.Range(ws.Cells(3, c), ws.Cells(lastRow, c)).ClearContents
            End If
        Next c
        On Error GoTo 0
    End If
    
    Application.ScreenUpdating = True
    MsgBox "Sheet is cleared and ready for " & newMonth & "! The fixed salaries/details have been preserved.", vbInformation, "Ready"
End Sub
'''

try:
    wb = excel.Workbooks.Open(file_path)
    
    # 1. Purchase Sheets Modifications
    for s_name in ['PURCHASE SHEET', 'DATABASE_PURCHASE']:
        ws = wb.Sheets(s_name)
        # Find TOTAL AMOUNT column
        for c in range(1, 20):
            val = str(ws.Cells(1, c).Value or ws.Cells(2, c).Value or ws.Cells(3, c).Value).upper().strip()
            if val == "TOTAL AMOUNT":
                # Insert after this column
                ws.Columns(c + 1).Insert(1) # xlShiftToRight
                header_row = 1 if ws.Cells(1, c).Value else (2 if ws.Cells(2, c).Value else 3)
                ws.Cells(header_row, c + 1).Value = "PAID"
                ws.Cells(header_row, c + 1).Interior.Color = 10092543 # Highlight it lightly so they know it's input
                ws.Columns(c + 1).ColumnWidth = 12
                break

    # 2. Expenditure Sheets Modifications
    for s_name in ['todays_expenditure', 'database_expenditure']:
        ws = wb.Sheets(s_name)
        # Loop backwards so insertions don't mess up our index
        for c in range(100, 0, -1):
            val = str(ws.Cells(1, c).Value or ws.Cells(2, c).Value or ws.Cells(3, c).Value).upper().strip()
            if val == "AMOUNT":
                ws.Columns(c + 1).Insert(1)
                header_row = 1 if str(ws.Cells(1, c).Value).upper().strip() == "AMOUNT" else (2 if str(ws.Cells(2, c).Value).upper().strip() == "AMOUNT" else 3)
                ws.Cells(header_row, c + 1).Value = "PAID"
                ws.Cells(header_row, c + 1).Interior.Color = 10092543
                ws.Columns(c + 1).ColumnWidth = 12

    # 3. Update ModDailyPurchases
    for comp in wb.VBProject.VBComponents:
        if comp.Name == 'ModDailyPurchases':
            comp.CodeModule.DeleteLines(1, comp.CodeModule.CountOfLines)
            comp.CodeModule.AddFromString(vba_mod_purchases)
            break
            
    # 4. Update Module1 StartNewMonth macro to be dynamic
    for comp in wb.VBProject.VBComponents:
        if comp.Name == 'Module1':
            code = comp.CodeModule.Lines(1, comp.CodeModule.CountOfLines)
            start_idx = code.find("Sub StartNewMonth()")
            if start_idx != -1:
                new_code = code[:start_idx] + vba_mod_expenditure_startnew
                comp.CodeModule.DeleteLines(1, comp.CodeModule.CountOfLines)
                comp.CodeModule.AddFromString(new_code)
            break

    wb.Save()
    print("SUCCESS")
except Exception as e:
    print(f"Error: {e}")
    traceback.print_exc()
finally:
    try:
        wb.Close(False)
    except:
        pass
    excel.Quit()
