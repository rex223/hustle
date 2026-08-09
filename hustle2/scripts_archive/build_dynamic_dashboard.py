import win32com.client
import os

excel = win32com.client.Dispatch('Excel.Application')
excel.Visible = False
excel.DisplayAlerts = False
file_path = os.path.abspath('RENTAL DATASHEET 1ST AUG.xlsm')

vba_ws = '''
Private Sub Worksheet_Change(ByVal Target As Range)
    On Error GoTo ErrorHandler
    
    Dim wsDB As Worksheet
    Dim tblEntry As ListObject
    Dim tblDB As ListObject
    Dim dateCell As Range
    
    Set dateCell = Me.Range("C1")
    
    On Error Resume Next
    Set tblEntry = Me.ListObjects("EntryFormTable")
    Set wsDB = ThisWorkbook.Sheets("PURCHASE SHEET DAILY")
    Set tblDB = wsDB.ListObjects("DailyPurchaseLedger")
    On Error GoTo ErrorHandler
    
    If Not Intersect(Target, dateCell) Is Nothing Then
        Application.EnableEvents = False
        Application.ScreenUpdating = False
        
        If tblEntry.DataBodyRange.Rows.Count > 0 Then
            tblEntry.ListColumns("QTY (KG)").DataBodyRange.ClearContents
            tblEntry.ListColumns("QTY (PCS)").DataBodyRange.ClearContents
            tblEntry.ListColumns("TRSP").DataBodyRange.ClearContents
        End If
        
        If IsDate(dateCell.Value) Then
            Dim searchDate As Date
            searchDate = CDate(dateCell.Value)
            
            Dim i As Long, j As Long
            Dim dbItem As String
            Dim entryItem As String
            
            If Not tblDB.DataBodyRange Is Nothing Then
                For i = 1 To tblDB.DataBodyRange.Rows.Count
                    If IsDate(tblDB.DataBodyRange.Cells(i, 1).Value) Then
                        If CDate(tblDB.DataBodyRange.Cells(i, 1).Value) = searchDate Then
                            dbItem = Trim(tblDB.DataBodyRange.Cells(i, 3).Value)
                            
                            For j = 1 To tblEntry.DataBodyRange.Rows.Count
                                entryItem = Trim(tblEntry.DataBodyRange.Cells(j, 2).Value)
                                If UCase(dbItem) = UCase(entryItem) Then
                                    tblEntry.DataBodyRange.Cells(j, 3).Value = tblDB.DataBodyRange.Cells(i, 4).Value
                                    tblEntry.DataBodyRange.Cells(j, 4).Value = tblDB.DataBodyRange.Cells(i, 5).Value
                                    tblEntry.DataBodyRange.Cells(j, 6).Value = tblDB.DataBodyRange.Cells(i, 7).Value
                                    Exit For
                                End If
                            Next j
                        End If
                    End If
                Next i
            End If
        End If
        
        Application.EnableEvents = True
        Application.ScreenUpdating = True
    End If

    Exit Sub
ErrorHandler:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub
'''

vba_mod = '''
Public Sub SaveDailyPurchases()
    Dim wsEntry As Worksheet
    Dim wsDB As Worksheet
    Dim tblEntry As ListObject
    Dim tblDB As ListObject
    Dim dateCell As Range
    
    On Error Resume Next
    Set wsEntry = ThisWorkbook.Sheets("TODAYS PURCHASES")
    Set wsDB = ThisWorkbook.Sheets("PURCHASE SHEET DAILY")
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
            itemsSaved = itemsSaved + 1
        End If
    Next i
    
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    
    MsgBox itemsSaved & " items successfully saved for " & Format(saveDate, "dd-mmm-yyyy") & "!", vbInformation, "Success"
End Sub
'''

try:
    wb = excel.Workbooks.Open(file_path)
    
    # 1. Fix "CHICHEN" to "CHICKEN"
    for s_name in ['PURCHASE SHEET DAILY', 'TODAYS PURCHASES']:
        ws = wb.Sheets(s_name)
        # Using Replace
        # What:="CHICHEN", Replacement:="CHICKEN"
        ws.Cells.Replace("CHICHEN", "CHICKEN", 2, 1, False, False)
        ws.Cells.Replace("chichen", "CHICKEN", 2, 1, False, False)

    ws_entry = wb.Sheets("TODAYS PURCHASES")
    
    # 2. Add Date Selector
    ws_entry.Range("B1").Value = "SELECT DATE:"
    ws_entry.Range("B1").Font.Bold = True
    ws_entry.Range("B1").Font.Size = 12
    ws_entry.Range("B1").HorizontalAlignment = -4152 # xlRight
    
    ws_entry.Range("C1").Value = "05-Aug-2026" # Default date to start
    ws_entry.Range("C1").Font.Bold = True
    ws_entry.Range("C1").Font.Size = 12
    ws_entry.Range("C1").Interior.Color = 10092543 # Light Yellow
    ws_entry.Range("C1").Borders.LineStyle = 1 # Solid
    
    # 3. Inject macros
    # Sheet macro
    ws_module = wb.VBProject.VBComponents(ws_entry.CodeName).CodeModule
    if ws_module.CountOfLines > 0:
        ws_module.DeleteLines(1, ws_module.CountOfLines)
    ws_module.AddFromString(vba_ws)
    
    # Standard macro
    for comp in wb.VBProject.VBComponents:
        if comp.Name == "ModDailyPurchases":
            comp.CodeModule.DeleteLines(1, comp.CodeModule.CountOfLines)
            comp.CodeModule.AddFromString(vba_mod)
            break
            
    # Trigger the event to load the default date
    excel.EnableEvents = True
    ws_entry.Range("C1").Value = "05-Aug-2026"

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
