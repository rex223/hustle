import win32com.client
import os

excel = win32com.client.Dispatch('Excel.Application')
excel.Visible = False
excel.DisplayAlerts = False
file_path = os.path.abspath('RENTAL DATASHEET 12TH AUG.xlsm')

vba_macro = '''Public Sub SaveDailyPurchases()
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
            r.Range(1, 9).Value = tblEntry.DataBodyRange.Cells(i, 8).Value
            itemsSaved = itemsSaved + 1
        End If
    Next i
    
    ' SORT ASCENDING
    If Not tblDB.DataBodyRange Is Nothing Then
        With tblDB.Sort
            .SortFields.Clear
            .SortFields.Add Key:=tblDB.ListColumns(1).Range, SortOn:=0, Order:=1 ' xlAscending
            .Header = 1 ' xlYes
            .Apply
        End With
        
        ' REDRAW BORDERS
        tblDB.DataBodyRange.Borders(9).LineStyle = -4142 ' xlBottom xlNone
        
        For i = 1 To tblDB.DataBodyRange.Rows.Count - 1
            If tblDB.DataBodyRange.Cells(i, 1).Value <> tblDB.DataBodyRange.Cells(i + 1, 1).Value Then
                With tblDB.DataBodyRange.Rows(i).Borders(9) ' xlEdgeBottom
                    .LineStyle = 1 ' xlContinuous
                    .Weight = -4138 ' xlMedium
                    .Color = 0 ' Black
                End With
            End If
        Next i
    End If
    
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    
    MsgBox itemsSaved & " items successfully saved for " & Format(saveDate, "d mmm yyyy") & "!", vbInformation, "Success"
End Sub
'''

try:
    wb = excel.Workbooks.Open(file_path)
    
    for comp in wb.VBProject.VBComponents:
        if comp.Name == 'ModDailyPurchases':
            comp.CodeModule.DeleteLines(1, comp.CodeModule.CountOfLines)
            comp.CodeModule.AddFromString(vba_macro)
            break
            
    wb.Save()
    print("SUCCESS")
except Exception as e:
    import traceback
    traceback.print_exc()
finally:
    try:
        wb.Close(False)
    except:
        pass
    excel.Quit()
