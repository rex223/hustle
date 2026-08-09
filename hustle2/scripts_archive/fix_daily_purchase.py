import win32com.client
import os

excel = win32com.client.Dispatch('Excel.Application')
excel.Visible = False
excel.DisplayAlerts = False
file_path = os.path.abspath('RENTAL DATASHEET 1ST AUG.xlsm')

vba_code = '''
Private Sub Worksheet_Change(ByVal Target As Range)
    On Error GoTo ErrorHandler
    If Target.Cells.Count > 1 Then Exit Sub
    
    Dim tbl As ListObject
    On Error Resume Next
    Set tbl = Me.ListObjects("DailyPurchaseLedger")
    On Error GoTo ErrorHandler
    
    If tbl Is Nothing Then Exit Sub
    
    If Not Intersect(Target, tbl.DataBodyRange) Is Nothing Then
        If Target.Column = tbl.ListColumns("ITEM NAME").Range.Column Or Target.Column = tbl.ListColumns("GROUP / FOR").Range.Column Then
            Dim rowIdx As Long
            rowIdx = Target.Row - tbl.HeaderRowRange.Row
            
            If rowIdx > 1 Then
                Dim dateCell As Range
                Set dateCell = tbl.DataBodyRange.Cells(rowIdx, 1)
                Dim prevDateCell As Range
                Set prevDateCell = tbl.DataBodyRange.Cells(rowIdx - 1, 1)
                
                If IsEmpty(dateCell.Value) Or dateCell.Value = "" Then
                    Application.EnableEvents = False
                    dateCell.Value = prevDateCell.Value
                    Application.EnableEvents = True
                End If
            End If
        End If
    End If
    Exit Sub
ErrorHandler:
    Application.EnableEvents = True
End Sub

Public Sub ApplyConditionalFormatting()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("PURCHASE SHEET DAILY")
    Dim tbl As ListObject
    Set tbl = ws.ListObjects("DailyPurchaseLedger")
    
    Dim colA As Range
    Set colA = tbl.ListColumns(1).DataBodyRange
    
    colA.FormatConditions.Delete
    
    Dim fc As FormatCondition
    Set fc = colA.FormatConditions.Add(Type:=xlExpression, Formula1:="=$A2=$A1")
    fc.NumberFormat = ";;;"
    fc.StopIfTrue = False
End Sub
'''

try:
    wb = excel.Workbooks.Open(file_path)
    ws = wb.Sheets('PURCHASE SHEET DAILY')
    
    ws_module = wb.VBProject.VBComponents(ws.CodeName).CodeModule
    if ws_module.CountOfLines > 0:
        ws_module.DeleteLines(1, ws_module.CountOfLines)
    ws_module.AddFromString(vba_code)
    
    excel.Application.Run(f"'{wb.Name}'!{ws.CodeName}.ApplyConditionalFormatting")
    
    wb.Save()
    print('SUCCESS')
except Exception as e:
    print(f'Error: {e}')
finally:
    try:
        wb.Close(False)
    except:
        pass
    excel.Quit()
