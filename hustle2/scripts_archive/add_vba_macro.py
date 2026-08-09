import win32com.client
import os

try:
    try:
        excel = win32com.client.GetActiveObject("Excel.Application")
    except:
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        
    excel.DisplayAlerts = False
    file_path = os.path.abspath('RENTAL DATASHEET 1ST AUG.xlsm')
    
    wb = None
    for b in excel.Workbooks:
        if b.FullName.lower() == file_path.lower():
            wb = b
            break
            
    if not wb:
        wb = excel.Workbooks.Open(file_path)

    workbook_module = wb.VBProject.VBComponents("ThisWorkbook").CodeModule

    vba_code = """
Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    On Error GoTo ErrorHandler
    
    Dim ignoreSheets As String
    ignoreSheets = "|RENT DATA|PAYMENT SHEET|EXPENDITURE|PURCHASE|SHEET17|"
    
    If InStr(1, ignoreSheets, "|" & UCase(Sh.Name) & "|") > 0 Then Exit Sub
    
    ' Only proceed if a single cell was changed
    If Target.Cells.Count > 1 Then Exit Sub
    
    Dim lastRow As Long
    ' Use column D (NAME) to find the last active row
    lastRow = Sh.Cells(Sh.Rows.Count, "D").End(xlUp).Row
    
    ' If the user modifies anything in the current last row
    If Target.Row = lastRow And Target.Value <> "" Then
        Dim nextRow As Long
        nextRow = lastRow + 1
        
        ' ONLY populate if the next row is currently empty!
        ' This prevents overwriting if they edit the last row multiple times.
        If Sh.Cells(nextRow, 4).Value = "" Then
            Application.EnableEvents = False
            
            ' 1. Auto increment Date by 1 month
            If IsDate(Sh.Cells(lastRow, 1).Value) Then
                Sh.Cells(nextRow, 1).Value = DateAdd("m", 1, Sh.Cells(lastRow, 1).Value)
            End If
            
            ' 2. Copy NAME
            Sh.Cells(nextRow, 4).Value = Sh.Cells(lastRow, 4).Value
            
            ' 3. Copy AMOUNT formula
            If Sh.Cells(lastRow, 5).HasFormula Then
                Sh.Cells(lastRow, 5).Copy Destination:=Sh.Cells(nextRow, 5)
            End If
            
            ' 4. Copy GROUP formula
            If Sh.Cells(lastRow, 7).HasFormula Then
                Sh.Cells(lastRow, 7).Copy Destination:=Sh.Cells(nextRow, 7)
            End If
            
            ' 5. Copy DUE formula
            If Sh.Cells(lastRow, 8).HasFormula Then
                Sh.Cells(lastRow, 8).Copy Destination:=Sh.Cells(nextRow, 8)
            End If
            
            Application.EnableEvents = True
        End If
    End If
    Exit Sub
    
ErrorHandler:
    Application.EnableEvents = True
End Sub
"""
    if workbook_module.CountOfLines > 0:
        workbook_module.DeleteLines(1, workbook_module.CountOfLines)
    
    workbook_module.AddFromString(vba_code)
    print("VBA Code added successfully!")
        
    wb.Save()
    print("Workbook saved successfully!")
    
except Exception as e:
    print(f"Error: {e}")
finally:
    excel.DisplayAlerts = True
