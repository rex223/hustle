import win32com.client
import os

excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False
file_path = os.path.abspath('RENTAL DATASHEET 1ST AUG.xlsm')

vba_code = """
Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    On Error GoTo ErrorHandler
    
    If UCase(Sh.Name) = "RENT DATA" Then
        If Not Intersect(Target, Sh.Range("A2:G1000")) Is Nothing Then
            Application.EnableEvents = False
            Application.ScreenUpdating = False
            
            Dim cell As Range
            If Not Intersect(Target, Sh.Range("A2:A1000")) Is Nothing Then
                For Each cell In Intersect(Target, Sh.Range("A2:A1000"))
                    If Trim(cell.Value) <> "" Then
                        Dim baseName As String
                        baseName = Trim(cell.Value)
                        
                        Dim prevLen As Long
                        Do
                            prevLen = Len(baseName)
                            If Right(baseName, 1) Like "#" Then baseName = Left(baseName, Len(baseName) - 1)
                            baseName = Trim(baseName)
                        Loop Until Len(baseName) = prevLen Or Len(baseName) = 0
                        
                        If baseName <> "" And UCase(baseName) <> "GRAND TOTAL" And UCase(Left(baseName, 5)) <> "TOTAL" Then
                            Dim ws As Worksheet
                            Dim wsExists As Boolean
                            wsExists = False
                            
                            For Each ws In ThisWorkbook.Sheets
                                If UCase(ws.Name) = UCase(baseName) Then
                                    wsExists = True
                                    Exit For
                                End If
                            Next ws
                            
                            If Not wsExists Then
                                Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
                                ws.Name = UCase(baseName)
                                
                                ws.Range("A1:H1").Value = Array("DATE", "TIME", "FOR THE MONTH", "NAME", "AMOUNT", "PAID AMOUNT", "GROUP", "DUE")
                                ws.Range("A1:H1").Font.Bold = True
                                
                                ws.Columns("A").ColumnWidth = 15
                                ws.Columns("B").ColumnWidth = 12
                                ws.Columns("C").ColumnWidth = 20
                                ws.Columns("D:H").ColumnWidth = 15
                            End If
                            
                            Set ws = ThisWorkbook.Sheets(UCase(baseName))
                            
                            ws.Range("D2").Value = UCase(baseName)
                            ws.Range("E2").Formula = "=SUMIF('RENT DATA'!A:A, """ & UCase(baseName) & "*"", 'RENT DATA'!C:C)"
                            ws.Range("G2").Formula = "=VLOOKUP(""" & UCase(baseName) & "*"", 'RENT DATA'!A:E, 5, FALSE)"
                            ws.Range("H2").Formula = "=E2-F2"
                        End If
                    End If
                Next cell
            End If
            
            Sh.Activate
            
            Sh.Range("A2:G1000").Interior.ColorIndex = xlNone
            Sh.Range("A2:G1000").Font.Bold = False
            
            On Error Resume Next
            Sh.Range("A1").CurrentRegion.RemoveSubtotal
            On Error GoTo ErrorHandler
            
            Dim i As Long, lastDataRow As Long
            lastDataRow = Sh.Cells(Sh.Rows.Count, "C").End(xlUp).Row
            If lastDataRow >= 2 Then
                For i = lastDataRow To 2 Step -1
                    If Left(UCase(Trim(Sh.Cells(i, 1).Value)), 5) = "TOTAL" Or UCase(Trim(Sh.Cells(i, 5).Value)) = "GRAND TOTAL" Then
                        Sh.Rows(i).Delete
                    End If
                Next i
            End If
            
            lastDataRow = Sh.Cells(Sh.Rows.Count, "A").End(xlUp).Row
            If lastDataRow > 1 Then
                With Sh.Sort
                    .SortFields.Clear
                    .SortFields.Add Key:=Sh.Range("E2:E" & lastDataRow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
                    .SetRange Sh.Range("A1:G" & lastDataRow)
                    .Header = xlYes
                    .Apply
                End With
                
                Sh.Range("A1:G" & lastDataRow).Subtotal GroupBy:=5, Function:=xlSum, TotalList:=Array(3), Replace:=True, PageBreaks:=False, SummaryBelowData:=True
                    
                ActiveWindow.DisplayOutline = False
                
                Dim finalRow As Long
                finalRow = Sh.Cells(Sh.Rows.Count, "C").End(xlUp).Row
                For i = 2 To finalRow
                    If Right(Trim(Sh.Cells(i, 5).Value), 5) = "Total" Then
                        Sh.Range(Sh.Cells(i, 1), Sh.Cells(i, 7)).Interior.Color = RGB(237, 237, 237)
                        Sh.Range(Sh.Cells(i, 1), Sh.Cells(i, 7)).Font.Bold = True
                    End If
                Next i
            End If
            
            If Not Sh.AutoFilterMode Then
                Sh.Range("A1:N1").AutoFilter
            End If
            
            UpdateGroupSummary
            
            Application.ScreenUpdating = True
            Application.EnableEvents = True
        End If

    Else
        Dim ignoreSheets As String
        ignoreSheets = "|PAYMENT SHEET|EXPENDITURE|PURCHASE|SHEET17|"
        
        If InStr(1, ignoreSheets, "|" & UCase(Sh.Name) & "|") > 0 Then Exit Sub
        
        If Target.Cells.Count > 1 Then Exit Sub
        
        Dim lastAliasRow As Long
        lastAliasRow = Sh.Cells(Sh.Rows.Count, "D").End(xlUp).Row
        
        If Target.Row = lastAliasRow And Target.Value <> "" Then
            Dim nextRow As Long
            nextRow = lastAliasRow + 1
            
            If Sh.Cells(nextRow, 4).Value = "" Then
                Application.EnableEvents = False
                
                If IsDate(Sh.Cells(lastAliasRow, 1).Value) Then
                    Sh.Cells(nextRow, 1).Value = DateAdd("m", 1, Sh.Cells(lastAliasRow, 1).Value)
                End If
                
                Sh.Cells(nextRow, 4).Value = Sh.Cells(lastAliasRow, 4).Value
                
                If Sh.Cells(lastAliasRow, 5).HasFormula Then Sh.Cells(lastAliasRow, 5).Copy Destination:=Sh.Cells(nextRow, 5)
                If Sh.Cells(lastAliasRow, 7).HasFormula Then Sh.Cells(lastAliasRow, 7).Copy Destination:=Sh.Cells(nextRow, 7)
                If Sh.Cells(lastAliasRow, 8).HasFormula Then Sh.Cells(lastAliasRow, 8).Copy Destination:=Sh.Cells(nextRow, 8)
                
                UpdateGroupSummary
                
                Application.EnableEvents = True
            End If
        Else
            Application.EnableEvents = False
            UpdateGroupSummary
            Application.EnableEvents = True
        End If
    End If
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
End Sub

Public Sub UpdateGroupSummary()
    Dim wsRent As Worksheet
    Set wsRent = ThisWorkbook.Sheets("RENT DATA")
    
    Dim dictTotal As Object, dictPaid As Object, dictDate As Object
    Set dictTotal = CreateObject("Scripting.Dictionary")
    Set dictPaid = CreateObject("Scripting.Dictionary")
    Set dictDate = CreateObject("Scripting.Dictionary")
    
    Dim lastRow As Long
    lastRow = wsRent.Cells(wsRent.Rows.Count, "E").End(xlUp).Row
    
    Dim i As Long
    Dim grp As String
    Dim amt As Double
    Dim cellVal As Variant
    
    For i = 2 To lastRow
        cellVal = wsRent.Cells(i, 5).Value
        If Not IsError(cellVal) Then
            grp = Trim(CStr(cellVal))
            If grp <> "" And UCase(grp) <> "GRAND TOTAL" And UCase(Left(grp, 5)) <> "TOTAL" And UCase(Right(grp, 5)) <> "TOTAL" And grp <> "Error 2042" Then
                amt = 0
                If IsNumeric(wsRent.Cells(i, 3).Value) Then amt = CDbl(wsRent.Cells(i, 3).Value)
                
                If Not dictTotal.Exists(grp) Then dictTotal(grp) = 0
                dictTotal(grp) = dictTotal(grp) + amt
                
                If Not dictPaid.Exists(grp) Then dictPaid(grp) = 0
                If Not dictDate.Exists(grp) Then dictDate(grp) = ""
            End If
        End If
    Next i
    
    Dim ws As Worksheet
    Dim ignoreSheets As String
    ignoreSheets = "|RENT DATA|PAYMENT SHEET|EXPENDITURE|PURCHASE|SHEET17|"
    
    Dim r As Long
    Dim lastAliasRow As Long
    Dim pDate As Variant
    
    For Each ws In ThisWorkbook.Sheets
        If InStr(1, ignoreSheets, "|" & UCase(ws.Name) & "|") = 0 Then
            lastAliasRow = ws.Cells(ws.Rows.Count, "D").End(xlUp).Row
            If lastAliasRow >= 2 Then
                For r = 2 To lastAliasRow
                    cellVal = ws.Cells(r, 7).Value
                    If Not IsError(cellVal) Then
                        grp = Trim(CStr(cellVal))
                        If grp <> "" And UCase(Right(grp, 5)) <> "TOTAL" Then
                            amt = 0
                            If IsNumeric(ws.Cells(r, 6).Value) Then amt = CDbl(ws.Cells(r, 6).Value)
                            
                            If Not dictPaid.Exists(grp) Then dictPaid(grp) = 0
                            dictPaid(grp) = dictPaid(grp) + amt
                            
                            If Not dictDate.Exists(grp) Then dictDate(grp) = ""
                            pDate = ws.Cells(r, 1).Value
                            If IsDate(pDate) And amt > 0 Then
                                If dictDate(grp) = "" Then
                                    dictDate(grp) = pDate
                                ElseIf CDate(pDate) > CDate(dictDate(grp)) Then
                                    dictDate(grp) = pDate
                                End If
                            End If
                            
                            If Not dictTotal.Exists(grp) Then dictTotal(grp) = 0
                        End If
                    End If
                Next r
            End If
        End If
    Next ws
    
    wsRent.Range("J1:N1000").Clear
    wsRent.Range("J1:N1").Value = Array("GROUP", "LAST PAID", "TOTAL AMOUNT", "TOTAL PAID", "GROUP DUE")
    
    With wsRent.Range("J1:N1")
        .Interior.Color = RGB(41, 128, 185)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .RowHeight = 25
    End With
    
    Dim outRow As Long
    outRow = 2
    
    Dim key As Variant
    For Each key In dictTotal.Keys
        wsRent.Cells(outRow, 10).Value = key
        
        If IsDate(dictDate(key)) Then
            wsRent.Cells(outRow, 11).Value = dictDate(key)
            wsRent.Cells(outRow, 11).NumberFormat = "dd-mmm-yyyy"
        Else
            wsRent.Cells(outRow, 11).Value = "-"
        End If
        
        wsRent.Cells(outRow, 12).Value = dictTotal(key)
        wsRent.Cells(outRow, 13).Value = dictPaid(key)
        wsRent.Cells(outRow, 14).Value = dictTotal(key) - dictPaid(key)
        
        If outRow Mod 2 = 0 Then
            wsRent.Range(wsRent.Cells(outRow, 10), wsRent.Cells(outRow, 14)).Interior.Color = RGB(245, 245, 245)
        End If
        
        outRow = outRow + 1
    Next key
    
    If outRow > 2 Then
        With wsRent.Range("J1:N" & outRow - 1).Borders
            .LineStyle = xlContinuous
            .Color = RGB(180, 180, 180)
            .Weight = xlThin
        End With
    End If
    
    wsRent.Range("L2:N" & outRow).NumberFormat = "#,##0"
    wsRent.Columns("J:N").AutoFit
End Sub
"""

try:
    wb = excel.Workbooks.Open(file_path)
    workbook_module = wb.VBProject.VBComponents("ThisWorkbook").CodeModule
    
    if workbook_module.CountOfLines > 0:
        workbook_module.DeleteLines(1, workbook_module.CountOfLines)
        
    workbook_module.AddFromString(vba_code)
    
    excel.EnableEvents = True
    excel.Application.Run("ThisWorkbook.UpdateGroupSummary")
    
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
