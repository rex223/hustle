import win32com.client
import os

excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False
file_path = os.path.abspath('RENTAL DATASHEET 1ST AUG.xlsm')

vba_code = """
Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    On Error GoTo ErrorHandler
    
    ' =========================================================================
    ' SECTION 1: LOGIC FOR "RENT DATA" SHEET
    ' =========================================================================
    If UCase(Sh.Name) = "RENT DATA" Then
        ' Trigger only if you edit columns A to G
        If Not Intersect(Target, Sh.Range("A2:G1000")) Is Nothing Then
            Application.EnableEvents = False
            Application.ScreenUpdating = False
            
            ' --- PART 1: AUTO-CREATE & UPDATE NAME SHEETS ---
            Dim cell As Range
            If Not Intersect(Target, Sh.Range("A2:A1000")) Is Nothing Then
                For Each cell In Intersect(Target, Sh.Range("A2:A1000"))
                    If Trim(cell.Value) <> "" Then
                        Dim baseName As String
                        baseName = Trim(cell.Value)
                        
                        ' Strip trailing numbers and spaces
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
                            
                            ' If sheet doesn't exist, create it!
                            If Not wsExists Then
                                Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
                                ws.Name = UCase(baseName)
                                
                                ' Build the header (Includes PAID AMOUNT)
                                ws.Range("A1:H1").Value = Array("DATE", "TIME", "FOR THE MONTH", "NAME", "AMOUNT", "PAID AMOUNT", "GROUP", "DUE")
                                ws.Range("A1:H1").Font.Bold = True
                                
                                ' Adjust widths
                                ws.Columns("A").ColumnWidth = 15
                                ws.Columns("B").ColumnWidth = 12
                                ws.Columns("C").ColumnWidth = 20
                                ws.Columns("D:H").ColumnWidth = 15
                            End If
                            
                            ' Set ws to point to our sheet
                            Set ws = ThisWorkbook.Sheets(UCase(baseName))
                            
                            ' ALWAYS ensure dynamic formulas are present in row 2
                            ws.Range("D2").Value = UCase(baseName)
                            ws.Range("E2").Formula = "=SUMIF('RENT DATA'!A:A, """ & UCase(baseName) & "*"", 'RENT DATA'!C:C)"
                            ws.Range("G2").Formula = "=VLOOKUP(""" & UCase(baseName) & "*"", 'RENT DATA'!A:E, 5, FALSE)"
                            ws.Range("H2").Formula = "=E2-F2"
                        End If
                    End If
                Next cell
            End If
            
            ' --- PART 2: AUTO-SORT AND CLEAN SUBTOTALS ---
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
            
            ' ---> THIS ENSURES FILTERS STAY ON FOR ALL COLUMNS A THROUGH M <---
            If Not Sh.AutoFilterMode Then
                Sh.Range("A1:M1").AutoFilter
            End If
            
            ' Update the Group Summary Table
            UpdateGroupSummary
            
            Application.ScreenUpdating = True
            Application.EnableEvents = True
        End If

    ' =========================================================================
    ' SECTION 2: LOGIC FOR INDIVIDUAL NAME SHEETS (The Smart Auto-Fill)
    ' =========================================================================
    Else
        Dim ignoreSheets As String
        ignoreSheets = "|PAYMENT SHEET|EXPENDITURE|PURCHASE|SHEET17|"
        
        If InStr(1, ignoreSheets, "|" & UCase(Sh.Name) & "|") > 0 Then Exit Sub
        
        ' Only proceed if a single cell was changed
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
                
                ' Update the Group Summary since a new payment row might have been generated
                UpdateGroupSummary
                
                Application.EnableEvents = True
            End If
        Else
            ' Even if it's not the last row, update the summary in case they modified an old payment
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

' =========================================================================
' SUBROUTINE: UPDATE GROUP SUMMARY ON RENT DATA
' =========================================================================
Private Sub UpdateGroupSummary()
    Dim wsRent As Worksheet
    Set wsRent = ThisWorkbook.Sheets("RENT DATA")
    
    Dim dictTotal As Object, dictPaid As Object
    Set dictTotal = CreateObject("Scripting.Dictionary")
    Set dictPaid = CreateObject("Scripting.Dictionary")
    
    ' 1. Calculate Group Totals from RENT DATA (Columns C and E)
    Dim lastRow As Long
    lastRow = wsRent.Cells(wsRent.Rows.Count, "E").End(xlUp).Row
    
    Dim i As Long
    Dim grp As String
    Dim amt As Double
    
    For i = 2 To lastRow
        grp = Trim(CStr(wsRent.Cells(i, 5).Value))
        ' Skip the subtotal/grand total rows
        If grp <> "" And UCase(grp) <> "GRAND TOTAL" And UCase(Left(grp, 5)) <> "TOTAL" Then
            amt = 0
            If IsNumeric(wsRent.Cells(i, 3).Value) Then amt = CDbl(wsRent.Cells(i, 3).Value)
            dictTotal(grp) = dictTotal(grp) + amt
            ' Initialize paid dictionary
            If Not dictPaid.Exists(grp) Then dictPaid(grp) = 0
        End If
    Next i
    
    ' 2. Calculate Group Paid from all individual alias sheets
    Dim ws As Worksheet
    Dim ignoreSheets As String
    ignoreSheets = "|RENT DATA|PAYMENT SHEET|EXPENDITURE|PURCHASE|SHEET17|"
    
    Dim r As Long
    Dim lastAliasRow As Long
    For Each ws In ThisWorkbook.Sheets
        If InStr(1, ignoreSheets, "|" & UCase(ws.Name) & "|") = 0 Then
            lastAliasRow = ws.Cells(ws.Rows.Count, "D").End(xlUp).Row
            If lastAliasRow >= 2 Then
                For r = 2 To lastAliasRow
                    grp = Trim(CStr(ws.Cells(r, 7).Value)) ' Group is Col G
                    If grp <> "" Then
                        amt = 0
                        ' Sum up their Paid Amounts (Col F)
                        If IsNumeric(ws.Cells(r, 6).Value) Then amt = CDbl(ws.Cells(r, 6).Value)
                        dictPaid(grp) = dictPaid(grp) + amt
                        ' Ensure the group exists in our total dictionary
                        If Not dictTotal.Exists(grp) Then dictTotal(grp) = 0
                    End If
                Next r
            End If
        End If
    Next ws
    
    ' 3. Write Summary to RENT DATA (Columns J, K, L, M)
    wsRent.Range("J1:M1000").ClearContents
    wsRent.Range("J1:M1").Value = Array("GROUP", "TOTAL AMOUNT", "TOTAL PAID", "GROUP DUE")
    wsRent.Range("J1:M1").Font.Bold = True
    
    Dim outRow As Long
    outRow = 2
    
    Dim key As Variant
    For Each key In dictTotal.Keys
        wsRent.Cells(outRow, 10).Value = key ' J
        wsRent.Cells(outRow, 11).Value = dictTotal(key) ' K
        wsRent.Cells(outRow, 12).Value = dictPaid(key) ' L
        wsRent.Cells(outRow, 13).Value = dictTotal(key) - dictPaid(key) ' M (Group Due)
        outRow = outRow + 1
    Next key
    
    wsRent.Columns("J:M").AutoFit
End Sub
"""

try:
    wb = excel.Workbooks.Open(file_path)
    workbook_module = wb.VBProject.VBComponents("ThisWorkbook").CodeModule
    
    if workbook_module.CountOfLines > 0:
        workbook_module.DeleteLines(1, workbook_module.CountOfLines)
        
    workbook_module.AddFromString(vba_code)
    
    # Let's forcefully trigger a calculation so it builds the summary immediately
    excel.EnableEvents = True
    # Modify a random out-of-bounds cell to trigger it if we want, but wait, 
    # we can just run the subroutine manually!
    # UpdateGroupSummary is Private, so we can't easily Application.Run it unless we make it Public.
    # Let's just make a dummy change in RENT DATA Z100
    ws = wb.Sheets("RENT DATA")
    ws.Range("Z100").Value = "Trigger"
    ws.Range("Z100").ClearContents()
    
    wb.Save()
    print("Code successfully injected and saved.")
except Exception as e:
    print(f"Error: {e}")
finally:
    try:
        wb.Close(False)
    except:
        pass
    excel.Quit()
