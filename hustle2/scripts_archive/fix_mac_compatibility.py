import win32com.client
import os

excel = win32com.client.Dispatch('Excel.Application')
excel.Visible = False
excel.DisplayAlerts = False
file_path = os.path.abspath('RENTAL DATASHEET 1ST AUG.xlsm')

mac_compatible_sub = '''Public Sub UpdateGroupSummary()
    Dim wsRent As Worksheet
    Set wsRent = ThisWorkbook.Sheets("RENT DATA")
    
    ' Mac-compatible storage (Arrays instead of Scripting.Dictionary)
    Dim grpNames() As String
    Dim grpTotals() As Double
    Dim grpPaids() As Double
    Dim grpDates() As Variant
    Dim grpCount As Long
    grpCount = 0
    
    Dim lastRow As Long
    lastRow = wsRent.Cells(wsRent.Rows.Count, "E").End(xlUp).Row
    
    Dim i As Long, j As Long
    Dim grp As String
    Dim amt As Double
    Dim cellVal As Variant
    Dim idx As Long
    
    ' 1. Read Rent Totals
    For i = 2 To lastRow
        cellVal = wsRent.Cells(i, 5).Value
        If Not IsError(cellVal) Then
            grp = Trim(CStr(cellVal))
            If grp <> "" And UCase(grp) <> "GRAND TOTAL" And UCase(Left(grp, 5)) <> "TOTAL" And UCase(Right(grp, 5)) <> "TOTAL" And grp <> "Error 2042" Then
                amt = 0
                If IsNumeric(wsRent.Cells(i, 3).Value) Then amt = CDbl(wsRent.Cells(i, 3).Value)
                
                ' Find or Add Group
                idx = -1
                If grpCount > 0 Then
                    For j = 1 To grpCount
                        If UCase(grpNames(j)) = UCase(grp) Then
                            idx = j
                            Exit For
                        End If
                    Next j
                End If
                
                If idx = -1 Then
                    grpCount = grpCount + 1
                    ReDim Preserve grpNames(1 To grpCount)
                    ReDim Preserve grpTotals(1 To grpCount)
                    ReDim Preserve grpPaids(1 To grpCount)
                    ReDim Preserve grpDates(1 To grpCount)
                    
                    grpNames(grpCount) = grp
                    grpTotals(grpCount) = 0
                    grpPaids(grpCount) = 0
                    grpDates(grpCount) = ""
                    idx = grpCount
                End If
                
                grpTotals(idx) = grpTotals(idx) + amt
            End If
        End If
    Next i
    
    ' 2. Read Payments from other sheets
    Dim ws As Worksheet
    Dim ignoreSheets As String
    ignoreSheets = "|RENT DATA|PAYMENT SHEET|EXPENDITURE|PURCHASE|DATABASE_PURCHASE|PURCHASE SHEET|SHEET17|"
    
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
                            
                            ' Find or Add Group
                            idx = -1
                            If grpCount > 0 Then
                                For j = 1 To grpCount
                                    If UCase(grpNames(j)) = UCase(grp) Then
                                        idx = j
                                        Exit For
                                    End If
                                Next j
                            End If
                            
                            If idx = -1 Then
                                grpCount = grpCount + 1
                                ReDim Preserve grpNames(1 To grpCount)
                                ReDim Preserve grpTotals(1 To grpCount)
                                ReDim Preserve grpPaids(1 To grpCount)
                                ReDim Preserve grpDates(1 To grpCount)
                                
                                grpNames(grpCount) = grp
                                grpTotals(grpCount) = 0
                                grpPaids(grpCount) = 0
                                grpDates(grpCount) = ""
                                idx = grpCount
                            End If
                            
                            grpPaids(idx) = grpPaids(idx) + amt
                            
                            pDate = ws.Cells(r, 1).Value
                            If IsDate(pDate) And amt > 0 Then
                                If grpDates(idx) = "" Then
                                    grpDates(idx) = pDate
                                ElseIf CDate(pDate) > CDate(grpDates(idx)) Then
                                    grpDates(idx) = pDate
                                End If
                            End If
                        End If
                    End If
                Next r
            End If
        End If
    Next ws
    
    ' 3. Output to RENT DATA
    wsRent.Range("J1:N1000").Clear
    wsRent.Range("J1:N1").Value = Array("GROUP", "LAST PAID", "TOTAL AMOUNT", "TOTAL PAID", "GROUP DUE")
    
    With wsRent.Range("J1:N1")
        .Interior.Color = RGB(41, 128, 185)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .HorizontalAlignment = -4108 ' xlCenter
        .VerticalAlignment = -4108 ' xlCenter
        .RowHeight = 25
    End With
    
    Dim outRow As Long
    outRow = 2
    
    If grpCount > 0 Then
        For i = 1 To grpCount
            wsRent.Cells(outRow, 10).Value = grpNames(i)
            
            If IsDate(grpDates(i)) Then
                wsRent.Cells(outRow, 11).Value = grpDates(i)
                wsRent.Cells(outRow, 11).NumberFormat = "dd-mmm-yyyy"
            Else
                wsRent.Cells(outRow, 11).Value = "-"
            End If
            
            wsRent.Cells(outRow, 12).Value = grpTotals(i)
            wsRent.Cells(outRow, 13).Value = grpPaids(i)
            wsRent.Cells(outRow, 14).Value = grpTotals(i) - grpPaids(i)
            
            If outRow Mod 2 = 0 Then
                wsRent.Range(wsRent.Cells(outRow, 10), wsRent.Cells(outRow, 14)).Interior.Color = RGB(245, 245, 245)
            End If
            
            outRow = outRow + 1
        Next i
    End If
    
    If outRow > 2 Then
        With wsRent.Range("J1:N" & outRow - 1).Borders
            .LineStyle = 1 ' xlContinuous
            .Color = RGB(180, 180, 180)
            .Weight = 2 ' xlThin
        End With
    End If
    
    wsRent.Range("L2:N" & outRow).NumberFormat = "#,##0"
    wsRent.Columns("J:N").AutoFit
End Sub'''

try:
    wb = excel.Workbooks.Open(file_path)
    module = wb.VBProject.VBComponents('ThisWorkbook').CodeModule
    
    code = module.Lines(1, module.CountOfLines)
    
    # We want to replace everything from "Public Sub UpdateGroupSummary()" to the end
    start_idx = code.find("Public Sub UpdateGroupSummary()")
    if start_idx != -1:
        new_code = code[:start_idx] + mac_compatible_sub
        module.DeleteLines(1, module.CountOfLines)
        module.AddFromString(new_code)
    
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
