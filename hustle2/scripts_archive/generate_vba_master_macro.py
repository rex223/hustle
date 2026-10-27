'''
VBA Procedure: UpdateMasterExecutiveSummary
To be added to ThisWorkbook in Excel.
'''

VBA_CODE = '''
' =========================================================================
' MASTER EXECUTIVE FINANCIAL SUMMARY REFRESH ENGINE
' =========================================================================
Public Sub UpdateMasterExecutiveSummary()
    On Error GoTo ErrHandler
    
    Dim wsMaster As Worksheet
    On Error Resume Next
    Set wsMaster = ThisWorkbook.Sheets("MASTER_EXECUTIVE_SUMMARY")
    On Error GoTo ErrHandler
    
    If wsMaster Is Nothing Then
        MsgBox "MASTER_EXECUTIVE_SUMMARY sheet not found!", vbExclamation, "Executive Summary"
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    
    Dim wsDB As Worksheet, wsPurch As Worksheet, wsRentMisc As Worksheet
    Set wsDB = ThisWorkbook.Sheets("database_expenditure")
    
    On Error Resume Next
    Set wsPurch = ThisWorkbook.Sheets("DATABASE_PURCHASE")
    Set wsRentMisc = ThisWorkbook.Sheets("RENT PERSONS MISCELLANEOUS")
    On Error GoTo ErrHandler
    
    ' -------------------------------------------------------------------------
    ' 1. COLLECT RENT DATA FOR ALL MONTHS ACROSS TENANT SHEETS
    ' -------------------------------------------------------------------------
    Dim rentByMonth(1 To 12) As Double
    Dim m As Long
    For m = 1 To 12
        rentByMonth(m) = 0
    Next m
    
    Dim ws As Worksheet, lastR As Long, r As Long, pDate As Variant
    Dim forMonthStr As String, amt As Double, targetM As Long
    
    For Each ws In ThisWorkbook.Worksheets
        Select Case UCase(Trim(ws.Name))
            Case "TODAYS_RENT_DATA", "TODAYS_EXPENDITURE", "DATABASE_EXPENDITURE", _
                 "PURCHASE SHEET", "DATABASE_PURCHASE", "MISCELLANEOUS", _
                 "RENT PERSONS MISCELLANEOUS", "MASTER_EXECUTIVE_SUMMARY"
                ' Skip non-tenant sheets
            Case Else
                lastR = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
                If lastR >= 2 Then
                    For r = 2 To lastR
                        forMonthStr = ""
                        If Not IsError(ws.Cells(r, 3).Value) Then
                            forMonthStr = UCase(Trim(CStr(ws.Cells(r, 3).Value)))
                        End If
                        pDate = ws.Cells(r, 1).Value
                        
                        targetM = 0
                        If InStr(forMonthStr, "JUL") > 0 Then targetM = 7
                        If InStr(forMonthStr, "AUG") > 0 Then targetM = 8
                        If InStr(forMonthStr, "SEP") > 0 Then targetM = 9
                        If InStr(forMonthStr, "OCT") > 0 Then targetM = 10
                        If InStr(forMonthStr, "NOV") > 0 Then targetM = 11
                        If InStr(forMonthStr, "DEC") > 0 Then targetM = 12
                        If InStr(forMonthStr, "JAN") > 0 Then targetM = 1
                        If InStr(forMonthStr, "FEB") > 0 Then targetM = 2
                        If InStr(forMonthStr, "MAR") > 0 Then targetM = 3
                        If InStr(forMonthStr, "APR") > 0 Then targetM = 4
                        If InStr(forMonthStr, "MAY") > 0 Then targetM = 5
                        If InStr(forMonthStr, "JUN") > 0 Then targetM = 6
                        
                        If targetM = 0 And IsDate(pDate) Then
                            targetM = Month(CDate(pDate))
                        End If
                        
                        If targetM >= 1 And targetM <= 12 Then
                            If IsNumeric(ws.Cells(r, 6).Value) Then
                                rentByMonth(targetM) = rentByMonth(targetM) + CDbl(ws.Cells(r, 6).Value)
                            End If
                        End If
                    Next r
                End If
        End Select
    Next ws
    
    ' Rent Persons Misc Incoming
    If Not wsRentMisc Is Nothing Then
        lastR = wsRentMisc.Cells(wsRentMisc.Rows.Count, "B").End(xlUp).Row
        For r = 2 To lastR
            If UCase(Trim(wsRentMisc.Cells(r, 6).Value)) = "RECEIVED" Then
                If IsDate(wsRentMisc.Cells(r, 2).Value) Then
                    targetM = Month(CDate(wsRentMisc.Cells(r, 2).Value))
                    If targetM >= 1 And targetM <= 12 Then
                        If IsNumeric(wsRentMisc.Cells(r, 5).Value) Then
                            rentByMonth(targetM) = rentByMonth(targetM) + CDbl(wsRentMisc.Cells(r, 5).Value)
                        End If
                    End If
                End If
            End If
        Next r
    End If
    
    ' -------------------------------------------------------------------------
    ' 2. COLLECT PURCHASES BY MONTH
    ' -------------------------------------------------------------------------
    Dim purchByMonth(1 To 12) As Double
    For m = 1 To 12
        purchByMonth(m) = 0
    Next m
    
    If Not wsPurch Is Nothing Then
        Dim amtCol As Long: amtCol = 8
        Dim cIdx As Long
        For cIdx = 1 To 20
            If UCase(Trim(wsPurch.Cells(1, cIdx).Value)) = "AMOUNT" Or UCase(Trim(wsPurch.Cells(1, cIdx).Value)) = "TOTAL AMOUNT" Then
                amtCol = cIdx
                Exit For
            End If
        Next cIdx
        
        lastR = wsPurch.Cells(wsPurch.Rows.Count, "C").End(xlUp).Row
        For r = 2 To lastR
            If IsDate(wsPurch.Cells(r, 1).Value) Then
                ' Accounting month = 1 month prior to purchase date
                Dim acctDate As Date
                acctDate = DateAdd("m", -1, CDate(wsPurch.Cells(r, 1).Value))
                targetM = Month(acctDate)
                If targetM >= 1 And targetM <= 12 Then
                    If IsNumeric(wsPurch.Cells(r, amtCol).Value) Then
                        purchByMonth(targetM) = purchByMonth(targetM) + CDbl(wsPurch.Cells(r, amtCol).Value)
                    End If
                End If
            End If
        Next r
    End If
    
    ' Rent Persons Misc Outgoing
    If Not wsRentMisc Is Nothing Then
        lastR = wsRentMisc.Cells(wsRentMisc.Rows.Count, "H").End(xlUp).Row
        For r = 2 To lastR
            If IsDate(wsRentMisc.Cells(r, 8).Value) Then
                targetM = Month(CDate(wsRentMisc.Cells(r, 8).Value))
                If targetM >= 1 And targetM <= 12 Then
                    If IsNumeric(wsRentMisc.Cells(r, 11).Value) Then
                        purchByMonth(targetM) = purchByMonth(targetM) + CDbl(wsRentMisc.Cells(r, 11).Value)
                    End If
                End If
            End If
        Next r
    End If
    
    ' -------------------------------------------------------------------------
    ' 3. POPULATE MONTH ROWS IN MASTER EXECUTIVE SUMMARY
    ' -------------------------------------------------------------------------
    ' Rows 13 to 24 correspond to Months:
    ' 13: July 2026 (M=7)
    ' 14: August 2026 (M=8)
    ' 15: September 2026 (M=9)
    ' 16: October 2026 (M=10)
    ' 17: November 2026 (M=11)
    ' 18: December 2026 (M=12)
    ' 19: January 2027 (M=1)
    ' 20: February 2027 (M=2)
    ' 21: March 2027 (M=3)
    ' 22: April 2027 (M=4)
    ' 23: May 2027 (M=5)
    ' 24: June 2027 (M=6)
    
    Dim monthMap(13 To 24) As Long
    monthMap(13) = 7:  monthMap(14) = 8:  monthMap(15) = 9
    monthMap(16) = 10: monthMap(17) = 11: monthMap(18) = 12
    monthMap(19) = 1:  monthMap(20) = 2:  monthMap(21) = 3
    monthMap(22) = 4:  monthMap(23) = 5:  monthMap(24) = 6
    
    Dim rowIdx As Long
    For rowIdx = 13 To 24
        targetM = monthMap(rowIdx)
        If rentByMonth(targetM) > 0 Then
            wsMaster.Cells(rowIdx, 3).Value = rentByMonth(targetM)
        End If
        If purchByMonth(targetM) > 0 Then
            wsMaster.Cells(rowIdx, 4).Value = purchByMonth(targetM)
        End If
    Next rowIdx
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    
    wsMaster.Calculate
    MsgBox "Executive Master Summary successfully synchronized with all rental and purchase records!", vbInformation, "Executive Master Refresh"
    Exit Sub
    
ErrHandler:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    MsgBox "Error refreshing Master Executive Summary: " & Err.Description, vbCritical, "Refresh Error"
End Sub
'''

with open("vba_current/UpdateMasterExecutiveSummary.txt", "w", encoding="utf-8") as f:
    f.write(VBA_CODE)
print("Saved vba_current/UpdateMasterExecutiveSummary.txt")
