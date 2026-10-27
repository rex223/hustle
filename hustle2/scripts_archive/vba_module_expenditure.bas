Public Const EXP_FIRST As Long = 3
Public Const EXP_LAST  As Long = 24
Public Const EXP_COLS  As Long = 27

Public Sub SaveMonthlyExpenditures()
    Dim wsEntry As Worksheet, wsDB As Worksheet
    Set wsEntry = ThisWorkbook.Sheets("TODAYS EXPENDITURE")
    Set wsDB    = ThisWorkbook.Sheets("DATABASE_EXPENDITURE")

    Dim mLabel As String
    mLabel = Trim(CStr(wsEntry.Range("C1").Value))
    If mLabel = "" Then
        MsgBox "Select a month label in C1 first (e.g. Aug-2026).", vbExclamation
        Exit Sub
    End If

    Dim ans As Integer
    ans = MsgBox("Save data for " & Chr(34) & mLabel & Chr(34) & "?" & Chr(10) & _
                 "Any existing entry for this month will be replaced.", _
                 vbYesNo + vbQuestion, "Confirm Save")
    If ans <> vbYes Then Exit Sub

    Application.ScreenUpdating = False
    Application.EnableEvents   = False

    Dim lastDB As Long, r As Long, cv As String
    lastDB = wsDB.Cells(wsDB.Rows.Count, 1).End(xlUp).Row

    For r = lastDB To 2 Step -1
        cv = Trim(CStr(wsDB.Cells(r, 1).Value))
        If cv = mLabel Or cv = "--- " & mLabel & " ---" Then
            wsDB.Rows(r).Delete
        End If
    Next r

    lastDB = wsDB.Cells(wsDB.Rows.Count, 1).End(xlUp).Row + 1
    If lastDB < 2 Then lastDB = 2

    With wsDB.Range(wsDB.Cells(lastDB, 1), wsDB.Cells(lastDB, EXP_COLS + 1))
        .Merge
        .Value = "--- " & mLabel & " ---"
        .Font.Bold = True
        .Font.Size = 11
        .Interior.Color = RGB(73, 80, 87)
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment   = xlCenter
    End With
    wsDB.Rows(lastDB).RowHeight = 20
    lastDB = lastDB + 1

    Dim eRow As Long, c As Long, dbRow As Long
    dbRow = lastDB
    For eRow = EXP_FIRST To EXP_LAST
        wsDB.Cells(dbRow, 1).Value = mLabel
        wsDB.Cells(dbRow, 1).Interior.Color = RGB(220, 220, 220)
        wsDB.Cells(dbRow, 1).Borders.LineStyle = xlContinuous
        For c = 1 To EXP_COLS
            wsDB.Cells(dbRow, c + 1).Value = wsEntry.Cells(eRow, c).Value
            wsDB.Cells(dbRow, c + 1).Borders.LineStyle = xlContinuous
        Next c
        dbRow = dbRow + 1
    Next eRow

    wsDB.Columns("A:AB").AutoFit

    Dim ic As Variant
    Dim inputCols As Variant
    inputCols = Array(4, 5, 7, 9, 11, 13, 15, 16, 18, 19, 21, 22, 25)
    For Each ic In inputCols
        wsEntry.Range(wsEntry.Cells(EXP_FIRST, ic), wsEntry.Cells(EXP_LAST, ic)).ClearContents
    Next ic

    Application.ScreenUpdating = True
    Application.EnableEvents   = True
    MsgBox "Saved for " & mLabel & "!", vbInformation, "Done"
End Sub

Public Sub LoadMonthlyExpenditures()
    Dim wsEntry As Worksheet, wsDB As Worksheet
    Set wsEntry = ThisWorkbook.Sheets("TODAYS EXPENDITURE")
    Set wsDB    = ThisWorkbook.Sheets("DATABASE_EXPENDITURE")

    Dim mLabel As String
    mLabel = Trim(CStr(wsEntry.Range("C1").Value))
    If mLabel = "" Then
        MsgBox "Select a month from the C1 dropdown first.", vbExclamation
        Exit Sub
    End If

    Dim lastDB As Long, r As Long, startRow As Long
    lastDB = wsDB.Cells(wsDB.Rows.Count, 1).End(xlUp).Row
    startRow = 0
    For r = 2 To lastDB
        If Trim(CStr(wsDB.Cells(r, 1).Value)) = mLabel Then
            startRow = r: Exit For
        End If
    Next r

    Application.ScreenUpdating = False
    Application.EnableEvents   = False

    Dim inputCols As Variant
    inputCols = Array(4, 5, 7, 9, 11, 13, 15, 16, 18, 19, 21, 22, 25)
    Dim ic As Variant
    For Each ic In inputCols
        wsEntry.Range(wsEntry.Cells(EXP_FIRST, ic), wsEntry.Cells(EXP_LAST, ic)).ClearContents
    Next ic

    If startRow = 0 Then
        Application.ScreenUpdating = True
        Application.EnableEvents   = True
        MsgBox "Prepared new empty sheet for " & Chr(34) & mLabel & Chr(34) & ".", vbInformation, "New Month"
        Exit Sub
    End If

    Dim eRow As Long, dbRow As Long, c As Long
    dbRow = startRow
    For eRow = EXP_FIRST To EXP_LAST
        If dbRow > lastDB Then Exit For
        For Each ic In inputCols
            c = CLng(ic)
            wsEntry.Cells(eRow, c).Value = wsDB.Cells(dbRow, c + 1).Value
        Next ic
        dbRow = dbRow + 1
    Next eRow

    Application.ScreenUpdating = True
    Application.EnableEvents   = True
    MsgBox "Loaded data for " & Chr(34) & mLabel & Chr(34) & " into TODAYS EXPENDITURE.", vbInformation, "Done"
End Sub


Public Sub AutoPaySalaries()
    On Error GoTo ErrorHandler
    Dim wsRent As Worksheet
    Dim wsExp As Worksheet
    
    On Error Resume Next
    Set wsRent = ThisWorkbook.Sheets("RENT DATA")
    Set wsExp = ThisWorkbook.Sheets("TODAYS EXPENDITURE")
    On Error GoTo ErrorHandler
    
    If wsRent Is Nothing Or wsExp Is Nothing Then
        MsgBox "Could not find required sheets.", vbExclamation
        Exit Sub
    End If
    
    ' 1. Calculate Total Rent Funds
    Dim lastRentRow As Long
    lastRentRow = wsRent.Cells(wsRent.Rows.Count, 13).End(xlUp).Row
    Dim totalRent As Double: totalRent = 0
    Dim r As Long
    For r = 2 To lastRentRow
        If IsNumeric(wsRent.Cells(r, 13).Value) And wsRent.Cells(r, 13).Value <> "" Then
            totalRent = totalRent + CDbl(wsRent.Cells(r, 13).Value)
        End If
    Next r
    
    ' 2. Calculate Already Paid Salaries
    Dim paidSalaries As Double: paidSalaries = 0
    Dim salaryAmount As Double
    For r = 3 To 24
        If UCase(Trim(CStr(wsExp.Cells(r, 5).Value))) = "PAID" Then
            If IsNumeric(wsExp.Cells(r, 2).Value) Then
                paidSalaries = paidSalaries + CDbl(wsExp.Cells(r, 2).Value)
            End If
        End If
    Next r
    
    Dim remainingFunds As Double
    remainingFunds = totalRent - paidSalaries
    
    ' Quick update to Ledger UI before processing
    Call UpdateLedgerUI(totalRent, paidSalaries, remainingFunds)
    
    ' Prompt user
    Dim ans As Integer
    ans = MsgBox("Funds Available: " & Format(remainingFunds, "Currency") & vbCrLf & _
                 "Do you want to automatically pay the remaining salaries from top to bottom?", _
                 vbYesNo + vbQuestion, "Auto-Pay Salaries")
    If ans <> vbYes Then Exit Sub
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    ' 3. Process Payments
    Dim newlyPaid As Integer: newlyPaid = 0
    For r = 3 To 24
        If UCase(Trim(CStr(wsExp.Cells(r, 5).Value))) <> "PAID" Then
            If IsNumeric(wsExp.Cells(r, 2).Value) And wsExp.Cells(r, 2).Value <> "" Then
                salaryAmount = CDbl(wsExp.Cells(r, 2).Value)
                If remainingFunds >= salaryAmount Then
                    ' Pay it!
                    wsExp.Cells(r, 5).Value = "PAID"
                    wsExp.Cells(r, 4).Value = Date ' Today's date
                    remainingFunds = remainingFunds - salaryAmount
                    paidSalaries = paidSalaries + salaryAmount
                    newlyPaid = newlyPaid + 1
                End If
            End If
        End If
    Next r
    
    ' Final update to Ledger UI
    Call UpdateLedgerUI(totalRent, paidSalaries, remainingFunds)
    
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    
    MsgBox "Successfully paid " & newlyPaid & " salaries!" & vbCrLf & _
           "Funds Left: " & Format(remainingFunds, "Currency"), vbInformation, "Auto-Pay Complete"
    
    Exit Sub
ErrorHandler:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    MsgBox "An error occurred: " & Err.Description, vbCritical
End Sub

Public Sub UpdateLedgerUI(totalRent As Double, paidSalaries As Double, remainingFunds As Double)
    On Error Resume Next
    Dim wsExp As Worksheet
    Set wsExp = ThisWorkbook.Sheets("TODAYS EXPENDITURE")
    If Not wsExp Is Nothing Then
        wsExp.Range("K1").Value = totalRent
        wsExp.Range("M1").Value = paidSalaries
        wsExp.Range("O1").Value = remainingFunds
    End If
End Sub

