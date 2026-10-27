import win32com.client
import os

excel = win32com.client.Dispatch('Excel.Application')
excel.Visible = False
excel.DisplayAlerts = False
file_path = os.path.abspath('RENTAL DATASHEET 12TH AUG.xlsm')

vba_update_summary = '''Public Sub UpdateExpenditureSummary()
    Dim wsExp As Worksheet
    Set wsExp = ThisWorkbook.Sheets("todays_expenditure")
    
    Application.ScreenUpdating = False
    
    ' 1. Calculate Total Rent Collected
    Dim targetMonthStr As String
    targetMonthStr = UCase(Trim(wsExp.Range("A2").Value))
    
    Dim totalRentCollected As Double
    totalRentCollected = 0
    
    If targetMonthStr <> "" Then
        Dim targetMonth As Integer, targetYear As Integer
        If InStr(targetMonthStr, "_") > 0 Or InStr(targetMonthStr, " ") > 0 Then
            Dim parts() As String
            If InStr(targetMonthStr, "_") > 0 Then parts = Split(targetMonthStr, "_") Else parts = Split(targetMonthStr, " ")
            On Error Resume Next
            targetMonth = Month(DateValue("1 " & parts(0) & " 2000"))
            targetYear = CInt(parts(1))
            On Error GoTo 0
        End If
        
        If targetMonth > 0 And targetYear > 0 Then
            Dim ws As Worksheet
            Dim r As Long, lastAliasRow As Long
            Dim pDate As Variant, forMonthStr As String, monthChecked As Boolean, addPayment As Boolean, amt As Double
            Dim ignoreSheets As String
            ignoreSheets = "|RENT DATA|TODAYS_RENT_DATA|DATABASE_RENT_DATA|PAYMENT SHEET|EXPENDITURE|PURCHASE|DATABASE_PURCHASE|PURCHASE SHEET|SHEET17|TODAYS_EXPENDITURE|DATABASE_EXPENDITURE|"
            
            For Each ws In ThisWorkbook.Sheets
                If InStr(1, ignoreSheets, "|" & UCase(ws.Name) & "|") = 0 Then
                    lastAliasRow = ws.Cells(ws.Rows.Count, "D").End(xlUp).Row
                    If lastAliasRow >= 2 Then
                        For r = 2 To lastAliasRow
                            pDate = ws.Cells(r, 1).Value
                            addPayment = False
                            monthChecked = False
                            
                            If Not IsError(ws.Cells(r, 3).Value) Then
                                forMonthStr = UCase(Trim(CStr(ws.Cells(r, 3).Value)))
                                If forMonthStr <> "" Then
                                    monthChecked = True
                                    If InStr(forMonthStr, UCase(MonthName(targetMonth))) > 0 Then addPayment = True
                                End If
                            End If
                            
                            If Not monthChecked Then
                                If IsDate(pDate) Then
                                    If Month(pDate) = targetMonth And Year(pDate) = targetYear Then addPayment = True
                                End If
                            End If
                            
                            If addPayment Then
                                amt = 0
                                If IsNumeric(ws.Cells(r, 6).Value) Then amt = CDbl(ws.Cells(r, 6).Value)
                                totalRentCollected = totalRentCollected + amt
                            End If
                        Next r
                    End If
                End If
            Next ws
        End If
    End If
    
    ' 2. Calculate Total Paid and Due
    Dim totalPaid As Double, totalDue As Double
    totalPaid = 0: totalDue = 0
    
    Dim lastRow As Long
    lastRow = wsExp.Cells(wsExp.Rows.Count, "A").End(xlUp).Row
    If lastRow < 3 Then lastRow = 30
    
    ' Find the summary row (starts with TOTAL RENT COLLECTED)
    Dim summaryRow As Long
    summaryRow = 0
    For r = lastRow To 1 Step -1
        If UCase(Trim(wsExp.Cells(r, 1).Value)) = "TOTAL RENT COLLECTED" Then
            summaryRow = r
            Exit For
        End If
    Next r
    
    If summaryRow = 0 Then
        ' Create summary block if missing
        summaryRow = lastRow + 3
        wsExp.Cells(summaryRow, 1).Value = "TOTAL RENT COLLECTED"
        wsExp.Cells(summaryRow, 2).Value = "PAID AMOUNT"
        wsExp.Cells(summaryRow, 3).Value = "DUE AMOUNT"
        wsExp.Cells(summaryRow, 4).Value = "AVAILABLE BALANCE"
        With wsExp.Range(wsExp.Cells(summaryRow, 1), wsExp.Cells(summaryRow, 4))
            .Interior.Color = RGB(41, 128, 185)
            .Font.Color = RGB(255, 255, 255)
            .Font.Bold = True
        End With
    End If
    
    Dim c As Long, header As String, isPaid As Boolean, cellAmt As Double
    For c = 2 To 100
        header = UCase(Trim(wsExp.Cells(1, c).Value))
        If header = "" Then header = UCase(Trim(wsExp.Cells(2, c).Value))
        If header = "" Then header = UCase(Trim(wsExp.Cells(3, c).Value))
        
        If header = "AMOUNT" Or header = "SALARIES" Or header = "OVERALL TOTAL" Then
            For r = 3 To summaryRow - 1
                If IsNumeric(wsExp.Cells(r, c).Value) And wsExp.Cells(r, c).Value <> "" Then
                    cellAmt = CDbl(wsExp.Cells(r, c).Value)
                    ' Check if next column is PAID and says PAID
                    isPaid = False
                    If UCase(Trim(wsExp.Cells(r, c + 1).Value)) = "PAID" Then isPaid = True
                    
                    If isPaid Then
                        totalPaid = totalPaid + cellAmt
                    Else
                        totalDue = totalDue + cellAmt
                    End If
                End If
            Next r
        End If
    Next c
    
    ' 3. Write Summary
    wsExp.Cells(summaryRow + 1, 1).Value = totalRentCollected
    wsExp.Cells(summaryRow + 1, 2).Value = totalPaid
    wsExp.Cells(summaryRow + 1, 3).Value = totalDue
    wsExp.Cells(summaryRow + 1, 4).Value = totalRentCollected - totalPaid
    
    With wsExp.Range(wsExp.Cells(summaryRow + 1, 1), wsExp.Cells(summaryRow + 1, 4))
        .Font.Bold = True
        .NumberFormat = "#,##0"
    End With
    
    Application.ScreenUpdating = True
End Sub
'''

vba_ws_change = '''Private Sub Worksheet_Change(ByVal Target As Range)
    ' Auto update summary when PAID or AMOUNT is typed
    Dim updateSummary As Boolean
    updateSummary = False
    
    Dim cell As Range
    For Each cell In Target
        If cell.Row > 2 Then
            updateSummary = True
            Exit For
        End If
    Next cell
    
    If updateSummary Then
        Application.EnableEvents = False
        On Error Resume Next
        Call ThisWorkbook.UpdateExpenditureSummary
        On Error GoTo 0
        Application.EnableEvents = True
    End If
End Sub
'''

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
            r.Range(1, 9).Value = tblEntry.DataBodyRange.Cells(i, 8).Value
            itemsSaved = itemsSaved + 1
        End If
    Next i
    
    ' Sort DB by Date Descending
    If Not tblDB.DataBodyRange Is Nothing Then
        With tblDB.Sort
            .SortFields.Clear
            .SortFields.Add Key:=tblDB.ListColumns(1).Range, SortOn:=0, Order:=2 ' xlDescending
            .Header = 1 ' xlYes
            .Apply
        End With
    End If
    
    ' Draw Borders Between Dates
    If Not tblDB.DataBodyRange Is Nothing Then
        ' Clear old borders
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
    
    MsgBox itemsSaved & " items successfully saved for " & Format(saveDate, "dd-mmm-yyyy") & "!", vbInformation, "Success"
End Sub
'''

try:
    wb = excel.Workbooks.Open(file_path)
    
    # 1. Update todays_rent_data Conditional Formatting
    wsRent = wb.Sheets('todays_rent_data')
    # Column M (TOTAL PAID)
    # Green if M >= L (Paid >= Amount)
    fc1 = wsRent.Range("M2:M1000").FormatConditions.Add(Type=2, Formula1='=AND($M2<>"", $L2<>"", $M2>=$L2)')
    fc1.Interior.Color = 13565902 # Light Green
    fc1.Font.Color = 24832 # Dark Green
    fc1.Font.Bold = True
    # Yellow/Light Green if 0 < M < L
    fc2 = wsRent.Range("M2:M1000").FormatConditions.Add(Type=2, Formula1='=AND($M2<>"", $L2<>"", $M2>0, $M2<$L2)')
    fc2.Interior.Color = 10092543 # Very Light Yellow/Green
    fc2.Font.Color = 0
    # Column N (GROUP DUE)
    # Red if N > 0
    fc3 = wsRent.Range("N2:N1000").FormatConditions.Add(Type=2, Formula1='=AND($N2<>"", $N2>0)')
    fc3.Interior.Color = 13421823 # Light Red
    fc3.Font.Color = 255 # Red
    fc3.Font.Bold = True

    # 2. Inject Summary Macro for todays_expenditure
    comp_thiswb = wb.VBProject.VBComponents('ThisWorkbook')
    thiswb_code = comp_thiswb.CodeModule.Lines(1, comp_thiswb.CodeModule.CountOfLines)
    if "Public Sub UpdateExpenditureSummary()" not in thiswb_code:
        comp_thiswb.CodeModule.AddFromString(vba_update_summary)

    # 3. Add Worksheet_Change to todays_expenditure
    wsExp = wb.Sheets('todays_expenditure')
    comp_exp = wb.VBProject.VBComponents(wsExp.CodeName)
    comp_exp.CodeModule.DeleteLines(1, comp_exp.CodeModule.CountOfLines)
    comp_exp.CodeModule.AddFromString(vba_ws_change)
    
    # 4. Re-create Buttons on todays_expenditure
    # Delete existing shapes to avoid duplicates
    for shp in wsExp.Shapes:
        shp.Delete()
    
    btn1 = wsExp.Buttons().Add(600, 10, 100, 30) # Save
    btn1.Caption = "SAVE DATA"
    btn1.OnAction = "SaveExpenditureData"
    btn2 = wsExp.Buttons().Add(720, 10, 100, 30) # Load
    btn2.Caption = "LOAD DATA"
    btn2.OnAction = "LoadExpenditureData"
    btn3 = wsExp.Buttons().Add(840, 10, 100, 30) # Start New
    btn3.Caption = "START NEW"
    btn3.OnAction = "StartNewMonth"
    btn4 = wsExp.Buttons().Add(960, 10, 100, 30) # AutoPay
    btn4.Caption = "AUTOPAY"
    btn4.OnAction = "AutoPaySalaries"
    
    # Run the summary update once right now
    excel.Application.Run("ThisWorkbook.UpdateExpenditureSummary")
    
    # 5. Fix ModDailyPurchases
    for comp in wb.VBProject.VBComponents:
        if comp.Name == 'ModDailyPurchases':
            comp.CodeModule.DeleteLines(1, comp.CodeModule.CountOfLines)
            comp.CodeModule.AddFromString(vba_mod_purchases)
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
