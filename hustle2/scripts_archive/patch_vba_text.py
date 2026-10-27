import os

with open("vba_current/ThisWorkbook.txt", "r", encoding="utf-8") as f:
    content = f.read()

target_marker = "    ' 2. Clear previous summary table"
idx = content.find(target_marker)

if idx == -1:
    print("Marker not found!")
else:
    new_tail = """    ' -------------------------------------------------------------------------
    ' 2. FIND TOTAL ROW AND DATA BOUNDS
    ' -------------------------------------------------------------------------
    Dim totRow As Long, lastDataRow As Long
    totRow = 0
    For r = 1 To 100
        If UCase(Trim(wsExp.Cells(r, 1).Value)) = "TOTAL" Or UCase(Trim(wsExp.Cells(r, 1).Value)) = "GRAND TOTAL" Then
            totRow = r
            Exit For
        End If
    Next r
    If totRow > 0 Then
        lastDataRow = totRow - 1
    Else
        lastDataRow = wsExp.Cells(wsExp.Rows.Count, "A").End(xlUp).Row
    End If
    If lastDataRow < 2 Then lastDataRow = 2
    
    ' -------------------------------------------------------------------------
    ' 3. CALCULATE PAID & DUE (STRICTLY IGNORING TOTALS AND RELIGION)
    ' -------------------------------------------------------------------------
    Dim totalPaid As Double, totalDue As Double
    totalPaid = 0
    totalDue = 0
    
    Dim c As Long, header As String, prevHeader As String, isPaid As Boolean, cellAmt As Double
    Dim rowName As String, itemDesc As String
    
    For c = 2 To 44
        header = UCase(Trim(wsExp.Cells(1, c).Value))
        prevHeader = ""
        If c > 1 Then prevHeader = UCase(Trim(wsExp.Cells(1, c - 1).Value))
        
        ' Process only salary and operating expense amount columns (skip Religion)
        If (header = "AMOUNT" Or header = "SALARIES") And prevHeader <> "RELIGION" Then
            For r = 3 To lastDataRow
                rowName = UCase(Trim(CStr(wsExp.Cells(r, 1).Value)))
                itemDesc = ""
                If c > 1 Then itemDesc = UCase(Trim(CStr(wsExp.Cells(r, c - 1).Value)))
                
                ' STRICT FILTER: Never count total / subtotal rows into Paid or Due!
                If rowName <> "TOTAL" And rowName <> "GRAND TOTAL" And _
                   itemDesc <> "TOTAL" And itemDesc <> "SUBTOTAL" And _
                   InStr(rowName, "TOTAL") = 0 And InStr(itemDesc, "TOTAL") = 0 Then
                   
                    If IsNumeric(wsExp.Cells(r, c).Value) And wsExp.Cells(r, c).Value <> "" Then
                        cellAmt = CDbl(wsExp.Cells(r, c).Value)
                        
                        If cellAmt > 0 Then
                            isPaid = False
                            
                            ' Look for "PAID" status in adjacent status columns
                            Dim pCol As Long
                            For pCol = c + 1 To c + 3
                                If pCol <= 44 Then
                                    If UCase(Trim(wsExp.Cells(r, pCol).Value)) = "PAID" Then
                                        isPaid = True
                                        Exit For
                                    End If
                                End If
                            Next pCol
                            
                            If isPaid Then
                                totalPaid = totalPaid + cellAmt
                            Else
                                totalDue = totalDue + cellAmt
                            End If
                        End If
                    End If
                End If
            Next r
        End If
    Next c
    
    Dim availBalance As Double
    availBalance = totalRentCollected - totalPurchases - totalPaid
    
    ' -------------------------------------------------------------------------
    ' 4. RENDER 6-CARD EXECUTIVE FINANCIAL POSITION BAR
    ' -------------------------------------------------------------------------
    Dim cardRow As Long
    cardRow = totRow + 3
    If cardRow < 36 Then cardRow = 36
    
    ' Row 36: Title Banner
    wsExp.Cells(cardRow, 1).Value = "FINANCIAL SUMMARY & EXECUTIVE CASH FLOW POSITION"
    With wsExp.Range(wsExp.Cells(cardRow, 1), wsExp.Cells(cardRow, 6))
        .Interior.Color = RGB(30, 41, 51)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .Font.Size = 11
        .HorizontalAlignment = xlCenter
    End With
    
    ' Row 37: 6 Card Headers
    wsExp.Cells(cardRow + 1, 1).Value = "TOTAL RENT COLLECTED"
    wsExp.Cells(cardRow + 1, 2).Value = "TOTAL PURCHASES"
    wsExp.Cells(cardRow + 1, 3).Value = "PAID AMOUNT"
    wsExp.Cells(cardRow + 1, 4).Value = "MISC AMOUNT PAID"
    wsExp.Cells(cardRow + 1, 5).Value = "DUE AMOUNT"
    wsExp.Cells(cardRow + 1, 6).Value = "AVAILABLE BALANCE"
    
    With wsExp.Range(wsExp.Cells(cardRow + 1, 1), wsExp.Cells(cardRow + 1, 6))
        .Interior.Color = RGB(41, 128, 185)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .Font.Size = 10
        .HorizontalAlignment = xlCenter
    End With
    
    ' Row 38: 6 Metric Values
    wsExp.Cells(cardRow + 2, 1).Value = totalRentCollected
    wsExp.Cells(cardRow + 2, 2).Value = totalPurchases
    wsExp.Cells(cardRow + 2, 3).Value = totalPaid
    wsExp.Cells(cardRow + 2, 4).Value = 0
    wsExp.Cells(cardRow + 2, 5).Value = totalDue
    wsExp.Cells(cardRow + 2, 6).Value = availBalance
    
    With wsExp.Range(wsExp.Cells(cardRow + 2, 1), wsExp.Cells(cardRow + 2, 6))
        .Font.Bold = True
        .Font.Size = 12
        .NumberFormat = "#,##0"
        .HorizontalAlignment = xlCenter
        .Interior.Color = RGB(255, 255, 255)
    End With
    
    ' Subtitles
    wsExp.Cells(cardRow + 3, 1).Value = "TOTAL INFLOW"
    wsExp.Cells(cardRow + 3, 2).Value = "SUPPLY EXPENSES"
    wsExp.Cells(cardRow + 3, 3).Value = "DISBURSED SALARIES"
    wsExp.Cells(cardRow + 3, 4).Value = "OUTGOING MISC"
    wsExp.Cells(cardRow + 3, 5).Value = "PENDING PAYABLES"
    wsExp.Cells(cardRow + 3, 6).Value = "NET LIQUIDITY"
    
    With wsExp.Range(wsExp.Cells(cardRow + 3, 1), wsExp.Cells(cardRow + 3, 6))
        .Font.Size = 8
        .Font.Color = RGB(100, 116, 139)
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With
    
    If availBalance < totalDue Or availBalance < 0 Then
        wsExp.Cells(cardRow + 2, 6).Font.Color = RGB(192, 0, 0)
    Else
        wsExp.Cells(cardRow + 2, 6).Font.Color = RGB(0, 128, 0)
    End If
End Sub
"""
    updated_content = content[:idx] + new_tail
    with open("vba_current/ThisWorkbook.txt", "w", encoding="utf-8") as f:
        f.write(updated_content)
    print("SUCCESS: vba_current/ThisWorkbook.txt updated successfully!")
