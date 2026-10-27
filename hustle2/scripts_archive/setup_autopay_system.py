"""
setup_autopay_system.py
Adds the Auto-Pay Salaries button, Ledger UI, and injects the macro.
Run with Excel CLOSED.
"""
import win32com.client
import os

WORKBOOK_PATH = os.path.abspath('hustle2/RENTAL DATASHEET 1ST AUG.xlsm')

VBA_AUTOPAY = """
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
"""

def excol(r, g, b):
    return b | (g << 8) | (r << 16)

def setup():
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    
    try:
        wb = excel.Workbooks.Open(WORKBOOK_PATH)
        ws = wb.Sheets('TODAYS EXPENDITURE')
        
        # ==========================================
        # 1. SETUP LEDGER UI IN ROW 1 (J1 to O1)
        # ==========================================
        labels = [("J1", "TOTAL FUNDS:"), ("L1", "USED:"), ("N1", "LEFT:")]
        for cell, text in labels:
            rng = ws.Range(cell)
            rng.Value = text
            rng.Font.Bold = True
            rng.Font.Size = 10
            rng.Font.Color = excol(255, 255, 255)
            rng.Interior.Color = excol(52, 58, 64) # Dark gray
            rng.HorizontalAlignment = -4108 # Center
            rng.VerticalAlignment = -4108
            rng.Borders.LineStyle = 1
        
        values = ["K1", "M1", "O1"]
        for cell in values:
            rng = ws.Range(cell)
            rng.Value = 0.0
            rng.Font.Bold = True
            rng.Font.Size = 11
            rng.Interior.Color = excol(248, 249, 250) # Light gray
            rng.HorizontalAlignment = -4152 # Right align
            rng.VerticalAlignment = -4108
            rng.Borders.LineStyle = 1
            rng.NumberFormat = "[$₹-en-IN]#,##0.00"
            
        ws.Columns("J:J").ColumnWidth = 14
        ws.Columns("K:K").ColumnWidth = 12
        ws.Columns("L:L").ColumnWidth = 8
        ws.Columns("M:M").ColumnWidth = 12
        ws.Columns("N:N").ColumnWidth = 8
        ws.Columns("O:O").ColumnWidth = 12
        
        # ==========================================
        # 2. ADD AUTO-PAY BUTTON (Cell I1)
        # ==========================================
        # Check if it already exists
        btn_exists = False
        for shape in ws.Shapes:
            if shape.Name == "BtnAutoPay":
                btn_exists = True
                break
                
        if not btn_exists:
            target_cell = ws.Range("I1")
            btn = ws.Shapes.AddShape(1, target_cell.Left + 2, target_cell.Top + 2, target_cell.Width - 4, target_cell.Height - 4)
            btn.Name = "BtnAutoPay"
            btn.TextFrame2.TextRange.Text = "AUTO-PAY\nSALARIES"
            btn.TextFrame2.TextRange.Font.Bold = True
            btn.TextFrame2.TextRange.Font.Size = 9
            btn.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = excol(255, 255, 255)
            btn.TextFrame2.VerticalAnchor = 3 # msoAnchorMiddle
            btn.TextFrame2.TextRange.ParagraphFormat.Alignment = 2 # msoAlignCenter
            btn.Fill.ForeColor.RGB = excol(255, 193, 7) # Amber/Gold
            btn.Line.Visible = False
            btn.OnAction = "ModExpenditureSystem.AutoPaySalaries"
            
        ws.Columns("I:I").ColumnWidth = 12
        
        # ==========================================
        # 3. INJECT VBA
        # ==========================================
        bas_dir = os.path.normpath(os.path.join(os.path.dirname(WORKBOOK_PATH), 'scripts_archive'))
        module_bas = os.path.join(bas_dir, 'vba_module_expenditure.bas')
        
        # We append the new VBA to our existing bas file for persistence
        with open(module_bas, 'a', encoding='utf-8') as f:
            f.write("\n" + VBA_AUTOPAY)
            
        # And also inject it into the workbook directly
        with open(module_bas, 'r', encoding='utf-8') as f:
            full_vba = f.read()
            
        mod_name = "ModExpenditureSystem"
        for comp in wb.VBProject.VBComponents:
            if comp.Name == mod_name:
                comp.CodeModule.DeleteLines(1, comp.CodeModule.CountOfLines)
                comp.CodeModule.AddFromString(full_vba)
                break
                
        wb.Save()
        print("SUCCESS: UI and Macro added.")
        
    except Exception as e:
        import traceback
        traceback.print_exc()
        print(f"Error: {e}")
    finally:
        try:
            wb.Close(False)
        except:
            pass
        excel.Quit()

if __name__ == '__main__':
    setup()
