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

    # 1. ADD ModGroupSummary MODULE
    vba_summary_code = """
Public Sub UpdateGroupSummary()
    Dim wsRent As Worksheet
    Set wsRent = ThisWorkbook.Sheets("RENT DATA")
    
    Dim dictTotal As Object, dictPaid As Object
    Set dictTotal = CreateObject("Scripting.Dictionary")
    Set dictPaid = CreateObject("Scripting.Dictionary")
    
    ' 1. Calculate Group Totals from RENT DATA
    Dim lastRow As Long
    lastRow = wsRent.Cells(wsRent.Rows.Count, "E").End(xlUp).Row
    
    Dim i As Long
    Dim grp As String
    Dim amt As Double
    
    For i = 2 To lastRow
        grp = Trim(CStr(wsRent.Cells(i, 5).Value))
        If grp <> "" And UCase(grp) <> "GRAND TOTAL" And UCase(Left(grp, 5)) <> "TOTAL" Then
            amt = 0
            If IsNumeric(wsRent.Cells(i, 3).Value) Then amt = CDbl(wsRent.Cells(i, 3).Value)
            dictTotal(grp) = dictTotal(grp) + amt
            ' Initialize paid to 0
            If Not dictPaid.Exists(grp) Then dictPaid(grp) = 0
        End If
    Next i
    
    ' 2. Calculate Group Paid from all individual sheets
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
                    grp = Trim(CStr(ws.Cells(r, 7).Value)) ' Group is col 7 (G)
                    If grp <> "" Then
                        amt = 0
                        If IsNumeric(ws.Cells(r, 6).Value) Then amt = CDbl(ws.Cells(r, 6).Value)
                        dictPaid(grp) = dictPaid(grp) + amt
                        If Not dictTotal.Exists(grp) Then dictTotal(grp) = 0 ' Just in case
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
        wsRent.Cells(outRow, 13).Value = dictTotal(key) - dictPaid(key) ' M (Due)
        outRow = outRow + 1
    Next key
    
    wsRent.Columns("J:M").AutoFit
End Sub
"""
    # Check if ModGroupSummary exists
    mod_summary = None
    for comp in wb.VBProject.VBComponents:
        if comp.Name == "ModGroupSummary":
            mod_summary = comp
            break
            
    if not mod_summary:
        mod_summary = wb.VBProject.VBComponents.Add(1) # vbext_ct_StdModule
        mod_summary.Name = "ModGroupSummary"
        
    mod_summary.CodeModule.DeleteLines(1, mod_summary.CodeModule.CountOfLines) if mod_summary.CodeModule.CountOfLines > 0 else None
    mod_summary.CodeModule.AddFromString(vba_summary_code)

    # 2. UPDATE ThisWorkbook MODULE to call UpdateGroupSummary
    workbook_module = wb.VBProject.VBComponents("ThisWorkbook").CodeModule
    vba_code = workbook_module.Lines(1, workbook_module.CountOfLines)
    
    # We will inject a call to UpdateGroupSummary right before Application.EnableEvents = True 
    # at the end of the sections. Actually, it's safer to just let the user run it or call it in Worksheet_Change
    
    wb.Save()
    print("VBA Summary Module added successfully!")
    
except Exception as e:
    print(f"Error: {e}")
finally:
    try:
        wb.Close(False)
    except:
        pass
    excel.DisplayAlerts = True
