import win32com.client
import os

def run():
    file_path = os.path.abspath('hustle2/RENTAL DATASHEET 1ST AUG.xlsm')
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    
    try:
        wb = excel.Workbooks.Open(file_path)
        
        ws_old = wb.Sheets("EXPENDITURE")
        ws_entry = wb.Sheets("TODAYS EXPENDITURE")
        ws_db = wb.Sheets("DATABASE_EXPENDITURE")
        
        # --- 1. SETUP TODAYS EXPENDITURE ---
        ws_entry.Cells.Clear()
        for shape in ws_entry.Shapes:
            shape.Delete()
        while ws_entry.ListObjects.Count > 0:
            ws_entry.ListObjects(1).Delete()
            
        # Copy exact layout from old to row 3
        ws_old.Range("A1:AA100").Copy(ws_entry.Range("A3"))
        
        # Define ranges for colors (Rows 3 to 100)
        colors = {
            "A3:E100": 15773696,
            "F3:G100": 10092543,
            "H3:I100": 13434828,
            "J3:K100": 16764159,
            "L3:M100": 16751103,
            "N3:P100": 16764057,
            "Q3:S100": 16777164,
            "T3:V100": 13421823,
            "W3:Y100": 15132390,
            "Z3:AA100": 10092441
        }
        
        for rng, col in colors.items():
            ws_entry.Range(rng).Interior.Color = col
            ws_entry.Range(rng).Borders.LineStyle = 1
            ws_entry.Range(rng).Borders.ColorIndex = 15
            
        cr = "4:100"
        ws_entry.Range(f"D4:E100").ClearContents()
        ws_entry.Range(f"G4:G100").ClearContents()
        ws_entry.Range(f"I4:I100").ClearContents()
        ws_entry.Range(f"K4:K100").ClearContents()
        ws_entry.Range(f"M4:M100").ClearContents()
        ws_entry.Range(f"O4:P100").ClearContents()
        ws_entry.Range(f"R4:S100").ClearContents()
        ws_entry.Range(f"U4:V100").ClearContents()
        ws_entry.Range(f"Y4:Y100").ClearContents()
        ws_entry.Range(f"Z4:Z100").ClearContents()
        
        ws_entry.Range("E3").Value = "AMOUNT"
        ws_entry.Range("D3").Value = "" 
        ws_entry.Range("O3").Value = ""
        ws_entry.Range("R3").Value = ""
        ws_entry.Range("V3").Value = ""
        
        ws_entry.Range("A1:AA2").Interior.ColorIndex = 0
        ws_entry.Range("A1:B1").Merge()
        ws_entry.Range("A1").Value = "SELECT DATE:"
        ws_entry.Range("A1").Font.Bold = True
        ws_entry.Range("A1").Font.Size = 12
        ws_entry.Range("A1").HorizontalAlignment = -4108
        ws_entry.Range("A1").VerticalAlignment = -4108
        
        ws_entry.Range("C1").Value = "01-Aug-2026"
        ws_entry.Range("C1").Font.Bold = True
        ws_entry.Range("C1").Font.Size = 12
        ws_entry.Range("C1").Interior.Color = 65535
        ws_entry.Range("C1").Borders.LineStyle = 1
        ws_entry.Range("C1").NumberFormat = "dd-mmm-yyyy"
        ws_entry.Range("C1").HorizontalAlignment = -4108
        ws_entry.Range("C1").VerticalAlignment = -4108
        
        btn = ws_entry.Shapes.AddShape(1, 280, 5, 120, 25)
        btn.TextFrame.Characters().Text = "💾 SAVE TO LEDGER"
        btn.TextFrame.Characters().Font.Bold = True
        btn.TextFrame.Characters().Font.Size = 9
        btn.TextFrame.Characters().Font.Color = 16777215
        btn.Fill.ForeColor.RGB = 5287936
        btn.TextFrame.HorizontalAlignment = -4108
        btn.TextFrame.VerticalAlignment = -4108
        btn.OnAction = "SaveDailyExpenditures"
        
        ws_entry.Rows("1:1").RowHeight = 25
        ws_entry.Activate()
        excel.ActiveWindow.SplitRow = 2
        excel.ActiveWindow.FreezePanes = True
        
        # --- 2. SETUP DATABASE EXPENDITURE ---
        ws_db.Cells.Clear()
        for shape in ws_db.Shapes:
            shape.Delete()
        while ws_db.ListObjects.Count > 0:
            ws_db.ListObjects(1).Delete()
            
        ws_entry.Range("A3:AA3").Copy(ws_db.Range("B1"))
        
        colors_db = {
            "B1:F1": 15773696,
            "G1:H1": 10092543,
            "I1:J1": 13434828,
            "K1:L1": 16764159,
            "M1:N1": 16751103,
            "O1:Q1": 16764057,
            "R1:T1": 16777164,
            "U1:W1": 13421823,
            "X1:Z1": 15132390,
            "AA1:AB1": 10092441
        }
        for rng, col in colors_db.items():
            ws_db.Range(rng).Interior.Color = col
            ws_db.Range(rng).Borders.LineStyle = 1
            ws_db.Range(rng).Borders.ColorIndex = 15
            
        ws_db.Range("A1").Value = "ENTRY DATE"
        ws_db.Range("A1").Interior.Color = 12632256
        ws_db.Range("A1").Borders.LineStyle = 1
        ws_db.Range("A1").Borders.ColorIndex = 15
        
        ws_db.Range("A1:AB1").Font.Bold = True
        ws_db.Columns("A:AB").AutoFit()
        
        ws_db.Activate()
        excel.ActiveWindow.SplitRow = 1
        excel.ActiveWindow.FreezePanes = True

        # --- 3. REWRITE MACROS ---
        vba_save = """Public Sub SaveDailyExpenditures()
    Dim wsEntry As Worksheet
    Dim wsDB As Worksheet
    Dim dateCell As Range
    
    On Error Resume Next
    Set wsEntry = ThisWorkbook.Sheets("TODAYS EXPENDITURE")
    Set wsDB = ThisWorkbook.Sheets("DATABASE_EXPENDITURE")
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
    
    Dim rIdx As Long
    Dim dbNextRow As Long
    Dim itemsSaved As Long
    Dim hasData As Boolean
    
    dbNextRow = wsDB.Cells(wsDB.Rows.Count, "A").End(xlUp).Row + 1
    If dbNextRow < 2 Then dbNextRow = 2
    itemsSaved = 0
    
    For rIdx = 4 To 100
        hasData = False
        If Trim(CStr(wsEntry.Cells(rIdx, 5).Value)) <> "" Then hasData = True
        If Trim(CStr(wsEntry.Cells(rIdx, 7).Value)) <> "" Then hasData = True
        If Trim(CStr(wsEntry.Cells(rIdx, 9).Value)) <> "" Then hasData = True
        If Trim(CStr(wsEntry.Cells(rIdx, 11).Value)) <> "" Then hasData = True
        If Trim(CStr(wsEntry.Cells(rIdx, 13).Value)) <> "" Then hasData = True
        If Trim(CStr(wsEntry.Cells(rIdx, 16).Value)) <> "" Then hasData = True
        If Trim(CStr(wsEntry.Cells(rIdx, 19).Value)) <> "" Then hasData = True
        If Trim(CStr(wsEntry.Cells(rIdx, 21).Value)) <> "" Then hasData = True
        If Trim(CStr(wsEntry.Cells(rIdx, 25).Value)) <> "" Then hasData = True
        If Trim(CStr(wsEntry.Cells(rIdx, 26).Value)) <> "" Then hasData = True
        
        If hasData Then
            wsDB.Cells(dbNextRow, 1).Value = saveDate
            wsDB.Cells(dbNextRow, 1).NumberFormat = "dd-mmm-yyyy"
            
            wsDB.Range(wsDB.Cells(dbNextRow, 2), wsDB.Cells(dbNextRow, 28)).Value = wsEntry.Range(wsEntry.Cells(rIdx, 1), wsEntry.Cells(rIdx, 27)).Value
            
            dbNextRow = dbNextRow + 1
            itemsSaved = itemsSaved + 1
        End If
    Next rIdx
    
    wsEntry.Range("D4:E100").ClearContents
    wsEntry.Range("G4:G100").ClearContents
    wsEntry.Range("I4:I100").ClearContents
    wsEntry.Range("K4:K100").ClearContents
    wsEntry.Range("M4:M100").ClearContents
    wsEntry.Range("O4:P100").ClearContents
    wsEntry.Range("R4:S100").ClearContents
    wsEntry.Range("U4:V100").ClearContents
    wsEntry.Range("Y4:Y100").ClearContents
    wsEntry.Range("Z4:Z100").ClearContents
    
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    
    MsgBox itemsSaved & " rows successfully saved for " & Format(saveDate, "dd-mmm-yyyy") & "!", vbInformation, "Success"
End Sub
"""
        mod_name = "ModDailyExpenditure"
        for comp in wb.VBProject.VBComponents:
            if comp.Name == mod_name:
                comp.CodeModule.DeleteLines(1, comp.CodeModule.CountOfLines)
                comp.CodeModule.AddFromString(vba_save)
                break
                
        vba_activate = """Private Sub Worksheet_Activate()
    On Error GoTo SafeExit
    Dim wsDB As Worksheet
    Set wsDB = ThisWorkbook.Sheets("DATABASE_EXPENDITURE")
    
    Dim lastRow As Long
    lastRow = wsDB.Cells(wsDB.Rows.Count, "A").End(xlUp).Row
    
    If lastRow < 2 Then Exit Sub
    
    Dim uniqueDates() As String
    Dim count As Long
    count = 0
    Dim val As Variant
    Dim i As Long, j As Long
    Dim isUnique As Boolean
    
    For i = 2 To lastRow
        val = wsDB.Cells(i, 1).Value
        If IsDate(val) Then
            Dim strDate As String
            strDate = Format(val, "dd-mmm-yyyy")
            
            isUnique = True
            If count > 0 Then
                For j = 1 To count
                    If uniqueDates(j) = strDate Then
                        isUnique = False
                        Exit For
                    End If
                Next j
            End If
            
            If isUnique Then
                count = count + 1
                ReDim Preserve uniqueDates(1 To count)
                uniqueDates(count) = strDate
            End If
        End If
    Next i
    
    If count > 0 Then
        Dim datesStr As String
        datesStr = uniqueDates(1)
        For j = 2 To count
            datesStr = datesStr & "," & uniqueDates(j)
        Next j
        
        With Me.Range("C1").Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:=datesStr
        End With
    End If
SafeExit:
End Sub
"""
        mod_exp = wb.VBProject.VBComponents(ws_entry.CodeName).CodeModule
        mod_exp.DeleteLines(1, mod_exp.CountOfLines)
        mod_exp.AddFromString(vba_activate)

        ws_entry.Activate()
        wb.Save()
        print("SUCCESS")
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
    run()
