import win32com.client
import os

def redo_expenditure_ui():
    file_path = os.path.abspath('hustle2/RENTAL DATASHEET 1ST AUG.xlsm')
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    
    try:
        wb = excel.Workbooks.Open(file_path)
        
        ws_old = wb.Sheets("EXPENDITURE")
        ws_new = wb.Sheets("TODAYS EXPENDITURE")
        
        # 1. Completely clear the new sheet
        ws_new.Cells.Clear()
        for shape in ws_new.Shapes:
            shape.Delete()
            
        # Remove any tables
        while ws_new.ListObjects.Count > 0:
            ws_new.ListObjects(1).Delete()
            
        # 2. Copy formatting and text from old sheet to new sheet (offset by 2 rows)
        ws_old.Range("A1:AA100").Copy(ws_new.Range("A3"))
        
        # Clear out the value columns (amounts & dates) to create a clean entry form
        clear_range = "4:100"
        
        # Salary: D (Date), E (Amount)
        ws_new.Range(f"D4:D100").ClearContents()
        ws_new.Range(f"E4:E100").ClearContents()
        ws_new.Range("E3").Value = "AMOUNT" # Rename JULY to AMOUNT
        ws_new.Range("D3").Value = "" # Remove DATE header as it's at the top now
        
        # Electricity: G, I
        ws_new.Range(f"G4:G100").ClearContents()
        ws_new.Range(f"I4:I100").ClearContents()
        
        # Labour: K
        ws_new.Range(f"K4:K100").ClearContents()
        
        # Digrum: M
        ws_new.Range(f"M4:M100").ClearContents()
        
        # Animal Feed: O (Date), P (Amount)
        ws_new.Range(f"O4:O100").ClearContents()
        ws_new.Range(f"P4:P100").ClearContents()
        ws_new.Range("O3").Value = "" # Remove DATE
        
        # Fuel/Gas: R (Date), S (Amount)
        ws_new.Range(f"R4:R100").ClearContents()
        ws_new.Range(f"S4:S100").ClearContents()
        ws_new.Range("R3").Value = "" # Remove DATE
        
        # Medical: U (Amount), V (Date)
        ws_new.Range(f"U4:U100").ClearContents()
        ws_new.Range(f"V4:V100").ClearContents()
        ws_new.Range("V3").Value = "" # Remove DATE
        
        # Repair: Y (Amount)
        ws_new.Range(f"Y4:Y100").ClearContents()
        
        # Religion: Z (Amount)
        ws_new.Range(f"Z4:Z100").ClearContents()
        
        # 3. Setup Row 1 (Header, Date Dropdown, Button)
        ws_new.Range("A1:AA2").Interior.ColorIndex = 0 # No fill
        ws_new.Range("A1:B1").Merge()
        ws_new.Range("A1").Value = "SELECT DATE:"
        ws_new.Range("A1").Font.Bold = True
        ws_new.Range("A1").Font.Size = 14
        ws_new.Range("A1").HorizontalAlignment = -4108 # Center
        ws_new.Range("A1").VerticalAlignment = -4108 # Center
        
        ws_new.Range("C1").Value = "01-Aug-2026"
        ws_new.Range("C1").Font.Bold = True
        ws_new.Range("C1").Font.Size = 14
        ws_new.Range("C1").Interior.Color = 10092543 # Yellow
        ws_new.Range("C1").Borders.LineStyle = 1
        ws_new.Range("C1").NumberFormat = "dd-mmm-yyyy"
        ws_new.Range("C1").HorizontalAlignment = -4108 # Center
        ws_new.Range("C1").VerticalAlignment = -4108 # Center
        
        # Add Save Button at E1
        btn = ws_new.Shapes.AddShape(1, 300, 5, 200, 35) # msoShapeRectangle
        btn.TextFrame.Characters().Text = "💾 SAVE TO LEDGER"
        btn.TextFrame.Characters().Font.Bold = True
        btn.TextFrame.Characters().Font.Size = 14
        btn.TextFrame.Characters().Font.Color = 16777215 # White
        btn.Fill.ForeColor.RGB = 5287936 # Green
        btn.TextFrame.HorizontalAlignment = -4108
        btn.TextFrame.VerticalAlignment = -4108
        btn.OnAction = "SaveDailyExpenditures"
        
        # Ensure row 1 is tall enough
        ws_new.Rows("1:1").RowHeight = 35
        
        # Freeze panes so row 1 and 2 are always visible
        ws_new.Activate()
        excel.ActiveWindow.SplitRow = 2
        excel.ActiveWindow.FreezePanes = True
        
        # 4. Inject VBA for Date Dropdown
        vba_activate = """Private Sub Worksheet_Activate()
    On Error GoTo SafeExit
    Dim wsDB As Worksheet
    Set wsDB = ThisWorkbook.Sheets("DATABASE_EXPENDITURE")
    Dim tblDB As ListObject
    Set tblDB = wsDB.ListObjects("ExpenditureLedger")
    
    If tblDB Is Nothing Then Exit Sub
    If tblDB.DataBodyRange Is Nothing Then Exit Sub
    
    Dim uniqueDates() As String
    Dim count As Long
    count = 0
    Dim val As Variant
    Dim i As Long, j As Long
    Dim isUnique As Boolean
    
    For i = 1 To tblDB.DataBodyRange.Rows.Count
        val = tblDB.DataBodyRange.Cells(i, 1).Value
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
        mod_exp = wb.VBProject.VBComponents(ws_new.CodeName).CodeModule
        mod_exp.DeleteLines(1, mod_exp.CountOfLines)
        mod_exp.AddFromString(vba_activate)

        # 5. Inject Global Macro
        vba_save = """Public Sub SaveDailyExpenditures()
    Dim wsEntry As Worksheet
    Dim wsDB As Worksheet
    Dim tblDB As ListObject
    Dim dateCell As Range
    
    On Error Resume Next
    Set wsEntry = ThisWorkbook.Sheets("TODAYS EXPENDITURE")
    Set wsDB = ThisWorkbook.Sheets("DATABASE_EXPENDITURE")
    Set tblDB = wsDB.ListObjects("ExpenditureLedger")
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
    
    Dim r As ListRow
    Dim itemsSaved As Long
    itemsSaved = 0
    Dim amt As Variant
    Dim rIdx As Long
    
    ' Salary (Cols A-E)
    For rIdx = 4 To 100
        amt = wsEntry.Cells(rIdx, 5).Value ' Col E
        If Trim(CStr(amt)) <> "" Then
            Dim finalAmt As Double
            finalAmt = 0
            If UCase(Trim(CStr(amt))) = "PAID" Then
                If IsNumeric(wsEntry.Cells(rIdx, 2).Value) Then finalAmt = CDbl(wsEntry.Cells(rIdx, 2).Value)
            ElseIf IsNumeric(amt) Then
                finalAmt = CDbl(amt)
            End If
            
            If finalAmt > 0 Then
                Set r = tblDB.ListRows.Add
                r.Range(1, 1).Value = saveDate
                r.Range(1, 2).Value = "SALARY"
                r.Range(1, 3).Value = "JULY" ' Optional month identifier
                r.Range(1, 4).Value = wsEntry.Cells(rIdx, 1).Value
                r.Range(1, 5).Value = wsEntry.Cells(rIdx, 3).Value
                r.Range(1, 6).Value = finalAmt
                r.Range(1, 7).Value = "Base: " & wsEntry.Cells(rIdx, 2).Value
                itemsSaved = itemsSaved + 1
            End If
        End If
    Next rIdx
    
    ' CESC URP (Cols F-G)
    For rIdx = 4 To 100
        amt = wsEntry.Cells(rIdx, 7).Value ' Col G
        If IsNumeric(amt) And amt > 0 Then
            Set r = tblDB.ListRows.Add
            r.Range(1, 1).Value = saveDate
            r.Range(1, 2).Value = "ELECTRICITY"
            r.Range(1, 4).Value = "CESC / URP"
            r.Range(1, 5).Value = wsEntry.Cells(rIdx, 6).Value
            r.Range(1, 6).Value = CDbl(amt)
            itemsSaved = itemsSaved + 1
        End If
    Next rIdx
    
    ' CESC RESD (Cols H-I)
    For rIdx = 4 To 100
        amt = wsEntry.Cells(rIdx, 9).Value ' Col I
        If IsNumeric(amt) And amt > 0 Then
            Set r = tblDB.ListRows.Add
            r.Range(1, 1).Value = saveDate
            r.Range(1, 2).Value = "ELECTRICITY"
            r.Range(1, 4).Value = "CESC / RESD"
            r.Range(1, 5).Value = wsEntry.Cells(rIdx, 8).Value
            r.Range(1, 6).Value = CDbl(amt)
            itemsSaved = itemsSaved + 1
        End If
    Next rIdx
    
    ' LABOUR (Cols J-K)
    For rIdx = 4 To 100
        amt = wsEntry.Cells(rIdx, 11).Value ' Col K
        If IsNumeric(amt) And amt > 0 Then
            Set r = tblDB.ListRows.Add
            r.Range(1, 1).Value = saveDate
            r.Range(1, 2).Value = "LABOUR"
            r.Range(1, 4).Value = wsEntry.Cells(rIdx, 10).Value
            r.Range(1, 6).Value = CDbl(amt)
            itemsSaved = itemsSaved + 1
        End If
    Next rIdx
    
    ' DIGRUM URP (Cols L-M)
    For rIdx = 4 To 100
        amt = wsEntry.Cells(rIdx, 13).Value ' Col M
        If IsNumeric(amt) And amt > 0 Then
            Set r = tblDB.ListRows.Add
            r.Range(1, 1).Value = saveDate
            r.Range(1, 2).Value = "DIGRUM"
            r.Range(1, 4).Value = wsEntry.Cells(rIdx, 12).Value
            r.Range(1, 5).Value = "URP"
            r.Range(1, 6).Value = CDbl(amt)
            itemsSaved = itemsSaved + 1
        End If
    Next rIdx
    
    ' ANIMAL FEED (Cols N-P)
    For rIdx = 4 To 100
        amt = wsEntry.Cells(rIdx, 16).Value ' Col P
        If IsNumeric(amt) And amt > 0 Then
            Set r = tblDB.ListRows.Add
            r.Range(1, 1).Value = saveDate
            r.Range(1, 2).Value = "ANIMAL FEED"
            r.Range(1, 4).Value = wsEntry.Cells(rIdx, 14).Value
            r.Range(1, 6).Value = CDbl(amt)
            itemsSaved = itemsSaved + 1
        End If
    Next rIdx
    
    ' FUEL / GAS (Cols Q-S)
    For rIdx = 4 To 100
        amt = wsEntry.Cells(rIdx, 19).Value ' Col S
        If IsNumeric(amt) And amt > 0 Then
            Set r = tblDB.ListRows.Add
            r.Range(1, 1).Value = saveDate
            r.Range(1, 2).Value = "FUEL / GAS"
            r.Range(1, 4).Value = wsEntry.Cells(rIdx, 17).Value
            r.Range(1, 6).Value = CDbl(amt)
            itemsSaved = itemsSaved + 1
        End If
    Next rIdx
    
    ' MEDICAL (Cols T-V)
    For rIdx = 4 To 100
        amt = wsEntry.Cells(rIdx, 21).Value ' Col U
        If IsNumeric(amt) And amt > 0 Then
            Set r = tblDB.ListRows.Add
            r.Range(1, 1).Value = saveDate
            r.Range(1, 2).Value = "MEDICAL"
            r.Range(1, 4).Value = wsEntry.Cells(rIdx, 20).Value
            r.Range(1, 6).Value = CDbl(amt)
            itemsSaved = itemsSaved + 1
        End If
    Next rIdx
    
    ' REPAIR (Cols W-Y)
    For rIdx = 4 To 100
        amt = wsEntry.Cells(rIdx, 25).Value ' Col Y
        If IsNumeric(amt) And amt > 0 Then
            Set r = tblDB.ListRows.Add
            r.Range(1, 1).Value = saveDate
            r.Range(1, 2).Value = "REPAIR"
            r.Range(1, 4).Value = wsEntry.Cells(rIdx, 24).Value
            r.Range(1, 5).Value = wsEntry.Cells(rIdx, 23).Value
            r.Range(1, 6).Value = CDbl(amt)
            itemsSaved = itemsSaved + 1
        End If
    Next rIdx
    
    ' RELIGION (Cols Z-AA)
    For rIdx = 4 To 100
        amt = wsEntry.Cells(rIdx, 26).Value ' Col Z
        If IsNumeric(amt) And amt > 0 Then
            Set r = tblDB.ListRows.Add
            r.Range(1, 1).Value = saveDate
            r.Range(1, 2).Value = "RELIGION"
            r.Range(1, 4).Value = wsEntry.Cells(rIdx, 27).Value
            r.Range(1, 6).Value = CDbl(amt)
            itemsSaved = itemsSaved + 1
        End If
    Next rIdx
    
    ' Clear Inputs for next entry
    wsEntry.Range("E4:E100").ClearContents
    wsEntry.Range("G4:G100").ClearContents
    wsEntry.Range("I4:I100").ClearContents
    wsEntry.Range("K4:K100").ClearContents
    wsEntry.Range("M4:M100").ClearContents
    wsEntry.Range("P4:P100").ClearContents
    wsEntry.Range("S4:S100").ClearContents
    wsEntry.Range("U4:U100").ClearContents
    wsEntry.Range("Y4:Y100").ClearContents
    wsEntry.Range("Z4:Z100").ClearContents
    
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    
    MsgBox itemsSaved & " expenditures successfully saved for " & Format(saveDate, "dd-mmm-yyyy") & "!", vbInformation, "Success"
End Sub
"""
        mod_name = "ModDailyExpenditure"
        module_exists = False
        for comp in wb.VBProject.VBComponents:
            if comp.Name == mod_name:
                module_exists = True
                comp.CodeModule.DeleteLines(1, comp.CodeModule.CountOfLines)
                comp.CodeModule.AddFromString(vba_save)
                break

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
    redo_expenditure_ui()
