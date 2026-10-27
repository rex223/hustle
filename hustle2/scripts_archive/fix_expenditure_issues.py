import win32com.client
import os

def fix_issues():
    file_path = os.path.abspath('hustle2/RENTAL DATASHEET 1ST AUG.xlsm')
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    
    try:
        wb = excel.Workbooks.Open(file_path)
        
        # 1. Fix ThisWorkbook ignoreSheets
        wb_module = wb.VBProject.VBComponents("ThisWorkbook").CodeModule
        code = wb_module.Lines(1, wb_module.CountOfLines)
        
        # mac compatibility version has slightly different ignore strings
        old_ignore = 'ignoreSheets = "|RENT DATA|PAYMENT SHEET|EXPENDITURE|PURCHASE|DATABASE_PURCHASE|PURCHASE SHEET|SHEET17|"'
        new_ignore = 'ignoreSheets = "|RENT DATA|PAYMENT SHEET|EXPENDITURE|PURCHASE|DATABASE_PURCHASE|PURCHASE SHEET|DATABASE_EXPENDITURE|TODAYS EXPENDITURE|SHEET17|"'
        
        code = code.replace(old_ignore, new_ignore)
        
        wb_module.DeleteLines(1, wb_module.CountOfLines)
        wb_module.AddFromString(code)
        
        # 2. Run UpdateGroupSummary to clean up RENT DATA
        try:
            excel.Application.Run("ThisWorkbook.UpdateGroupSummary")
        except Exception as e:
            print("Error running UpdateGroupSummary:", e)

        # 3. Fix TODAYS EXPENDITURE UI and Logic
        ws_entry = wb.Sheets("TODAYS EXPENDITURE")
        ws_db = wb.Sheets("DATABASE_EXPENDITURE")
        
        # Fix Button Position
        for shape in ws_entry.Shapes:
            try:
                if "SAVE" in shape.TextFrame.Characters().Text:
                    shape.Left = 350
                    shape.Top = 10
            except:
                pass
                
        # Remove Worksheet_Change filter macro, keep only Worksheet_Activate
        vba_expenditure_ws = '''Private Sub Worksheet_Activate()
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
'''
        mod_exp = wb.VBProject.VBComponents(ws_entry.CodeName).CodeModule
        mod_exp.DeleteLines(1, mod_exp.CountOfLines)
        mod_exp.AddFromString(vba_expenditure_ws)
        
        # Pre-fill TODAYS EXPENDITURE with unique items
        tbl_entry = ws_entry.ListObjects("EntryFormExp")
        tbl_db = ws_db.ListObjects("ExpenditureLedger")
        
        unique_items = []
        seen = set()
        
        if tbl_db.DataBodyRange is not None:
            for i in range(1, tbl_db.DataBodyRange.Rows.Count + 1):
                cat = tbl_db.DataBodyRange.Cells(i, 2).Value
                paid_to = tbl_db.DataBodyRange.Cells(i, 4).Value
                loc = tbl_db.DataBodyRange.Cells(i, 5).Value
                rem = tbl_db.DataBodyRange.Cells(i, 7).Value
                
                cat = str(cat).strip() if cat else ""
                paid_to = str(paid_to).strip() if paid_to else ""
                loc = str(loc).strip() if loc else ""
                rem = str(rem).strip() if rem else ""
                
                key = (cat, paid_to, loc, rem)
                if key not in seen and (cat or paid_to):
                    seen.add(key)
                    unique_items.append(key)
                    
        # Sort items: Salary first, then others
        unique_items.sort(key=lambda x: (0 if x[0] == 'SALARY' else 1, x[0], x[1]))
        
        # Clear existing table data
        if tbl_entry.DataBodyRange is not None:
            tbl_entry.DataBodyRange.ClearContents()
            
            # Delete extra rows if any
            while tbl_entry.ListRows.Count > 1:
                tbl_entry.ListRows(1).Delete()
                
        # Populate table
        for idx, item in enumerate(unique_items):
            if idx >= tbl_entry.ListRows.Count:
                tbl_entry.ListRows.Add()
                
            r = idx + 1
            tbl_entry.DataBodyRange.Cells(r, 1).Value = item[0] # CATEGORY
            tbl_entry.DataBodyRange.Cells(r, 2).Value = "JULY" if item[0] == 'SALARY' else "" # MONTH/FOR default
            tbl_entry.DataBodyRange.Cells(r, 3).Value = item[1] # PAID TO
            tbl_entry.DataBodyRange.Cells(r, 4).Value = item[2] # LOCATION
            # Cell 5 is AMOUNT, leave blank
            tbl_entry.DataBodyRange.Cells(r, 6).Value = item[3] # REMARKS

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
    fix_issues()
