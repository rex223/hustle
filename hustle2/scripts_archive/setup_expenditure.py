import win32com.client
import os
import pandas as pd

def extract_expenditures(filepath):
    df = pd.read_excel(filepath, 'EXPENDITURE')
    records = []
    for idx, row in df.iterrows():
        # 1. Salary
        if pd.notna(row.get('NAME')):
            month_for = ""
            amount = ""
            date_val = ""
            if pd.notna(row.get('JULY')):
                if isinstance(row['JULY'], (int, float)):
                    amount = row['JULY']
                    month_for = "JULY"
                elif str(row['JULY']).upper() == 'PAID':
                    amount = row.get('SALARIES')
                    month_for = "JULY"
                else:
                    amount = row['JULY']
            if pd.notna(row.get('DATE')):
                date_val = row['DATE']
                
            records.append({
                'DATE': date_val, 'CATEGORY': 'SALARY', 'MONTH / FOR': month_for,
                'PAID TO / FOR': row['NAME'], 'LOCATION / TYPE': row.get('TYPE', ''),
                'AMOUNT': amount, 'REMARKS': f"Base: {row.get('SALARIES', '')}"
            })
            
        # 2. CESC / URP
        if 'CESC / URP ' in row and pd.notna(row['CESC / URP ']):
            records.append({
                'DATE': '', 'CATEGORY': 'ELECTRICITY', 'MONTH / FOR': '',
                'PAID TO / FOR': 'CESC / URP', 'LOCATION / TYPE': row['CESC / URP '],
                'AMOUNT': row.get('AMOUNT', ''), 'REMARKS': ''
            })
            
        # 3. CESC / RESD
        if 'CESC / RESD' in row and pd.notna(row['CESC / RESD']):
            records.append({
                'DATE': '', 'CATEGORY': 'ELECTRICITY', 'MONTH / FOR': '',
                'PAID TO / FOR': 'CESC / RESD', 'LOCATION / TYPE': row['CESC / RESD'],
                'AMOUNT': row.get('AMOUNT.1', ''), 'REMARKS': ''
            })

        # 4. LABOUR
        if 'LABOUR' in row and pd.notna(row['LABOUR']):
            records.append({
                'DATE': '', 'CATEGORY': 'LABOUR', 'MONTH / FOR': '',
                'PAID TO / FOR': row['LABOUR'], 'LOCATION / TYPE': '',
                'AMOUNT': row.get('AMOUNT.2', ''), 'REMARKS': ''
            })
            
        # 5. DIGRUM / URP
        if 'DIGRUM / URP' in row and pd.notna(row['DIGRUM / URP']):
            records.append({
                'DATE': '', 'CATEGORY': 'DIGRUM', 'MONTH / FOR': '',
                'PAID TO / FOR': row['DIGRUM / URP'], 'LOCATION / TYPE': 'URP',
                'AMOUNT': row.get('AMOUNT.3', ''), 'REMARKS': ''
            })
            
        # 6. ANIMAL FEED
        if 'ANIMAL FEED' in row and pd.notna(row['ANIMAL FEED']):
            records.append({
                'DATE': row.get('DATE.1', ''), 'CATEGORY': 'ANIMAL FEED', 'MONTH / FOR': '',
                'PAID TO / FOR': row['ANIMAL FEED'], 'LOCATION / TYPE': '',
                'AMOUNT': row.get('AMOUNT.4', ''), 'REMARKS': ''
            })
            
        # 7. FUEL / GAS
        if 'FUEL / GAS' in row and pd.notna(row['FUEL / GAS']):
            records.append({
                'DATE': row.get('DATE.2', ''), 'CATEGORY': 'FUEL / GAS', 'MONTH / FOR': '',
                'PAID TO / FOR': row['FUEL / GAS'], 'LOCATION / TYPE': '',
                'AMOUNT': row.get('AMOUNT.5', ''), 'REMARKS': ''
            })
            
        # 8. MEDICAL
        if 'MEDICAL' in row and pd.notna(row['MEDICAL']):
            records.append({
                'DATE': row.get('DATE.3', ''), 'CATEGORY': 'MEDICAL', 'MONTH / FOR': '',
                'PAID TO / FOR': row['MEDICAL'], 'LOCATION / TYPE': '',
                'AMOUNT': row.get('AMOUNT.6', ''), 'REMARKS': ''
            })
            
        # 9. REPAIR
        if 'REPAIR' in row and pd.notna(row['REPAIR']):
            records.append({
                'DATE': '', 'CATEGORY': 'REPAIR', 'MONTH / FOR': '',
                'PAID TO / FOR': row['REPAIR'], 'LOCATION / TYPE': row.get('LOCATION', ''),
                'AMOUNT': row.get('AMOUNT.7', ''), 'REMARKS': ''
            })
            
        # 10. RELIGION
        if 'RELIGION' in row and pd.notna(row['RELIGION']):
            records.append({
                'DATE': '', 'CATEGORY': 'RELIGION', 'MONTH / FOR': '',
                'PAID TO / FOR': row['RELIGION'], 'LOCATION / TYPE': '',
                'AMOUNT': row.get('OVERALL TOTAL', ''), 'REMARKS': ''
            })
    return records

vba_purchases_ws = '''Private Sub Worksheet_Activate()
    On Error GoTo SafeExit
    Dim wsDB As Worksheet
    Set wsDB = ThisWorkbook.Sheets("DATABASE_PURCHASE")
    Dim tblDB As ListObject
    Set tblDB = wsDB.ListObjects("DailyPurchaseLedger")
    
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

Private Sub Worksheet_Change(ByVal Target As Range)
    On Error GoTo ErrorHandler
    Dim dateCell As Range
    Set dateCell = Me.Range("C1")
    If Intersect(Target, dateCell) Is Nothing Then Exit Sub
    
    Dim wsDB As Worksheet
    Dim tblEntry As ListObject
    Dim tblDB As ListObject
    
    On Error Resume Next
    Set tblEntry = Me.ListObjects("EntryFormTable")
    Set wsDB = ThisWorkbook.Sheets("DATABASE_PURCHASE")
    Set tblDB = wsDB.ListObjects("DailyPurchaseLedger")
    On Error GoTo ErrorHandler
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    
    If tblEntry.DataBodyRange.Rows.Count > 0 Then
        tblEntry.ListColumns("QTY (KG)").DataBodyRange.ClearContents
        tblEntry.ListColumns("QTY (PCS)").DataBodyRange.ClearContents
        tblEntry.ListColumns("TRSP").DataBodyRange.ClearContents
    End If
    
    If IsDate(dateCell.Value) Then
        Dim searchDate As Date
        searchDate = CDate(dateCell.Value)
        
        Dim i As Long, j As Long
        Dim dbItem As String
        Dim entryItem As String
        
        If Not tblDB.DataBodyRange Is Nothing Then
            For i = 1 To tblDB.DataBodyRange.Rows.Count
                If IsDate(tblDB.DataBodyRange.Cells(i, 1).Value) Then
                    If CDate(tblDB.DataBodyRange.Cells(i, 1).Value) = searchDate Then
                        dbItem = Trim(tblDB.DataBodyRange.Cells(i, 3).Value)
                        
                        For j = 1 To tblEntry.DataBodyRange.Rows.Count
                            entryItem = Trim(tblEntry.DataBodyRange.Cells(j, 2).Value)
                            If UCase(dbItem) = UCase(entryItem) Then
                                tblEntry.DataBodyRange.Cells(j, 3).Value = tblDB.DataBodyRange.Cells(i, 4).Value
                                tblEntry.DataBodyRange.Cells(j, 4).Value = tblDB.DataBodyRange.Cells(i, 5).Value
                                tblEntry.DataBodyRange.Cells(j, 6).Value = tblDB.DataBodyRange.Cells(i, 7).Value
                                Exit For
                            End If
                        Next j
                    End If
                End If
            Next i
        End If
    End If
    
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub
ErrorHandler:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub
'''

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

Private Sub Worksheet_Change(ByVal Target As Range)
    On Error GoTo ErrorHandler
    Dim dateCell As Range
    Set dateCell = Me.Range("C1")
    If Intersect(Target, dateCell) Is Nothing Then Exit Sub
    
    Dim wsDB As Worksheet
    Dim tblEntry As ListObject
    Dim tblDB As ListObject
    
    On Error Resume Next
    Set tblEntry = Me.ListObjects("EntryFormExp")
    Set wsDB = ThisWorkbook.Sheets("DATABASE_EXPENDITURE")
    Set tblDB = wsDB.ListObjects("ExpenditureLedger")
    On Error GoTo ErrorHandler
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    
    If tblEntry.DataBodyRange.Rows.Count > 0 Then
        tblEntry.DataBodyRange.ClearContents
    End If
    
    If IsDate(dateCell.Value) Then
        Dim searchDate As Date
        searchDate = CDate(dateCell.Value)
        
        Dim i As Long
        Dim r As Long
        r = 1
        
        If Not tblDB.DataBodyRange Is Nothing Then
            For i = 1 To tblDB.DataBodyRange.Rows.Count
                If IsDate(tblDB.DataBodyRange.Cells(i, 1).Value) Then
                    If CDate(tblDB.DataBodyRange.Cells(i, 1).Value) = searchDate Then
                        If r > tblEntry.DataBodyRange.Rows.Count Then
                            tblEntry.ListRows.Add
                        End If
                        tblEntry.DataBodyRange.Cells(r, 1).Value = tblDB.DataBodyRange.Cells(i, 2).Value
                        tblEntry.DataBodyRange.Cells(r, 2).Value = tblDB.DataBodyRange.Cells(i, 3).Value
                        tblEntry.DataBodyRange.Cells(r, 3).Value = tblDB.DataBodyRange.Cells(i, 4).Value
                        tblEntry.DataBodyRange.Cells(r, 4).Value = tblDB.DataBodyRange.Cells(i, 5).Value
                        tblEntry.DataBodyRange.Cells(r, 5).Value = tblDB.DataBodyRange.Cells(i, 6).Value
                        tblEntry.DataBodyRange.Cells(r, 6).Value = tblDB.DataBodyRange.Cells(i, 7).Value
                        r = r + 1
                    End If
                End If
            Next i
        End If
    End If
    
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub
ErrorHandler:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub
'''

vba_mod_exp = '''Public Sub SaveDailyExpenditures()
    Dim wsEntry As Worksheet
    Dim wsDB As Worksheet
    Dim tblEntry As ListObject
    Dim tblDB As ListObject
    Dim dateCell As Range
    
    On Error Resume Next
    Set wsEntry = ThisWorkbook.Sheets("TODAYS EXPENDITURE")
    Set wsDB = ThisWorkbook.Sheets("DATABASE_EXPENDITURE")
    Set tblEntry = wsEntry.ListObjects("EntryFormExp")
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
    
    Dim amt As Variant
    Dim r As ListRow
    Dim itemsSaved As Long
    itemsSaved = 0
    
    For i = 1 To tblEntry.DataBodyRange.Rows.Count
        amt = tblEntry.DataBodyRange.Cells(i, 5).Value
        
        If IsNumeric(amt) And amt > 0 Then
            Set r = tblDB.ListRows.Add
            r.Range(1, 1).Value = saveDate
            r.Range(1, 2).Value = tblEntry.DataBodyRange.Cells(i, 1).Value
            r.Range(1, 3).Value = tblEntry.DataBodyRange.Cells(i, 2).Value
            r.Range(1, 4).Value = tblEntry.DataBodyRange.Cells(i, 3).Value
            r.Range(1, 5).Value = tblEntry.DataBodyRange.Cells(i, 4).Value
            r.Range(1, 6).Value = amt
            r.Range(1, 7).Value = tblEntry.DataBodyRange.Cells(i, 6).Value
            itemsSaved = itemsSaved + 1
        End If
    Next i
    
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    
    MsgBox itemsSaved & " expenditures successfully saved for " & Format(saveDate, "dd-mmm-yyyy") & "!", vbInformation, "Success"
End Sub
'''


def run_setup():
    file_path = os.path.abspath('hustle2/RENTAL DATASHEET 1ST AUG.xlsm')
    records = extract_expenditures(file_path)
    
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    wb = excel.Workbooks.Open(file_path)
    
    try:
        # 1. Setup DATABASE_EXPENDITURE
        sheet_db = "DATABASE_EXPENDITURE"
        ws_db = None
        for s in wb.Sheets:
            if s.Name.upper() == sheet_db.upper():
                ws_db = s
                break
                
        if ws_db:
            ws_db.Cells.Clear()
            for tbl in ws_db.ListObjects:
                tbl.Delete()
        else:
            ws_db = wb.Sheets.Add(After=wb.Sheets(wb.Sheets.Count))
            ws_db.Name = sheet_db
            
        headers = ["DATE", "CATEGORY", "MONTH / FOR", "PAID TO / FOR", "LOCATION / TYPE", "AMOUNT", "REMARKS"]
        ws_db.Range("A1:G1").Value = headers
        
        row_idx = 2
        for r in records:
            if pd.isna(r['AMOUNT']) and pd.isna(r['REMARKS']):
                continue
                
            date_val = r['DATE']
            if not pd.isna(date_val) and date_val != "":
                if hasattr(date_val, 'strftime'):
                    date_val = date_val.strftime('%Y-%m-%d')
                ws_db.Cells(row_idx, 1).Value = date_val
            else:
                ws_db.Cells(row_idx, 1).Value = ""
            ws_db.Cells(row_idx, 2).Value = r['CATEGORY']
            ws_db.Cells(row_idx, 3).Value = r['MONTH / FOR']
            ws_db.Cells(row_idx, 4).Value = r['PAID TO / FOR']
            ws_db.Cells(row_idx, 5).Value = r['LOCATION / TYPE']
            
            amt = r['AMOUNT']
            if not pd.isna(amt) and amt != "":
                ws_db.Cells(row_idx, 6).Value = float(amt)
                
            ws_db.Cells(row_idx, 7).Value = r['REMARKS']
            row_idx += 1
            
        table_range = f"A1:G{row_idx - 1 if row_idx > 2 else 2}"
        tbl_db = ws_db.ListObjects.Add(1, ws_db.Range(table_range), 1, 1)
        tbl_db.Name = "ExpenditureLedger"
        tbl_db.TableStyle = "TableStyleMedium14"
        
        ws_db.Columns("A").ColumnWidth = 12
        ws_db.Columns("B:E").ColumnWidth = 18
        ws_db.Columns("F").ColumnWidth = 12
        ws_db.Columns("G").ColumnWidth = 20
        ws_db.Columns("A").NumberFormat = "dd-mmm-yyyy"
        ws_db.Columns("F").NumberFormat = "#,##0.00"
        
        # 2. Setup TODAYS EXPENDITURE
        sheet_entry = "TODAYS EXPENDITURE"
        ws_entry = None
        for s in wb.Sheets:
            if s.Name.upper() == sheet_entry.upper():
                ws_entry = s
                break
                
        if ws_entry:
            ws_entry.Cells.Clear()
            for shape in ws_entry.Shapes:
                shape.Delete()
            for tbl in ws_entry.ListObjects:
                tbl.Delete()
        else:
            ws_entry = wb.Sheets.Add(Before=wb.Sheets(1))
            ws_entry.Name = sheet_entry
            
        ws_entry.Range("B1").Value = "SELECT DATE:"
        ws_entry.Range("B1").Font.Bold = True
        ws_entry.Range("B1").Font.Size = 12
        ws_entry.Range("B1").HorizontalAlignment = -4152
        
        ws_entry.Range("C1").Value = "01-Aug-2026"
        ws_entry.Range("C1").Font.Bold = True
        ws_entry.Range("C1").Font.Size = 12
        ws_entry.Range("C1").Interior.Color = 10092543
        ws_entry.Range("C1").Borders.LineStyle = 1
        ws_entry.Range("C1").NumberFormat = "dd-mmm-yyyy"
        
        ws_entry.Range("A3:F3").Value = ["CATEGORY", "MONTH / FOR", "PAID TO / FOR", "LOCATION / TYPE", "AMOUNT", "REMARKS"]
        table_range = "A3:F18"
        tbl_entry = ws_entry.ListObjects.Add(1, ws_entry.Range(table_range), 1, 1)
        tbl_entry.Name = "EntryFormExp"
        tbl_entry.TableStyle = "TableStyleMedium14"
        
        ws_entry.Columns("A:F").ColumnWidth = 18
        ws_entry.Columns("E").NumberFormat = "#,##0.00"
        
        tbl_entry.ListColumns("AMOUNT").DataBodyRange.Interior.Color = 14348258
        
        btn = ws_entry.Shapes.AddShape(1, 10, 10, 200, 40)
        btn.TextFrame.Characters().Text = "💾 SAVE TO LEDGER"
        btn.TextFrame.Characters().Font.Bold = True
        btn.TextFrame.Characters().Font.Size = 14
        btn.TextFrame.Characters().Font.Color = 16777215
        btn.Fill.ForeColor.RGB = 5287936
        btn.TextFrame.HorizontalAlignment = -4108
        btn.TextFrame.VerticalAlignment = -4108
        btn.OnAction = "SaveDailyExpenditures"
        
        ws_entry.Activate()
        excel.ActiveWindow.DisplayGridlines = False

        # 3. Inject VBA Code
        ws_purch = wb.Sheets("PURCHASE SHEET")
        mod_purch = wb.VBProject.VBComponents(ws_purch.CodeName).CodeModule
        mod_purch.DeleteLines(1, mod_purch.CountOfLines)
        mod_purch.AddFromString(vba_purchases_ws)
        
        mod_exp = wb.VBProject.VBComponents(ws_entry.CodeName).CodeModule
        mod_exp.DeleteLines(1, mod_exp.CountOfLines)
        mod_exp.AddFromString(vba_expenditure_ws)
        
        mod_name = "ModDailyExpenditure"
        module_exists = False
        for comp in wb.VBProject.VBComponents:
            if comp.Name == mod_name:
                module_exists = True
                comp.CodeModule.DeleteLines(1, comp.CodeModule.CountOfLines)
                comp.CodeModule.AddFromString(vba_mod_exp)
                break
                
        if not module_exists:
            new_module = wb.VBProject.VBComponents.Add(1)
            new_module.Name = mod_name
            new_module.CodeModule.AddFromString(vba_mod_exp)
            
        ws_purch.Activate()
        ws_purch.Range("C1").Select()
        ws_entry.Activate()
        ws_entry.Range("C1").Select()

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
    run_setup()
