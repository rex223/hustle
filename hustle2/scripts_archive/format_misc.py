import win32com.client
import os

xl = win32com.client.Dispatch('Excel.Application')
xl.DisplayAlerts = False
wb = xl.Workbooks.Open(os.path.abspath('RENTAL DATASHEET 9 SEP.xlsm'))

try:
    ws = wb.Sheets("RENT PERSONS MISCELLANEOUS")
    
    # Unmerge any cells just in case
    ws.Cells.UnMerge()
    
    # Clear all existing formats
    ws.Cells.ClearFormats()
    
    # 1. Format INCOMING (Left Side) - Columns A to F
    # Headers
    headers_in = ["INCOMING", "DATE", "RECEIVE FROM", "DESCRIPTION", "AMOUNT", "STATUS"]
    for i, h in enumerate(headers_in):
        ws.Cells(1, i + 1).Value = h
    
    # Format Headers (Green)
    rng_in = ws.Range("A1:F1")
    rng_in.Interior.Color = 5287936 # RGB(0, 176, 80) roughly
    rng_in.Font.Color = 16777215 # White
    rng_in.Font.Bold = True
    rng_in.HorizontalAlignment = -4108 # Center
    
    # 2. Format OUTGOING (Right Side) - Columns H to L
    # We will put it in H to K (4 columns)
    headers_out = ["OUTGOING", "DATE", "PAID TO", "DESCRIPTION", "AMOUNT"]
    for i, h in enumerate(headers_out):
        ws.Cells(1, i + 8).Value = h
        
    # Format Headers (Red)
    rng_out = ws.Range("H1:L1")
    rng_out.Interior.Color = 255 # RGB(255, 0, 0)
    rng_out.Font.Color = 16777215 # White
    rng_out.Font.Bold = True
    rng_out.HorizontalAlignment = -4108 # Center
    
    # 3. Add Borders to Data (Assumes max 100 rows for now)
    last_row_in = max(ws.Cells(ws.Rows.Count, "B").End(-4162).Row, 2)
    last_row_out = max(ws.Cells(ws.Rows.Count, "I").End(-4162).Row, 2)
    
    if last_row_in >= 2:
        ws.Range(f"A1:F{max(last_row_in, 10)}").Borders.LineStyle = 1
        
    if last_row_out >= 2:
        ws.Range(f"H1:L{max(last_row_out, 10)}").Borders.LineStyle = 1
        
    # 4. Number formats
    ws.Columns("B:B").NumberFormat = "dd-mmm-yyyy"
    ws.Columns("I:I").NumberFormat = "dd-mmm-yyyy"
    ws.Columns("E:E").NumberFormat = "#,##0"
    ws.Columns("L:L").NumberFormat = "#,##0"
    
    # 5. Column Widths
    ws.Columns("A:F").AutoFit()
    ws.Columns("G:G").ColumnWidth = 3 # Spacer
    ws.Columns("H:L").AutoFit()
    
    # Minor adjustments
    ws.Columns("A:A").ColumnWidth = max(ws.Columns("A:A").ColumnWidth, 12)
    ws.Columns("C:C").ColumnWidth = max(ws.Columns("C:C").ColumnWidth, 15)
    ws.Columns("D:D").ColumnWidth = max(ws.Columns("D:D").ColumnWidth, 20)
    
    ws.Columns("H:H").ColumnWidth = max(ws.Columns("H:H").ColumnWidth, 12)
    ws.Columns("J:J").ColumnWidth = max(ws.Columns("J:J").ColumnWidth, 15)
    ws.Columns("K:K").ColumnWidth = max(ws.Columns("K:K").ColumnWidth, 20)

    wb.Save()
    print("SUCCESS: RENT PERSONS MISCELLANEOUS formatted")
except Exception as e:
    import traceback
    traceback.print_exc()
finally:
    wb.Close(False)
    xl.Quit()
