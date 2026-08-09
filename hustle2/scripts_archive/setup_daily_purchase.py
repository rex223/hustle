import win32com.client
import os

excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False
file_path = os.path.abspath('RENTAL DATASHEET 1ST AUG.xlsm')

try:
    wb = excel.Workbooks.Open(file_path)
    
    # 1. Create or clear 'PURCHASE SHEET DAILY'
    sheet_name = "PURCHASE SHEET DAILY"
    ws_new = None
    
    # Find existing
    for s in wb.Sheets:
        if s.Name.upper() == sheet_name.upper():
            ws_new = s
            break
            
    if ws_new:
        ws_new.Cells.Clear()
    else:
        # Add after PURCHASE if possible
        target_index = wb.Sheets.Count
        for i in range(1, wb.Sheets.Count + 1):
            if wb.Sheets(i).Name.upper() == "PURCHASE":
                target_index = i
                break
        ws_new = wb.Sheets.Add(After=wb.Sheets(target_index))
        ws_new.Name = sheet_name
        
    ws_old = wb.Sheets("PURCHASE")
    
    # 2. Extract Data from old sheet
    last_row = ws_old.Cells(ws_old.Rows.Count, "A").End(3).Row
    if last_row < 2:
        last_row = ws_old.Cells(ws_old.Rows.Count, "C").End(3).Row
        
    common_date = None
    data_rows = []
    
    for r in range(2, last_row + 1):
        # Read the row
        row_data = [ws_old.Cells(r, c).Value for c in range(1, 9)]
        
        # Check if row is entirely empty
        if all(v is None or str(v).strip() == "" for v in row_data):
            continue
            
        # Capture the common date if it exists
        if row_data[0] is not None and str(row_data[0]).strip() != "":
            if not common_date:
                common_date = row_data[0]
        
        # Add to our list
        data_rows.append(row_data)

    # 3. Setup New Headers
    headers = ["DATE", "GROUP / FOR", "ITEM NAME", "QTY (KG)", "QTY (PCS)", "RATE", "TRSP", "TOTAL AMOUNT"]
    ws_new.Range("A1:H1").Value = headers
    
    # 4. Write data to new sheet, filling in the date
    out_row = 2
    for row_data in data_rows:
        # If there's no date for this row, use the common_date
        date_val = row_data[0]
        if date_val is None or str(date_val).strip() == "":
            date_val = common_date
            
        ws_new.Cells(out_row, 1).Value = date_val
        ws_new.Cells(out_row, 2).Value = row_data[1] # GROUP / FOR
        ws_new.Cells(out_row, 3).Value = row_data[2] # ITEM
        ws_new.Cells(out_row, 4).Value = row_data[3] # QTY KG
        ws_new.Cells(out_row, 5).Value = row_data[4] # QTY PCS
        ws_new.Cells(out_row, 6).Value = row_data[5] # RATE
        ws_new.Cells(out_row, 7).Value = row_data[6] # TRSP
        ws_new.Cells(out_row, 8).Value = row_data[7] # TOTAL AMOUNT (User manually enters)
        
        out_row += 1

    # 5. Formatting as a Table (ListObject)
    # Check if table already exists and delete it
    for tbl in ws_new.ListObjects:
        tbl.Delete()
        
    table_range = f"A1:H{out_row - 1 if out_row > 2 else 2}"
    tbl = ws_new.ListObjects.Add(1, ws_new.Range(table_range), 1, 1) # xlSrcRange, Range, xlYes
    tbl.Name = "DailyPurchaseLedger"
    tbl.TableStyle = "TableStyleMedium9" # Nice blue style
    
    # Adjust widths
    ws_new.Columns("A").ColumnWidth = 12
    ws_new.Columns("B:C").ColumnWidth = 20
    ws_new.Columns("D:H").ColumnWidth = 12
    
    # Date Format
    ws_new.Columns("A").NumberFormat = "dd-mmm-yyyy"
    
    wb.Save()
    print("SUCCESS")
except Exception as e:
    print(f"Error: {e}")
finally:
    try:
        wb.Close(False)
    except:
        pass
    excel.Quit()
