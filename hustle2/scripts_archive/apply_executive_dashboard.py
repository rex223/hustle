import win32com.client
import os

def rgb_to_int(r, g, b):
    return r + (g << 8) + (b << 16)

def run():
    file_path = os.path.abspath('RENTAL DATASHEET 19 OCT.xlsm')
    print(f"Applying Executive Dashboard Styling to: {file_path}")
    
    xl = win32com.client.Dispatch('Excel.Application')
    xl.DisplayAlerts = False
    xl.Visible = False
    xl.EnableEvents = False
    
    try:
        wb = xl.Workbooks.Open(file_path)
        wsExp = wb.Sheets('TODAYS_EXPENDITURE')
        
        # 1. Executive Top Header (Row 1 and Row 2)
        wsExp.Rows(1).RowHeight = 28
        wsExp.Rows(2).RowHeight = 24
        
        # Header style for Row 1
        navy_dark = rgb_to_int(15, 23, 42)    # Slate 900
        slate_hdr = rgb_to_int(30, 41, 59)    # Slate 800
        white_int = rgb_to_int(255, 255, 255)
        
        headers_col = ["NAME", "SALARIES", "TYPE / DEPT", "DATE", "PAID STATUS"]
        for c, h in enumerate(headers_col, start=1):
            cell = wsExp.Cells(1, c)
            cell.Value = h
            cell.Interior.Color = navy_dark
            cell.Font.Color = white_int
            cell.Font.Bold = True
            cell.Font.Size = 11
            cell.Font.Name = "Segoe UI"
            cell.HorizontalAlignment = -4108 # Center
            cell.VerticalAlignment = -4108   # Center
            
        # Month marker in Row 2
        for c in range(1, 6):
            cell = wsExp.Cells(2, c)
            cell.Interior.Color = slate_hdr
            cell.Font.Color = white_int
            cell.Font.Bold = True
            cell.Font.Size = 11
            cell.Font.Name = "Segoe UI"
            cell.HorizontalAlignment = -4108
            cell.VerticalAlignment = -4108
            
        # 2. Re-apply the 29 rows styling with Executive Spacing and Hairline Borders
        border_light = rgb_to_int(203, 213, 225) # Slate 300
        
        for r in range(3, 32):
            wsExp.Rows(r).RowHeight = 22
            # Set font & formatting for each column
            for c in range(1, 6):
                cell = wsExp.Cells(r, c)
                cell.Font.Name = "Segoe UI"
                cell.Font.Size = 10
                cell.VerticalAlignment = -4108
                
                # Hairline borders
                for b_idx in [7, 8, 9, 10]: # Left, Top, Bottom, Right
                    cell.Borders(b_idx).LineStyle = 1
                    cell.Borders(b_idx).Weight = 1 # Hairline
                    cell.Borders(b_idx).Color = border_light
            
            # Left align name with indent
            wsExp.Cells(r, 1).HorizontalAlignment = -4131 # Left
            wsExp.Cells(r, 1).InsertIndent(1)
            
            # Right align salaries with currency/accounting format
            wsExp.Cells(r, 2).HorizontalAlignment = -4152 # Right
            wsExp.Cells(r, 2).NumberFormat = "#,##0"
            
            # Center type, date, status
            wsExp.Cells(r, 3).HorizontalAlignment = -4108
            wsExp.Cells(r, 4).HorizontalAlignment = -4108
            wsExp.Cells(r, 4).NumberFormat = "dd-mmm-yyyy"
            wsExp.Cells(r, 5).HorizontalAlignment = -4108
            wsExp.Cells(r, 5).Font.Bold = True
            
        # 3. Executive TOTAL Row (Row 32)
        wsExp.Rows(32).RowHeight = 28
        crimson_red = rgb_to_int(185, 28, 28)
        
        wsExp.Cells(32, 1).Value = "TOTAL"
        wsExp.Cells(32, 1).HorizontalAlignment = -4108
        wsExp.Cells(32, 2).Formula = "=SUM(B3:B31)"
        wsExp.Cells(32, 2).HorizontalAlignment = -4152
        wsExp.Cells(32, 2).NumberFormat = "#,##0"
        
        for c in range(1, 6):
            cell = wsExp.Cells(32, c)
            cell.Interior.Color = crimson_red
            cell.Font.Color = white_int
            cell.Font.Bold = True
            cell.Font.Size = 12
            cell.Font.Name = "Segoe UI"
            cell.VerticalAlignment = -4108
            cell.Borders(9).LineStyle = -4119 # xlDouble bottom border
            cell.Borders(9).Color = white_int
            
        # 4. Executive KPI Dashboard Cards (Rows 35 to 39)
        wsExp.Rows(34).RowHeight = 15 # Breathing room
        wsExp.Rows(35).RowHeight = 26 # Section Banner
        wsExp.Rows(36).RowHeight = 24 # Card Headers
        wsExp.Rows(37).RowHeight = 32 # Card Big Numbers
        wsExp.Rows(38).RowHeight = 18 # Card Subtitle
        
        # Clean any old summary artifacts
        wsExp.Rows("35:45").Clear
        
        # Section Header Banner
        banner = wsExp.Range("A35:F35")
        banner.Merge()
        banner.Value = "FINANCIAL SUMMARY & EXECUTIVE CASH FLOW POSITION"
        banner.Interior.Color = navy_dark
        banner.Font.Color = white_int
        banner.Font.Bold = True
        banner.Font.Size = 12
        banner.Font.Name = "Segoe UI"
        banner.HorizontalAlignment = -4108
        banner.VerticalAlignment = -4108
        
        # Cards definition
        cards = [
            ("TOTAL RENT COLLECTED", rgb_to_int(30, 58, 138),  "TOTAL INFLOW"),          # Blue 800
            ("TOTAL PURCHASES",      rgb_to_int(71, 85, 105),  "SUPPLY EXPENSES"),      # Slate 600
            ("PAID AMOUNT",          rgb_to_int(37, 99, 235),  "DISBURSED SALARIES"),   # Blue 600
            ("MISC AMOUNT PAID",     rgb_to_int(185, 28, 28),  "OUTGOING MISC"),        # Red 700
            ("DUE AMOUNT",           rgb_to_int(180, 83, 9),   "PENDING PAYABLES"),     # Amber 700
            ("AVAILABLE BALANCE",    rgb_to_int(22, 101, 52),  "NET LIQUIDITY"),        # Green 800
        ]
        
        # Values for August statement (Strict calendar date matching for misc = 0 in August!)
        # Total Rent Collected: 828,100
        # Total Purchases: 7,900
        # Paid Salaries/Expenses: 549,810
        # Misc Outgoing: 0 (since Sept dates do not belong to August)
        # Due Amount: 675,150
        # Avail Balance = 828,100 - 7,900 - 549,810 - 0 = 270,390
        card_values = [
            828100.0,
            7900.0,
            549810.0,
            0.0,
            675150.0,
            270390.0
        ]
        
        card_bg = rgb_to_int(248, 250, 252) # Slate 50
        
        for col_idx, (hdr_text, hdr_color, sub_text) in enumerate(cards, start=1):
            # Row 36: Header
            h_cell = wsExp.Cells(36, col_idx)
            h_cell.Value = hdr_text
            h_cell.Interior.Color = hdr_color
            h_cell.Font.Color = white_int
            h_cell.Font.Bold = True
            h_cell.Font.Size = 10
            h_cell.Font.Name = "Segoe UI"
            h_cell.HorizontalAlignment = -4108
            h_cell.VerticalAlignment = -4108
            
            # Row 37: Big KPI Value
            v_cell = wsExp.Cells(37, col_idx)
            v_cell.Value = card_values[col_idx - 1]
            v_cell.Interior.Color = card_bg
            v_cell.Font.Bold = True
            v_cell.Font.Size = 15
            v_cell.Font.Name = "Segoe UI"
            v_cell.NumberFormat = "#,##0"
            v_cell.HorizontalAlignment = -4108
            v_cell.VerticalAlignment = -4108
            
            # Row 38: Sub-label
            s_cell = wsExp.Cells(38, col_idx)
            s_cell.Value = sub_text
            s_cell.Interior.Color = card_bg
            s_cell.Font.Color = rgb_to_int(100, 116, 139) # Slate 500
            s_cell.Font.Size = 8
            s_cell.Font.Bold = True
            s_cell.Font.Name = "Segoe UI"
            s_cell.HorizontalAlignment = -4108
            s_cell.VerticalAlignment = -4108
            
            # Specific value colors
            if col_idx == 4: # Misc Outgoing
                v_cell.Font.Color = rgb_to_int(185, 28, 28)
            elif col_idx == 5: # Due
                v_cell.Font.Color = rgb_to_int(180, 83, 9)
            elif col_idx == 6: # Avail Balance
                v_cell.Font.Color = rgb_to_int(22, 101, 52)
            else:
                v_cell.Font.Color = rgb_to_int(15, 23, 42)
                
            # Card borders
            for r_card in range(36, 39):
                for b_idx in [7, 8, 9, 10]:
                    wsExp.Cells(r_card, col_idx).Borders(b_idx).LineStyle = 1
                    wsExp.Cells(r_card, col_idx).Borders(b_idx).Color = border_light
                    wsExp.Cells(r_card, col_idx).Borders(b_idx).Weight = 1
                    
        # Column width optimization
        wsExp.Columns("A").ColumnWidth = 24
        wsExp.Columns("B").ColumnWidth = 16
        wsExp.Columns("C").ColumnWidth = 16
        wsExp.Columns("D").ColumnWidth = 18
        wsExp.Columns("E").ColumnWidth = 18
        wsExp.Columns("F").ColumnWidth = 22
        
        # 5. Also ensure database_expenditure August block matches this exact reorganized order
        wsDB = wb.Sheets('database_expenditure')
        aug_start = 0
        for r in range(1, wsDB.UsedRange.Rows.Count + 1):
            if str(wsDB.Cells(r, 1).Value or '').strip().upper() == 'AUGUST_2026':
                aug_start = r
                break
                
        if aug_start > 0:
            print(f"Updating database_expenditure AUGUST_2026 block starting at row {aug_start}...")
            # Copy Columns A-E from rows 2 to 32 of TODAYS_EXPENDITURE to database_expenditure
            for offset in range(0, 31): # 0 is row 2, 30 is row 32
                for c in range(1, 6):
                    wsDB.Cells(aug_start + offset, c).Value = wsExp.Cells(2 + offset, c).Value
                    wsDB.Cells(aug_start + offset, c).Interior.Color = wsExp.Cells(2 + offset, c).Interior.Color
                    wsDB.Cells(aug_start + offset, c).Font.Color = wsExp.Cells(2 + offset, c).Font.Color
                    wsDB.Cells(aug_start + offset, c).Font.Bold = wsExp.Cells(2 + offset, c).Font.Bold
                    wsDB.Cells(aug_start + offset, c).NumberFormat = wsExp.Cells(2 + offset, c).NumberFormat
                    
        wb.Save()
        print("Workbook saved with Executive Dashboard styling successfully!")
        
    finally:
        xl.EnableEvents = True
        wb.Close(False)
        xl.Quit()

if __name__ == '__main__':
    run()
