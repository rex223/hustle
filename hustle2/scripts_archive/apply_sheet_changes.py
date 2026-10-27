import win32com.client
import os

def rgb_to_int(r, g, b):
    return r + (g << 8) + (b << 16)

def run():
    file_path = os.path.abspath('RENTAL DATASHEET 19 OCT.xlsm')
    print(f"Opening workbook: {file_path}")
    
    xl = win32com.client.Dispatch('Excel.Application')
    xl.DisplayAlerts = False
    xl.Visible = False
    xl.EnableEvents = False  # Prevent old macros from firing while we format
    
    try:
        wb = xl.Workbooks.Open(file_path)
        wsExp = wb.Sheets('TODAYS_EXPENDITURE')
        
        # 1. Capture current Rows 3 to 31 data in Columns A to E
        records = []
        for r in range(3, 32):
            row_data = {
                'name': wsExp.Cells(r, 1).Value,
                'salary': wsExp.Cells(r, 2).Value,
                'type': wsExp.Cells(r, 3).Value,
                'date': wsExp.Cells(r, 4).Value,
                'status': wsExp.Cells(r, 5).Value,
            }
            records.append(row_data)
        
        # Helper to find record
        def pop_record(name_match, type_match=None):
            for i, rec in enumerate(records):
                n = str(rec['name'] or '').strip().upper()
                t = str(rec['type'] or '').strip().upper()
                if n == name_match.upper():
                    if type_match is None or t == type_match.upper():
                        return records.pop(i)
            for i, rec in enumerate(records):
                n = str(rec['name'] or '').strip().upper()
                if n == name_match.upper():
                    return records.pop(i)
            return None

        # Build new ordered list of 29 items
        new_order = [
            # 1. Bawa on first row
            ('BAWA', None, rgb_to_int(220, 230, 242)),
            # 2-13. Staff (Preexisting colors preserved!)
            ('NAIM', None, rgb_to_int(253, 233, 217)),
            ('NIZAM', None, rgb_to_int(253, 233, 217)),
            ('SHAFIQUE', None, rgb_to_int(253, 233, 217)),
            ('KAMAL', None, rgb_to_int(253, 233, 217)),
            ('NAZIR', None, rgb_to_int(146, 205, 220)),
            ('ZYED', None, rgb_to_int(146, 205, 220)),
            ('ISHRAQUE', None, rgb_to_int(255, 255, 0)),
            ('MERAB', None, rgb_to_int(255, 255, 0)),
            ('ENAYATULLA', None, rgb_to_int(255, 255, 0)),
            ('RASHEED', None, rgb_to_int(196, 189, 151)),
            ('MUNNA', None, rgb_to_int(242, 220, 219)),
            ('RAFIQUE', None, rgb_to_int(242, 220, 219)),
            # 14-15. Security / Site (Preexisting orange preserved!)
            ('SURAJ', None, rgb_to_int(250, 191, 143)),
            ('SECURITY', None, rgb_to_int(250, 191, 143)),
            # 16-17. Residential (Soft lavender group tint)
            ('SAM', None, rgb_to_int(232, 228, 242)),
            ('ZOE/ZARA', None, rgb_to_int(232, 228, 242)),
            # 18-19. Statutory (Soft mint/sage group tint)
            ('ESI', None, rgb_to_int(225, 240, 235)),
            ('NARAYAN', None, rgb_to_int(225, 240, 235)),
            # 20-24. Weekly wages (Soft warm wheat group tint)
            ('SHAHNAWAZ/108033', 'WEEK 1', rgb_to_int(255, 243, 205)),
            ('SHAHNAWAZ/108033', 'WEEK 2', rgb_to_int(255, 243, 205)),
            ('SHAHNAWAZ/108033', 'WEEK 3', rgb_to_int(255, 243, 205)),
            ('SHAHNAWAZ/108033', 'WEEK 4', rgb_to_int(255, 243, 205)),
            ('SHAHNAWAZ/108033', 'WEEK 5', rgb_to_int(255, 243, 205)),
            # 25. Utilities / Site bills (Soft slate group tint)
            ('CESC+RAJA+OTHERS', None, rgb_to_int(235, 238, 245)),
            # 26-28. Leadership: Adinul & Akhtar above Parvez (Soft pistachio group tint)
            ('ADINUL', None, rgb_to_int(226, 239, 218)),
            ('AKHTAR', None, rgb_to_int(226, 239, 218)),
            ('PARVEZ', None, rgb_to_int(226, 239, 218)),
            # 29. EC at the very bottom (Soft lilac/accent group tint)
            ('EC', None, rgb_to_int(245, 230, 240)),
        ]
        
        print(f"Placing {len(new_order)} ordered rows into TODAYS_EXPENDITURE (Rows 3 to 31)...")
        
        curr_row = 3
        total_salary = 0
        for name_key, type_key, fill_color in new_order:
            rec = pop_record(name_key, type_key)
            if rec is None:
                print(f"Warning: could not find record for {name_key} ({type_key})")
                continue
            
            wsExp.Cells(curr_row, 1).Value = rec['name']
            wsExp.Cells(curr_row, 2).Value = rec['salary']
            wsExp.Cells(curr_row, 3).Value = rec['type']
            wsExp.Cells(curr_row, 4).Value = rec['date']
            wsExp.Cells(curr_row, 5).Value = rec['status']
            
            if rec['salary'] is not None:
                total_salary += float(rec['salary'])
                
            for c in range(1, 6):
                wsExp.Cells(curr_row, c).Interior.Color = fill_color
                wsExp.Cells(curr_row, c).Font.Color = 0
                wsExp.Cells(curr_row, c).Font.Bold = False
            
            curr_row += 1
            
        print(f"All 29 rows written. Total Salary check: {total_salary}")
        
        # Row 32: TOTAL Row
        wsExp.Cells(32, 1).Value = "TOTAL"
        wsExp.Cells(32, 2).Formula = "=SUM(B3:B31)"
        wsExp.Cells(32, 3).Value = None
        wsExp.Cells(32, 4).Value = None
        wsExp.Cells(32, 5).Value = None
        
        red_int = rgb_to_int(198, 40, 40)
        white_int = rgb_to_int(255, 255, 255)
        for c in range(1, 6):
            wsExp.Cells(32, c).Interior.Color = red_int
            wsExp.Cells(32, c).Font.Color = white_int
            wsExp.Cells(32, c).Font.Bold = True
        wsExp.Cells(32, 2).NumberFormat = "#,##0"
        
        # 2. Update the Summary Table Card at Rows 36-37
        print("Formatting Summary Table card at Rows 36-37...")
        summary_headers = [
            "TOTAL RENT COLLECTED",
            "TOTAL PURCHASES",
            "PAID AMOUNT",
            "MISCELLANEOUS AMOUNT PAID",
            "DUE AMOUNT",
            "AVAILABLE BALANCE"
        ]
        
        blue_hdr = rgb_to_int(41, 128, 185)
        for idx, h_text in enumerate(summary_headers, start=1):
            cell = wsExp.Cells(36, idx)
            cell.Value = h_text
            cell.Interior.Color = blue_hdr
            cell.Font.Color = white_int
            cell.Font.Bold = True
            cell.HorizontalAlignment = -4108
            
        wsExp.Cells(36, 7).Value = None
        wsExp.Cells(36, 7).Interior.ColorIndex = -4142
        wsExp.Cells(37, 7).Value = None
        wsExp.Cells(37, 7).Interior.ColorIndex = -4142
        
        total_rent = 828100.0
        total_purch = 7900.0
        total_sal_paid = 549810.0
        total_misc = 72800.0
        total_due = 675150.0
        avail_bal = total_rent - total_purch - total_sal_paid - total_misc
        
        wsExp.Cells(37, 1).Value = total_rent
        wsExp.Cells(37, 2).Value = total_purch
        wsExp.Cells(37, 3).Value = total_sal_paid
        wsExp.Cells(37, 4).Value = -total_misc
        wsExp.Cells(37, 5).Value = total_due
        wsExp.Cells(37, 6).Value = avail_bal
        
        dark_red = rgb_to_int(192, 0, 0)
        green_int = rgb_to_int(0, 128, 0)
        
        for c in range(1, 7):
            cell = wsExp.Cells(37, c)
            cell.Font.Bold = True
            cell.HorizontalAlignment = -4108
            cell.NumberFormat = "#,##0"
            if c == 4:
                cell.Font.Color = dark_red
            elif c == 6:
                cell.Font.Color = dark_red if avail_bal < total_due else green_int
            else:
                cell.Font.Color = 0
                
        wsExp.Columns("A:F").AutoFit()
        
        # 3. Handle Kamal Sheet: clear unbilled template row 4
        if 'KAMAL' in [s.Name for s in wb.Sheets]:
            wsKamal = wb.Sheets('KAMAL')
            val_r4 = wsKamal.Cells(4, 4).Value
            if str(val_r4 or '').strip().upper() == 'KAMAL' and wsKamal.Cells(4, 1).Value is None:
                print("Clearing unbilled template row 4 from KAMAL sheet...")
                wsKamal.Rows(4).ClearContents()
                
        wb.Save()
        print("Workbook saved successfully with 6-column layout!")
        
    finally:
        xl.EnableEvents = True
        wb.Close(False)
        xl.Quit()

if __name__ == '__main__':
    run()
