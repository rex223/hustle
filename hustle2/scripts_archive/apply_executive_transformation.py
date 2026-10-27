import win32com.client
import os

def rgb_to_int(r, g, b):
    return r + (g << 8) + (b << 16)

def run():
    file_path = os.path.abspath('RENTAL DATASHEET 19 OCT.xlsm')
    print(f"Applying Complete Executive Dashboard Transformation to: {file_path}")
    
    xl = win32com.client.Dispatch('Excel.Application')
    xl.DisplayAlerts = False
    xl.Visible = False
    xl.EnableEvents = False
    
    try:
        wb = xl.Workbooks.Open(file_path)
        
        # =========================================================================
        # PART 1: TODAYS_RENT_DATA (IMAGE 1)
        # =========================================================================
        print("Transforming todays_rent_data...")
        wsRent = wb.Sheets('todays_rent_data')
        
        # Color definitions for Groups (Clean, light executive pastel tints)
        grp1_tint = rgb_to_int(241, 245, 249)    # Slate / Ice 50
        grp1_subtotal = rgb_to_int(219, 234, 254)# Blue 100
        
        grp2_tint = rgb_to_int(240, 253, 244)    # Green / Mint 50
        grp2_subtotal = rgb_to_int(220, 252, 231)# Green 100
        
        grp3_tint = rgb_to_int(254, 252, 232)    # Amber / Sand 50
        grp3_subtotal = rgb_to_int(254, 240, 199)# Amber 100
        
        grp4_tint = rgb_to_int(245, 243, 255)    # Purple / Lavender 50
        grp4_subtotal = rgb_to_int(237, 233, 254)# Purple 100
        
        dark_slate = rgb_to_int(30, 41, 59)      # Slate 800
        navy_dark = rgb_to_int(15, 23, 42)       # Slate 900
        white_int = rgb_to_int(255, 255, 255)
        border_light = rgb_to_int(226, 232, 240) # Slate 200
        
        # 1. Header Row 1 (Cols A to J)
        wsRent.Rows(1).RowHeight = 28
        hdr_range = wsRent.Range("A1:J1")
        hdr_range.Interior.Color = dark_slate
        hdr_range.Font.Color = white_int
        hdr_range.Font.Bold = True
        hdr_range.Font.Size = 10
        hdr_range.Font.Name = "Segoe UI"
        hdr_range.HorizontalAlignment = -4108 # Center
        hdr_range.VerticalAlignment = -4108
        
        # 2. Iterate Rows 2 to 35 and apply Executive Group Colors
        # GRP 1: Rows 2 to 11 (Tenants), Row 12 (Subtotal)
        # GRP 2: Rows 13 to 23 (Tenants), Row 24 (Subtotal)
        # GRP 3: Rows 25 to 28 (Tenants), Row 29 (Subtotal)
        # GRP 4: Rows 30 to 33 (Tenants), Row 34 (Subtotal)
        # Grand Total: Row 35
        
        group_ranges = [
            (2, 11, grp1_tint, False),
            (12, 12, grp1_subtotal, True),
            (13, 23, grp2_tint, False),
            (24, 24, grp2_subtotal, True),
            (25, 28, grp3_tint, False),
            (29, 29, grp3_subtotal, True),
            (30, 33, grp4_tint, False),
            (34, 34, grp4_subtotal, True),
        ]
        
        for start_r, end_r, fill_color, is_subtotal in group_ranges:
            rng = wsRent.Range(wsRent.Cells(start_r, 1), wsRent.Cells(end_r, 8))
            rng.Interior.Color = fill_color
            rng.Font.Name = "Segoe UI"
            rng.Font.Size = 10
            rng.VerticalAlignment = -4108
            
            if is_subtotal:
                rng.Font.Bold = True
                rng.Font.Color = rgb_to_int(15, 23, 42)
                rng.Borders(8).LineStyle = 1 # Top border
                rng.Borders(8).Color = border_light
                rng.Borders(9).LineStyle = -4119 # xlDouble bottom border
                rng.Borders(9).Color = border_light
            else:
                rng.Font.Bold = False
                rng.Font.Color = 0
                for b_idx in [7, 8, 9, 10]:
                    rng.Borders(b_idx).LineStyle = 1
                    rng.Borders(b_idx).Color = border_light
                    rng.Borders(b_idx).Weight = 1
                    
        # Row 35: Grand Total
        gt_rng = wsRent.Range("A35:H35")
        gt_rng.Interior.Color = navy_dark
        gt_rng.Font.Color = white_int
        gt_rng.Font.Bold = True
        gt_rng.Font.Size = 11
        gt_rng.Font.Name = "Segoe UI"
        gt_rng.VerticalAlignment = -4108
        wsRent.Rows(35).RowHeight = 26
        
        # Row heights & alignments
        wsRent.Range("A2:A34").RowHeight = 22
        wsRent.Range("A2:A35").HorizontalAlignment = -4108
        wsRent.Range("A2:A35").NumberFormat = "dd-mmm-yyyy"
        wsRent.Range("B2:B35").HorizontalAlignment = -4131 # Left
        wsRent.Range("D2:D35").HorizontalAlignment = -4152 # Right
        wsRent.Range("D2:D35").NumberFormat = "#,##0"
        wsRent.Range("F2:F35").HorizontalAlignment = -4108 # Center
        
        # 3. Modernize Right-Hand Group Summary Table (Cols K to O)
        print("Styling Group Summary Table (Cols K to O)...")
        wsRent.Range("K1:O1").Interior.Color = rgb_to_int(30, 58, 138)
        wsRent.Range("K1:O1").Font.Color = white_int
        wsRent.Range("K1:O1").Font.Bold = True
        wsRent.Range("K1:O1").Font.Name = "Segoe UI"
        wsRent.Range("K1:O1").Font.Size = 10
        wsRent.Range("K1:O1").HorizontalAlignment = -4108
        wsRent.Range("K1:O1").VerticalAlignment = -4108
        
        wsRent.Range("K2:O5").RowHeight = 24
        wsRent.Range("K2:K5").Font.Bold = True
        wsRent.Range("K2:K5").Font.Name = "Segoe UI"
        wsRent.Range("K2:K5").HorizontalAlignment = -4108
        
        wsRent.Range("L2:L5").Font.Name = "Segoe UI"
        wsRent.Range("L2:L5").HorizontalAlignment = -4108
        wsRent.Range("L2:L5").NumberFormat = "dd-mmm-yyyy"
        
        # Col M: Total Amount (Soft Emerald)
        rng_m = wsRent.Range("M2:M5")
        rng_m.Interior.Color = rgb_to_int(240, 253, 244)
        rng_m.Font.Color = rgb_to_int(22, 101, 52)
        rng_m.Font.Bold = True
        rng_m.Font.Name = "Segoe UI"
        rng_m.Font.Size = 11
        rng_m.NumberFormat = "#,##0"
        rng_m.HorizontalAlignment = -4152
        
        # Col N: Total Paid (Soft Royal Blue)
        rng_n = wsRent.Range("N2:N5")
        rng_n.Interior.Color = rgb_to_int(239, 246, 255)
        rng_n.Font.Color = rgb_to_int(29, 78, 216)
        rng_n.Font.Bold = True
        rng_n.Font.Name = "Segoe UI"
        rng_n.Font.Size = 11
        rng_n.NumberFormat = "#,##0"
        rng_n.HorizontalAlignment = -4152
        
        # Col O: Group Due (Soft Rose / Red)
        rng_o = wsRent.Range("O2:O5")
        rng_o.Interior.Color = rgb_to_int(254, 242, 242)
        rng_o.Font.Color = rgb_to_int(185, 28, 28)
        rng_o.Font.Bold = True
        rng_o.Font.Name = "Segoe UI"
        rng_o.Font.Size = 11
        rng_o.NumberFormat = "#,##0"
        rng_o.HorizontalAlignment = -4152
        
        # Grand Total on Group Summary (Row 6)
        wsRent.Rows(6).RowHeight = 28
        gt_sum = wsRent.Range("K6:O6")
        gt_sum.Interior.Color = navy_dark
        gt_sum.Font.Color = white_int
        gt_sum.Font.Bold = True
        gt_sum.Font.Size = 12
        gt_sum.Font.Name = "Segoe UI"
        gt_sum.VerticalAlignment = -4108
        wsRent.Range("M6:O6").HorizontalAlignment = -4152
        wsRent.Range("M6:O6").NumberFormat = "#,##0"
        
        wsRent.Columns("A:J").AutoFit()
        wsRent.Columns("K:O").ColumnWidth = 16
        
        # =========================================================================
        # PART 2: TODAYS_EXPENDITURE & DATABASE_EXPENDITURE (IMAGE 2)
        # =========================================================================
        sections = [
            (1,  5,  "SALARIES & PAYROLL",          rgb_to_int(15, 23, 42)),   # Slate 900
            (6,  9,  "SALARY ADVANCES",             rgb_to_int(30, 58, 138)),  # Indigo 800
            (10, 13, "ELECTRICITY (CESC URP)",      rgb_to_int(180, 83, 9)),   # Amber 700
            (14, 17, "ELECTRICITY (CESC RESD)",     rgb_to_int(217, 119, 6)),  # Amber 600
            (18, 21, "LABOUR CONTRACTOR (URP)",     rgb_to_int(154, 52, 18)),  # Orange 800
            (22, 25, "DIGRUM WORKS (URP)",          rgb_to_int(194, 65, 12)),  # Orange 700
            (26, 29, "ANIMAL FEED",                 rgb_to_int(20, 83, 45)),   # Green 800
            (30, 33, "FUEL & LOGISTICS",            rgb_to_int(3, 105, 161)),  # Sky 700
            (34, 37, "MEDICAL & HEALTHCARE",        rgb_to_int(159, 18, 57)),  # Rose 800
            (38, 42, "FACILITIES & REPAIRS",        rgb_to_int(51, 65, 85)),   # Slate 600
            (43, 44, "COMMUNITY & RELIGION",        rgb_to_int(109, 40, 217))  # Purple 700
        ]
        
        separator_color = rgb_to_int(71, 85, 105) # Slate 600 heavy boundary
        
        for sheet_target in ['TODAYS_EXPENDITURE', 'database_expenditure']:
            print(f"Applying Section Theming and Visual Dividers to: {sheet_target}...")
            ws = wb.Sheets(sheet_target)
            ws.Rows(1).RowHeight = 28
            max_r = min(50, ws.UsedRange.Rows.Count)
            if max_r < 35: max_r = 35
            
            for start_c, end_c, sec_name, hdr_rgb in sections:
                # 1. Header Row 1 Styling for this section
                hdr_sec = ws.Range(ws.Cells(1, start_c), ws.Cells(1, end_c))
                hdr_sec.Interior.Color = hdr_rgb
                hdr_sec.Font.Color = white_int
                hdr_sec.Font.Bold = True
                hdr_sec.Font.Size = 10
                hdr_sec.Font.Name = "Segoe UI"
                hdr_sec.HorizontalAlignment = -4108 # Center
                hdr_sec.VerticalAlignment = -4108
                
                # 2. Heavy Right Border on end_c (1 single COM call per section!)
                border_rng = ws.Range(ws.Cells(1, end_c), ws.Cells(max_r, end_c))
                border_rng.Borders(10).LineStyle = 1
                border_rng.Borders(10).Weight = 3 # Thick boundary!
                border_rng.Borders(10).Color = separator_color
                
                # Format amounts and dates for this section
                for c in range(start_c, end_c + 1):
                    hdr_val = str(ws.Cells(1, c).Value or '').strip().upper()
                    if "AMOUNT" in hdr_val or "SALARIES" in hdr_val:
                        ws.Range(ws.Cells(2, c), ws.Cells(max_r, c)).NumberFormat = "#,##0"
                        ws.Range(ws.Cells(2, c), ws.Cells(max_r, c)).HorizontalAlignment = -4152
                    elif "DATE" in hdr_val:
                        ws.Range(ws.Cells(2, c), ws.Cells(max_r, c)).NumberFormat = "dd-mmm-yyyy"
                        ws.Range(ws.Cells(2, c), ws.Cells(max_r, c)).HorizontalAlignment = -4108
                    elif "PAID" in hdr_val or "STATUS" in hdr_val:
                        ws.Range(ws.Cells(2, c), ws.Cells(max_r, c)).HorizontalAlignment = -4108
                        ws.Range(ws.Cells(2, c), ws.Cells(max_r, c)).Font.Bold = True
                        
            # Generous column widths so no title is ever cut off
            ws.Columns(1).ColumnWidth = 24  # Name
            ws.Columns(2).ColumnWidth = 16  # Salary
            ws.Columns(3).ColumnWidth = 14  # Type
            ws.Columns(4).ColumnWidth = 15  # Date
            ws.Columns(5).ColumnWidth = 14  # Status
            ws.Columns(6).ColumnWidth = 14  # Adv Date
            ws.Columns(7).ColumnWidth = 16  # Adv Name
            ws.Columns(8).ColumnWidth = 15  # Adv Amount
            ws.Columns(9).ColumnWidth = 14  # Adv Paid
            ws.Columns(10).ColumnWidth = 14 # CESC Date
            ws.Columns(11).ColumnWidth = 22 # CESC URP Name
            ws.Columns(12).ColumnWidth = 15 # CESC URP Amount
            ws.Columns(13).ColumnWidth = 14 # CESC URP Paid
            ws.Columns(14).ColumnWidth = 14 # CESC RESD Date
            ws.Columns(15).ColumnWidth = 22 # CESC RESD Name
            ws.Columns(16).ColumnWidth = 15 # CESC RESD Amount
            ws.Columns(17).ColumnWidth = 14 # CESC RESD Paid
            ws.Columns(18).ColumnWidth = 14 # Labour Date
            ws.Columns(19).ColumnWidth = 24 # Labour URP
            ws.Columns(20).ColumnWidth = 15 # Labour Amount
            ws.Columns(21).ColumnWidth = 14 # Labour Paid
            ws.Columns(22).ColumnWidth = 14 # Digrum Date
            ws.Columns(23).ColumnWidth = 22 # Digrum URP
            ws.Columns(24).ColumnWidth = 15 # Digrum Amount
            ws.Columns(25).ColumnWidth = 14 # Digrum Paid
            ws.Columns(26).ColumnWidth = 20 # Animal Feed
            ws.Columns(27).ColumnWidth = 14 # Animal Date
            ws.Columns(28).ColumnWidth = 15 # Animal Amount
            ws.Columns(29).ColumnWidth = 14 # Animal Paid
            ws.Columns(30).ColumnWidth = 20 # Fuel/Gas
            ws.Columns(31).ColumnWidth = 14 # Fuel Date
            ws.Columns(32).ColumnWidth = 15 # Fuel Amount
            ws.Columns(33).ColumnWidth = 14 # Fuel Paid
            ws.Columns(34).ColumnWidth = 14 # Medical Date
            ws.Columns(35).ColumnWidth = 22 # Medical Name
            ws.Columns(36).ColumnWidth = 15 # Medical Amount
            ws.Columns(37).ColumnWidth = 14 # Medical Paid
            ws.Columns(38).ColumnWidth = 18 # Location
            ws.Columns(39).ColumnWidth = 14 # Repair Date
            ws.Columns(40).ColumnWidth = 26 # Repair Description
            ws.Columns(41).ColumnWidth = 15 # Repair Amount
            ws.Columns(42).ColumnWidth = 14 # Repair Paid
            ws.Columns(44).ColumnWidth = 20 # Religion
            
        wb.Save()
        print("Executive transformation applied and saved successfully!")
        
    finally:
        xl.EnableEvents = True
        wb.Close(False)
        xl.Quit()

if __name__ == '__main__':
    run()
