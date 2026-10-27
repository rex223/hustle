"""
Builds the MASTER_EXECUTIVE_SUMMARY sheet in 'RENTAL DATASHEET 19 OCT.xlsm'.
Features:
1. Executive KPI Summary Cards (Cumulative Inflow, Outflow, Net Saved, Savings %, Liabilities, Liquidity).
2. Month-by-Month Master Cash Flow Table (July 2026 to June 2027) with dynamic formulas.
3. Departmental Expenditure Breakdown & Excess Spend Anomaly Tracker (Salaries, CESC, Labour, Digrum, Feed/Fuel, Repairs, Purchases).
4. Money Saved & Cash Retention Scorecard (Variance, Retention per 1,000, Efficiency Index).
5. Executive Briefing & Actionable Recommendations.
"""

import win32com.client
import os
import sys

def build_master_sheet():
    excel_path = os.path.abspath('RENTAL DATASHEET 19 OCT.xlsm')
    if not os.path.exists(excel_path):
        print(f"Error: {excel_path} not found.")
        return False

    xl = win32com.client.Dispatch('Excel.Application')
    xl.DisplayAlerts = False
    xl.Visible = False
    xl.EnableEvents = False
    xl.ScreenUpdating = False

    try:
        wb = xl.Workbooks.Open(excel_path)
        xl.Calculation = -4135 # xlCalculationManual
        sheet_name = "MASTER_EXECUTIVE_SUMMARY"
        
        # Check if sheet exists
        for s in wb.Sheets:
            if s.Name.upper() == sheet_name.upper():
                s.Delete()
                print(f"Deleted existing {sheet_name} to rebuild cleanly.")
                break

        # Place right after TODAYS_EXPENDITURE
        after_sheet = None
        for s in wb.Sheets:
            if s.Name.upper() == "TODAYS_EXPENDITURE":
                after_sheet = s
                break
        
        if after_sheet is not None:
            ws = wb.Sheets.Add(None, after_sheet)
        else:
            ws = wb.Sheets.Add()
            
        ws.Name = sheet_name
        ws.Tab.Color = 0xAF401E # Executive deep blue (BGR)

        ws.Activate()
        xl.ActiveWindow.DisplayGridlines = True

        # Helper colors (RGB integers)
        NAVY = 0x33291E       # #1E2933 BGR: R=30, G=41, B=51
        HEADER_BLUE = 0x9B6A1A # BGR: R=26, G=106, B=155
        WHITE = 0xFFFFFF
        LIGHT_BG = 0xF8FAFC   # Slate 50
        BORDER_GRAY = 0xCBD5E1 # Slate 300

        # Set Column Widths
        col_widths = {
            'A': 3.0,
            'B': 18.0, # Month
            'C': 16.0, # Rent Collected
            'D': 14.0, # Purchases
            'E': 15.0, # Salaries Paid
            'F': 15.0, # Side Exp Paid
            'G': 13.0, # Misc Paid
            'H': 17.0, # Total Expenditure
            'I': 17.0, # Net Money Saved
            'J': 14.0, # Savings Rate %
            'K': 15.0, # Pending Dues
            'L': 16.0, # MoM Variance
            'M': 18.0, # Financial Status
            'N': 3.0
        }
        for col, width in col_widths.items():
            ws.Columns(col).ColumnWidth = width

        # Set standard font
        ws.Cells.Font.Name = "Segoe UI"
        ws.Cells.Font.Size = 9.5

        # =========================================================================
        # 1. TOP HEADER BANNER (Rows 1 to 3)
        # =========================================================================
        ws.Rows(1).RowHeight = 10
        ws.Rows(2).RowHeight = 36
        ws.Rows(3).RowHeight = 20
        ws.Rows(4).RowHeight = 10

        banner = ws.Range("B2:M2")
        banner.Merge()
        banner.Value = "EXECUTIVE FINANCIAL POSITION & MONTHLY CASH FLOW MASTER DASHBOARD"
        banner.Interior.Color = 0x33291E # Dark Navy
        banner.Font.Color = WHITE
        banner.Font.Bold = True
        banner.Font.Size = 13.5
        banner.HorizontalAlignment = -4108 # xlCenter
        banner.VerticalAlignment = -4108

        sub_banner = ws.Range("B3:M3")
        sub_banner.Merge()
        sub_banner.Value = "Consolidated Monthly Inflow, Expenditure Breakdown, Retained Cash Savings & Departmental Excess Spend Analytics"
        sub_banner.Interior.Color = 0x473B2D # Slate Navy
        sub_banner.Font.Color = 0xCBD5E1 # Slate 300
        sub_banner.Font.Size = 9.5
        sub_banner.HorizontalAlignment = -4108
        sub_banner.VerticalAlignment = -4108

        # =========================================================================
        # 2. EXECUTIVE KPI SUMMARY CARDS (Rows 5 to 7)
        # =========================================================================
        # Dynamic link directly to Total Row (Row 24)
        card_configs = [
            ("B", "C", "CUMULATIVE INFLOW", "=C24", "Gross Rental Revenue Collected", 0x9B6A1A, "#,##0"),
            ("D", "E", "CUMULATIVE EXPENDITURE", "=H24", "Purchases, Salaries & OpEx", 0x473B2D, "#,##0"),
            ("F", "G", "CUMULATIVE MONEY SAVED", "=I24", "Net Operational Retained Surplus", 0x1E7E34, "#,##0"),
            ("H", "I", "AVERAGE SAVINGS RATE", "=J24", "Retained Cash / Gross Inflow", 0x0078D4, "0.0%"),
            ("J", "K", "PENDING LIABILITIES", "=K24", "Unpaid Salaries & Supplier Dues", 0x2424B8, "#,##0"),
            ("L", "M", "NET LIQUIDITY POSITION", "=I24-K24", "Available Balance Less Total Dues", 0x155724, "#,##0")
        ]

        ws.Rows(5).RowHeight = 18
        ws.Rows(6).RowHeight = 28
        ws.Rows(7).RowHeight = 16
        ws.Rows(8).RowHeight = 8

        for col1, col2, title, formula, subtitle, hdr_color, num_fmt in card_configs:
            # Header
            rng_hdr = ws.Range(f"{col1}5:{col2}5")
            rng_hdr.Merge()
            rng_hdr.Value = title
            rng_hdr.Interior.Color = hdr_color
            rng_hdr.Font.Color = WHITE
            rng_hdr.Font.Bold = True
            rng_hdr.Font.Size = 8.5
            rng_hdr.HorizontalAlignment = -4108
            rng_hdr.VerticalAlignment = -4108

            # Value
            rng_val = ws.Range(f"{col1}6:{col2}6")
            rng_val.Merge()
            rng_val.Formula = formula
            rng_val.Interior.Color = WHITE
            rng_val.Font.Bold = True
            rng_val.Font.Size = 14
            rng_val.NumberFormat = num_fmt
            rng_val.HorizontalAlignment = -4108
            rng_val.VerticalAlignment = -4108

            # Subtitle
            rng_sub = ws.Range(f"{col1}7:{col2}7")
            rng_sub.Merge()
            rng_sub.Value = subtitle
            rng_sub.Interior.Color = 0xF1F5F9
            rng_sub.Font.Color = 0x64748B
            rng_sub.Font.Size = 7.5
            rng_sub.Font.Italic = True
            rng_sub.HorizontalAlignment = -4108
            rng_sub.VerticalAlignment = -4108

            card_box = ws.Range(f"{col1}5:{col2}7")
            card_box.Borders.Color = 0xCBD5E1

        # =========================================================================
        # 3. MONTHLY CASH FLOW MASTER TABLE (Rows 10 to 25)
        # =========================================================================
        ws.Rows(10).RowHeight = 24
        table_title = ws.Range("B10:M10")
        table_title.Merge()
        table_title.Value = "MONTHLY CASH FLOW & REVENUE ALLOCATION SUMMARY"
        table_title.Interior.Color = 0x33291E
        table_title.Font.Color = WHITE
        table_title.Font.Bold = True
        table_title.Font.Size = 11
        table_title.HorizontalAlignment = -4131 # xlLeft
        table_title.VerticalAlignment = -4108

        # Headers on Row 11
        ws.Rows(11).RowHeight = 30
        headers = [
            ("B", "MONTH / PERIOD"),
            ("C", "RENT COLLECTED\n(INFLOW)"),
            ("D", "PURCHASES\n(SUPPLIES)"),
            ("E", "SALARIES PAID\n(DISBURSED)"),
            ("F", "SIDE EXP PAID\n(CESC, LAB, ETC)"),
            ("G", "MISC PAID\n(OUTGOING)"),
            ("H", "TOTAL EXPENDITURE\n(OUTFLOW)"),
            ("I", "NET MONEY SAVED\n(BALANCE)"),
            ("J", "SAVINGS RATE\n(%)"),
            ("K", "PENDING DUES\n(LIABILITIES)"),
            ("L", "MOM VARIANCE\n(EXCESS SPEND)"),
            ("M", "EXECUTIVE STATUS\n(HEALTH)")
        ]
        for col, htext in headers:
            cell = ws.Range(f"{col}11")
            cell.Value = htext
            cell.WrapText = True
            cell.Interior.Color = 0x9B6A1A
            cell.Font.Color = WHITE
            cell.Font.Bold = True
            cell.Font.Size = 8.5
            cell.HorizontalAlignment = -4108
            cell.VerticalAlignment = -4108
            cell.Borders.Color = 0xCBD5E1

        # 12 Months Rows: Rows 12 to 23
        months_data = [
            ("JULY 2026", 829100, 114242, 291100, 441570, 0, 23300),
            ("AUGUST 2026", 828100, 7900, 260100, 450230, 0, 34500),
            ("SEPTEMBER 2026", 59500, 1100, 0, 0, 0, 0),
            ("OCTOBER 2026", 0, 0, 0, 0, 0, 0),
            ("NOVEMBER 2026", 0, 0, 0, 0, 0, 0),
            ("DECEMBER 2026", 0, 0, 0, 0, 0, 0),
            ("JANUARY 2027", 0, 0, 0, 0, 0, 0),
            ("FEBRUARY 2027", 0, 0, 0, 0, 0, 0),
            ("MARCH 2027", 0, 0, 0, 0, 0, 0),
            ("APRIL 2027", 0, 0, 0, 0, 0, 0),
            ("MAY 2027", 0, 0, 0, 0, 0, 0),
            ("JUNE 2027", 0, 0, 0, 0, 0, 0)
        ]

        for idx, (m_name, inf, purch, sal, side, misc, dues) in enumerate(months_data):
            r = 12 + idx
            ws.Rows(r).RowHeight = 20
            
            ws.Range(f"B{r}").Value = m_name
            ws.Range(f"B{r}").Font.Bold = True
            ws.Range(f"B{r}").HorizontalAlignment = -4131

            ws.Range(f"C{r}").Value = inf
            ws.Range(f"D{r}").Value = purch
            ws.Range(f"E{r}").Value = sal
            ws.Range(f"F{r}").Value = side
            ws.Range(f"G{r}").Value = misc
            ws.Range(f"K{r}").Value = dues

            # Dynamic Formulas
            ws.Range(f"H{r}").Formula = f"=SUM(D{r}:G{r})"
            ws.Range(f"I{r}").Formula = f"=C{r}-H{r}"
            ws.Range(f"J{r}").Formula = f"=IF(C{r}>0, I{r}/C{r}, 0)"
            
            if idx == 0:
                ws.Range(f"L{r}").Value = "-"
            else:
                prev_r = r - 1
                ws.Range(f"L{r}").Formula = f"=IF(AND(H{r}>0, H{prev_r}>0), H{r}-H{prev_r}, 0)"

            ws.Range(f"M{r}").Formula = f'=IF(C{r}=0, "⏳ UPCOMING", IF(I{r}<0, "🔴 DEFICIT", IF(I{r}>=100000, "🟢 HIGH SURPLUS", "🟢 SURPLUS")))'

            ws.Range(f"C{r}:I{r}").NumberFormat = "#,##0"
            ws.Range(f"J{r}").NumberFormat = "0.0%"
            ws.Range(f"K{r}").NumberFormat = "#,##0"
            ws.Range(f"L{r}").NumberFormat = "+#,##0;-#,##0;-"
            ws.Range(f"M{r}").HorizontalAlignment = -4108

            row_bg = WHITE if idx % 2 == 0 else 0xF8FAFC
            ws.Range(f"B{r}:M{r}").Interior.Color = row_bg
            ws.Range(f"B{r}:M{r}").Borders.Color = 0xE2E8F0

            if m_name == "AUGUST 2026":
                ws.Range(f"I{r}").Font.Color = 0x155724
                ws.Range(f"I{r}").Font.Bold = True
            elif m_name == "JULY 2026":
                ws.Range(f"I{r}").Font.Color = 0x1A1ACB
                ws.Range(f"I{r}").Font.Bold = True

        # Total Row (Row 24)
        ws.Rows(24).RowHeight = 22
        ws.Range("B24").Value = "YTD TOTALS"
        ws.Range("B24").Font.Bold = True
        ws.Range("C24").Formula = "=SUM(C12:C23)"
        ws.Range("D24").Formula = "=SUM(D12:D23)"
        ws.Range("E24").Formula = "=SUM(E12:E23)"
        ws.Range("F24").Formula = "=SUM(F12:F23)"
        ws.Range("G24").Formula = "=SUM(G12:G23)"
        ws.Range("H24").Formula = "=SUM(H12:H23)"
        ws.Range("I24").Formula = "=SUM(I12:I23)"
        ws.Range("J24").Formula = "=IF(C24>0, I24/C24, 0)"
        ws.Range("K24").Formula = "=SUM(K12:K23)"
        ws.Range("L24").Formula = "=SUM(L13:L23)"
        ws.Range("M24").Formula = '=IF(I24>0, "🟢 NET SURPLUS", "🔴 NET DEFICIT")'

        tot_rng = ws.Range("B24:M24")
        tot_rng.Font.Bold = True
        tot_rng.Interior.Color = 0xE2E8F0
        tot_rng.Borders.Color = 0x94A3B8
        ws.Range("C24:I24").NumberFormat = "#,##0"
        ws.Range("J24").NumberFormat = "0.0%"
        ws.Range("K24:L24").NumberFormat = "#,##0"
        ws.Range("M24").HorizontalAlignment = -4108

        # Average Row (Row 25)
        ws.Rows(25).RowHeight = 20
        ws.Range("B25").Value = "MONTHLY AVERAGE (ACTIVE)"
        ws.Range("B25").Font.Bold = True
        ws.Range("C25").Formula = "=AVERAGE(C12:C13)"
        ws.Range("D25").Formula = "=AVERAGE(D12:D13)"
        ws.Range("E25").Formula = "=AVERAGE(E12:E13)"
        ws.Range("F25").Formula = "=AVERAGE(F12:F13)"
        ws.Range("G25").Formula = "=AVERAGE(G12:G13)"
        ws.Range("H25").Formula = "=AVERAGE(H12:H13)"
        ws.Range("I25").Formula = "=AVERAGE(I12:I13)"
        ws.Range("J25").Formula = "=AVERAGE(J12:J13)"
        ws.Range("K25").Formula = "=AVERAGE(K12:K13)"
        ws.Range("L25").Value = "-"
        ws.Range("M25").Value = "BENCHMARK RUN-RATE"

        avg_rng = ws.Range("B25:M25")
        avg_rng.Interior.Color = 0xF1F5F9
        avg_rng.Font.Italic = True
        avg_rng.Borders.Color = 0xCBD5E1
        ws.Range("C25:I25").NumberFormat = "#,##0"
        ws.Range("J25").NumberFormat = "0.0%"
        ws.Range("K25").NumberFormat = "#,##0"
        ws.Range("M25").HorizontalAlignment = -4108

        # =========================================================================
        # 4. EXECUTIVE ANALYTICS 1: DEPARTMENTAL BREAKDOWN & EXCESS SPEND (Rows 28 to 37)
        # =========================================================================
        ws.Rows(27).RowHeight = 10
        ws.Rows(28).RowHeight = 24
        an1_title = ws.Range("B28:M28")
        an1_title.Merge()
        an1_title.Value = "EXECUTIVE ANALYTICS 1: DEPARTMENTAL EXPENDITURE BREAKDOWN & EXCESS SPEND ANOMALIES"
        an1_title.Interior.Color = 0x33291E
        an1_title.Font.Color = WHITE
        an1_title.Font.Bold = True
        an1_title.Font.Size = 10.5
        an1_title.HorizontalAlignment = -4131
        an1_title.VerticalAlignment = -4108

        ws.Rows(29).RowHeight = 26
        dept_headers = [
            ("B", "EXPENDITURE CATEGORY"),
            ("C", "JULY 2026\n(ACTUAL)"),
            ("D", "AUGUST 2026\n(ACTUAL)"),
            ("E", "MOM CHANGE\n(VARIANCE)"),
            ("F", "% CHANGE\n(GROWTH)"),
            ("G", "SHARE OF TOTAL\n(AUG %)"),
            ("H", "BUDGET IMPACT\n(CATEGORY STATUS)"),
            ("I", "M", "EXECUTIVE AUDIT FINDING & ACTIONABLE INSIGHT")
        ]
        for item in dept_headers:
            if len(item) == 2:
                col, htext = item
                cell = ws.Range(f"{col}29")
                cell.Value = htext
            else:
                c1, c2, htext = item
                cell = ws.Range(f"{c1}29:{c2}29")
                cell.Merge()
                cell.Value = htext

            cell.WrapText = True
            cell.Interior.Color = 0x473B2D
            cell.Font.Color = WHITE
            cell.Font.Bold = True
            cell.Font.Size = 8.5
            cell.HorizontalAlignment = -4108
            cell.VerticalAlignment = -4108
            cell.Borders.Color = 0xCBD5E1

        dept_data = [
            ("STAFF SALARIES & PAYROLL", 291100, 260100, "Controlled", "Payroll decreased by ₹31k (saving). 3 unpaid dues of ₹33.5k pending settlement."),
            ("CESC ELECTRICITY BILLS", 212200, 201530, "Stable", "Major utility cost stable at ~₹200k. Slight reduction of ₹10.6k (-5.0%) in August."),
            ("LABOUR & MAINTENANCE", 135270, 123800, "Controlled", "Ground labour spend reduced by ₹11.5k (-8.5%). Fully paid with zero pending dues."),
            ("DIGRUM & HAULAGE", 60000, 45000, "Controlled", "Haulage expenses reduced by ₹15k (-25.0%). Steady operation."),
            ("ANIMAL FEED & FUEL", 6500, 9000, "Slight Increase", "Routine feed/fuel fluctuation (+₹2.5k). Minor impact on cash flow (<1.5% of OpEx)."),
            ("MEDICAL & REPAIRS", 25000, 70900, "⚠️ EXCESS SPIKE", "CRITICAL EXCESS SPEND: Jumped +₹45,900 (+183.6%) driven by Kuddus (₹24.3k) & Hakim (₹10k)."),
            ("SUPPLY PURCHASES (STORE)", 114242, 7900, "🟢 MASSIVE SAVINGS", "EXCELLENT PROCUREMENT CONTROL: Slashed by ₹106,342 (-93.1%) enabling ₹109.8k net cash surplus!")
        ]

        for idx, (cat_name, j_val, a_val, status, audit_note) in enumerate(dept_data):
            r = 30 + idx
            ws.Rows(r).RowHeight = 22
            ws.Range(f"B{r}").Value = cat_name
            ws.Range(f"B{r}").Font.Bold = True
            ws.Range(f"C{r}").Value = j_val
            ws.Range(f"D{r}").Value = a_val
            
            ws.Range(f"E{r}").Formula = f"=D{r}-C{r}"
            ws.Range(f"F{r}").Formula = f"=IF(C{r}>0, (D{r}-C{r})/C{r}, 0)"
            ws.Range(f"G{r}").Formula = f"=D{r}/$D$37"
            ws.Range(f"H{r}").Value = status
            
            note_rng = ws.Range(f"I{r}:M{r}")
            note_rng.Merge()
            note_rng.Value = audit_note
            note_rng.HorizontalAlignment = -4131

            ws.Range(f"C{r}:E{r}").NumberFormat = "#,##0"
            ws.Range(f"F{r}:G{r}").NumberFormat = "0.0%"
            ws.Range(f"H{r}").HorizontalAlignment = -4108

            row_bg = WHITE if idx % 2 == 0 else 0xF8FAFC
            if "EXCESS SPIKE" in status:
                row_bg = 0xF0E6FF
                ws.Range(f"H{r}").Font.Color = 0x0000C8
                ws.Range(f"H{r}").Font.Bold = True
            elif "SAVINGS" in status:
                row_bg = 0xE6F9E6
                ws.Range(f"H{r}").Font.Color = 0x155724
                ws.Range(f"H{r}").Font.Bold = True

            ws.Range(f"B{r}:M{r}").Interior.Color = row_bg
            ws.Range(f"B{r}:M{r}").Borders.Color = 0xE2E8F0

        # Total Departmental Row (Row 37)
        ws.Rows(37).RowHeight = 22
        ws.Range("B37").Value = "TOTAL CONSOLIDATED OUTFLOW"
        ws.Range("B37").Font.Bold = True
        ws.Range("C37").Formula = "=SUM(C30:C36)"
        ws.Range("D37").Formula = "=SUM(D30:D36)"
        ws.Range("E37").Formula = "=D37-C37"
        ws.Range("F37").Formula = "=(D37-C37)/C37"
        ws.Range("G37").Formula = "=SUM(G30:G36)"
        ws.Range("H37").Formula = '=IF(E37<0, "🟢 COST REDUCED", "⚠️ COST EXPANDED")'
        
        sum_note = ws.Range("I37:M37")
        sum_note.Merge()
        sum_note.Value = "Net spending dropped by ₹128,682 (-15.2%), converting a ₹17.8k July deficit into a ₹109.8k August surplus."
        sum_note.Font.Bold = True
        sum_note.HorizontalAlignment = -4131

        ws.Range("B37:M37").Font.Bold = True
        ws.Range("B37:M37").Interior.Color = 0xE2E8F0
        ws.Range("C37:E37").NumberFormat = "#,##0"
        ws.Range("F37:G37").NumberFormat = "0.0%"
        ws.Range("H37").HorizontalAlignment = -4108

        # =========================================================================
        # 5. EXECUTIVE ANALYTICS 2: MONEY SAVED & CASH RETENTION (Rows 39 to 47)
        # =========================================================================
        ws.Rows(38).RowHeight = 10
        ws.Rows(39).RowHeight = 24
        an2_title = ws.Range("B39:M39")
        an2_title.Merge()
        an2_title.Value = "EXECUTIVE ANALYTICS 2: MONEY SAVED, CASH RETENTION & OPERATIONAL EFFICIENCY"
        an2_title.Interior.Color = 0x33291E
        an2_title.Font.Color = WHITE
        an2_title.Font.Bold = True
        an2_title.Font.Size = 10.5
        an2_title.HorizontalAlignment = -4131
        an2_title.VerticalAlignment = -4108

        ws.Rows(40).RowHeight = 26
        eff_headers = [
            ("B", "FINANCIAL METRIC"),
            ("C", "JULY 2026\n(ACTUAL)"),
            ("D", "AUGUST 2026\n(ACTUAL)"),
            ("E", "DELTA / SWING\n(IMPROVEMENT)"),
            ("F", "G", "TARGET / BENCHMARK"),
            ("H", "M", "EXECUTIVE ASSESSMENT & VALUE DELIVERED")
        ]
        for item in eff_headers:
            if len(item) == 2:
                col, htext = item
                cell = ws.Range(f"{col}40")
                cell.Value = htext
            else:
                c1, c2, htext = item
                cell = ws.Range(f"{c1}40:{c2}40")
                cell.Merge()
                cell.Value = htext

            cell.WrapText = True
            cell.Interior.Color = 0x473B2D
            cell.Font.Color = WHITE
            cell.Font.Bold = True
            cell.Font.Size = 8.5
            cell.HorizontalAlignment = -4108
            cell.VerticalAlignment = -4108
            cell.Borders.Color = 0xCBD5E1

        metrics_data = [
            ("Net Cash Saved / Retained", -17812, 109870, "=D41-C41", "Target: ≥ ₹80,000 / mo", 
             "MASSIVE POSITIVE SWING (+₹127,682). Business transitioned from capital drain to strong monthly cash generation."),
            ("Savings Retention Rate %", -0.0215, 0.1327, "=D42-C42", "Target: ≥ 10.0% of Inflow", 
             "Exceeded target by 3.27 percentage points in August, retaining 13.3% of total rental collections."),
            ("Retained Cash per ₹1,000 Collected", -21.48, 132.68, "=D43-C43", "Target: ₹100 per ₹1,000", 
             "For every ₹1,000 in rent received, ₹133 was converted into available net liquidity in August (vs -₹21 in July)."),
            ("Supply Procurement Savings", 114242, 7900, "=C44-D44", "Target: ≤ ₹25,000 / mo", 
             "Procurement costs slashed by ₹106,342. Eliminating bulk supply waste directly created 96.8% of August's retained profit."),
            ("Operating Cost to Revenue Ratio", 1.0215, 0.8673, "=D45-C45", "Target: ≤ 90.0% of Revenue", 
             "Operating efficiency improved dramatically from 102.1% (loss) to 86.7% (profitable operating window)."),
            ("Total Outstanding Liabilities (Dues)", 23300, 34500, "=D46-C46", "Limit: ≤ ₹40,000 / mo", 
             "Dues grew by ₹11,200 (₹33.5k unpaid salaries). Retained cash of ₹109.8k covers 100% of liabilities with ₹75.3k safety buffer.")
        ]

        for idx, (m_name, j_val, a_val, swing_formula, bench, desc) in enumerate(metrics_data):
            r = 41 + idx
            ws.Rows(r).RowHeight = 22
            ws.Range(f"B{r}").Value = m_name
            ws.Range(f"B{r}").Font.Bold = True
            ws.Range(f"C{r}").Value = j_val
            ws.Range(f"D{r}").Value = a_val
            ws.Range(f"E{r}").Formula = swing_formula
            
            b_rng = ws.Range(f"F{r}:G{r}")
            b_rng.Merge()
            b_rng.Value = bench
            b_rng.HorizontalAlignment = -4108

            d_rng = ws.Range(f"H{r}:M{r}")
            d_rng.Merge()
            d_rng.Value = desc
            d_rng.HorizontalAlignment = -4131

            if "%" in m_name or "Ratio" in m_name:
                ws.Range(f"C{r}:E{r}").NumberFormat = "0.0%"
            elif "per ₹1,000" in m_name:
                ws.Range(f"C{r}:E{r}").NumberFormat = "₹#,##0.00"
            else:
                ws.Range(f"C{r}:E{r}").NumberFormat = "#,##0"

            row_bg = WHITE if idx % 2 == 0 else 0xF8FAFC
            ws.Range(f"B{r}:M{r}").Interior.Color = row_bg
            ws.Range(f"B{r}:M{r}").Borders.Color = 0xE2E8F0

        # =========================================================================
        # 6. EXECUTIVE SUMMARY & KEY TAKEAWAYS (Rows 49 to 53)
        # =========================================================================
        ws.Rows(48).RowHeight = 10
        ws.Rows(49).RowHeight = 24
        rec_title = ws.Range("B49:M49")
        rec_title.Merge()
        rec_title.Value = "EXECUTIVE ACTION PLAN & SUMMARY TAKEAWAYS"
        rec_title.Interior.Color = 0x33291E
        rec_title.Font.Color = WHITE
        rec_title.Font.Bold = True
        rec_title.Font.Size = 10.5
        rec_title.HorizontalAlignment = -4131
        rec_title.VerticalAlignment = -4108

        takeaways = [
            ("1. REVENUE STABILITY", "Gross monthly rental collections are rock-solid at ~₹828,000 - ₹829,000 with strong tenant compliance across Ranjan, Mujib, Nehal, Pervaz, Azhar, Hyder, Moin, Rafat."),
            ("2. EXCESS SPEND ALERT", "Medical & Machine Repairs requires strict cap management: August incurred an unbudgeted surge of ₹45,900 (+183.6%), primarily from Kuddus (₹24,350) and Hakim Saheb (₹10,000). Pre-approval required for repairs > ₹5,000."),
            ("3. CAPITAL PRESERVATION", "August proved that limiting supply store purchases (reduced from ₹114,242 to ₹7,900) immediately unlocks over ₹100,000/month in net retained cash savings."),
            ("4. LIABILITY SETTLEMENT", "Current outstanding liabilities total ₹34,500 against available cash of ₹109,870. Immediate clearance of ₹33,500 unpaid payroll is recommended while preserving a ₹75,370 liquid buffer.")
        ]

        for idx, (head, text) in enumerate(takeaways):
            r = 50 + idx
            ws.Rows(r).RowHeight = 24
            h_cell = ws.Range(f"B{r}:C{r}")
            h_cell.Merge()
            h_cell.Value = head
            h_cell.Font.Bold = True
            h_cell.Interior.Color = 0xF1F5F9
            h_cell.HorizontalAlignment = -4131

            t_cell = ws.Range(f"D{r}:M{r}")
            t_cell.Merge()
            t_cell.Value = text
            t_cell.Interior.Color = WHITE
            t_cell.HorizontalAlignment = -4131

            ws.Range(f"B{r}:M{r}").Borders.Color = 0xE2E8F0

        # Bottom spacer
        ws.Rows(54).RowHeight = 16

        # Restore Excel application settings and calculate
        xl.Calculation = -4105 # xlCalculationAutomatic
        ws.Calculate()
        xl.EnableEvents = True
        xl.ScreenUpdating = True

        # Save workbook
        wb.Save()
        print("Workbook saved successfully with MASTER_EXECUTIVE_SUMMARY!")
        return True

    except Exception as e:
        print(f"Error during sheet construction: {e}")
        import traceback
        traceback.print_exc()
        return False
    finally:
        try:
            wb.Close(False)
        except:
            pass
        xl.Quit()

if __name__ == "__main__":
    success = build_master_sheet()
    if success:
        print("MASTER EXECUTIVE SHEET BUILD COMPLETE.")
    else:
        sys.exit(1)
