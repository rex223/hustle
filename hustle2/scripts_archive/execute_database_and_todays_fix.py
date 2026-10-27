import win32com.client
import os
import datetime

def excel_serial_to_num(val):
    if hasattr(val, 'year'): # datetime
        dt = datetime.date(val.year, val.month, val.day)
        base = datetime.date(1899, 12, 30)
        return float((dt - base).days)
    return float(val) if val is not None else None

def to_pywintypes_date(d):
    if d is None: return None
    if isinstance(d, datetime.datetime): return d
    if isinstance(d, datetime.date): return datetime.datetime(d.year, d.month, d.day)
    return d

def main():
    xl = win32com.client.Dispatch('Excel.Application')
    xl.DisplayAlerts = False
    xl.Visible = False

    wb_path = os.path.abspath('RENTAL DATASHEET 19 OCT.xlsm')
    wb = xl.Workbooks.Open(wb_path)

    wsDB = wb.Sheets('database_expenditure')
    wsToday = wb.Sheets('TODAYS_EXPENDITURE')

    # Standardized 44 Headers
    headers = [
        "NAME", "SALARIES", "TYPE", "DATE", "PAID STATUS",
        "DATE", "ADV", "AMOUNT", "PAID",
        "DATE", "CESC / URP", "AMOUNT", "PAID",
        "DATE", "CESC / RESD", "AMOUNT", "PAID",
        "DATE", "LABOUR/URP", "AMOUNT", "PAID",
        "DATE", "DIGRUM / URP", "AMOUNT", "PAID",
        "DATE", "ANIMAL FEED", "AMOUNT", "PAID",
        "DATE", "FUEL/GAS", "AMOUNT", "PAID",
        "DATE", "MEDICAL", "AMOUNT", "PAID",
        "DATE", "LOCATION", "REPAIR", "AMOUNT", "PAID",
        "RELIGION", "AMOUNT"
    ]

    print("=== STEP 1: READ JULY & AUGUST RAW DATA ===")
    personnel_order = [
        "BAWA", "NAIM", "NIZAM", "SHAFIQUE", "KAMAL", "NAZIR", "ZYED", "ISHRAQUE",
        "MERAB", "ENAYATULLA", "RASHEED", "MUNNA", "RAFIQUE", "SURAJ ", "SECURITY",
        "SAM", "SAM 2", "ZOE/ZARA", "ESI", "NARAYAN",
        "SHAHNAWAZ/108033_1", "SHAHNAWAZ/108033_2", "SHAHNAWAZ/108033_3", "SHAHNAWAZ/108033_4", "SHAHNAWAZ/108033_5",
        "ADINUL", "AKHTAR", "PARVEZ", "CESC+RAJA+OTHERS", "EC", "TOTAL"
    ]

    # --- JULY DATA ---
    july_people = {}
    shahnawaz_count = 0
    for r in range(3, 34):
        p_name = wsDB.Cells(r, 1).Value
        if not p_name: continue
        p_name = str(p_name).strip()
        key = p_name
        if "SHAHNAWAZ" in p_name.upper():
            shahnawaz_count += 1
            key = f"SHAHNAWAZ/108033_{shahnawaz_count}"
        
        sal = wsDB.Cells(r, 2).Value
        ptype = wsDB.Cells(r, 3).Value
        pdate = wsDB.Cells(r, 4).Value
        pstatus = wsDB.Cells(r, 5).Value
        july_people[key] = (p_name, sal, ptype, pdate, pstatus)

    july_side = []
    for r in range(3, 34):
        adv_date = wsToday.Cells(r, 6).Value
        adv_name = wsToday.Cells(r, 7).Value
        adv_amt = wsToday.Cells(r, 8).Value
        adv_paid = wsToday.Cells(r, 9).Value

        urp_date = wsToday.Cells(r, 10).Value
        urp_name = wsToday.Cells(r, 11).Value
        urp_amt = wsToday.Cells(r, 12).Value
        urp_paid = wsToday.Cells(r, 13).Value

        resd_date = wsToday.Cells(r, 14).Value
        resd_name = wsToday.Cells(r, 15).Value
        resd_amt = wsToday.Cells(r, 16).Value
        resd_paid = wsToday.Cells(r, 17).Value

        lab_date = wsToday.Cells(r, 18).Value
        lab_name = wsToday.Cells(r, 19).Value
        lab_amt = wsToday.Cells(r, 20).Value
        lab_paid = wsToday.Cells(r, 21).Value

        dig_date = wsToday.Cells(r, 22).Value
        dig_name = wsToday.Cells(r, 23).Value
        dig_amt = wsToday.Cells(r, 24).Value
        dig_paid = wsToday.Cells(r, 25).Value

        raw_af_item = wsToday.Cells(r, 26).Value
        raw_af_date = wsToday.Cells(r, 27).Value
        raw_af_amt = wsToday.Cells(r, 28).Value
        raw_af_paid = wsToday.Cells(r, 29).Value
        af_date = raw_af_date
        af_item = raw_af_item
        af_amt = raw_af_amt
        af_paid = raw_af_paid

        raw_fuel_item = wsToday.Cells(r, 30).Value
        raw_fuel_date = wsToday.Cells(r, 31).Value
        raw_fuel_amt = wsToday.Cells(r, 32).Value
        raw_fuel_paid = wsToday.Cells(r, 33).Value
        fuel_date = raw_fuel_date
        fuel_item = raw_fuel_item
        fuel_amt = raw_fuel_amt
        fuel_paid = raw_fuel_paid

        med_date = wsToday.Cells(r, 34).Value
        med_doc = wsToday.Cells(r, 35).Value
        med_amt = wsToday.Cells(r, 36).Value
        med_paid = wsToday.Cells(r, 37).Value

        raw_rep_38 = wsToday.Cells(r, 38).Value
        raw_rep_39 = wsToday.Cells(r, 39).Value
        raw_rep_40 = wsToday.Cells(r, 40).Value
        raw_rep_41 = wsToday.Cells(r, 41).Value
        raw_rep_42 = wsToday.Cells(r, 42).Value

        if hasattr(raw_rep_38, 'year'):
            rep_date = raw_rep_38
            rep_loc = raw_rep_39
        elif hasattr(raw_rep_39, 'year'):
            rep_date = raw_rep_39
            rep_loc = raw_rep_38
        else:
            rep_date = raw_rep_38
            rep_loc = raw_rep_39
        rep_item = raw_rep_40
        rep_amt = raw_rep_41
        rep_paid = raw_rep_42

        rel_amt = wsToday.Cells(r, 43).Value
        rel_name = wsToday.Cells(r, 44).Value

        july_side.append({
            'adv': (adv_date, adv_name, adv_amt, adv_paid),
            'urp': (urp_date, urp_name, urp_amt, urp_paid),
            'resd': (resd_date, resd_name, resd_amt, resd_paid),
            'lab': (lab_date, lab_name, lab_amt, lab_paid),
            'dig': (dig_date, dig_name, dig_amt, dig_paid),
            'af': (af_date, af_item, af_amt, af_paid),
            'fuel': (fuel_date, fuel_item, fuel_amt, fuel_paid),
            'med': (med_date, med_doc, med_amt, med_paid),
            'rep': (rep_date, rep_loc, rep_item, rep_amt, rep_paid),
            'rel': (rel_name, rel_amt)
        })

    # --- AUGUST DATA ---
    august_people = {}
    shahnawaz_count = 0
    for r in range(35, 65):
        p_name = wsDB.Cells(r, 1).Value
        if not p_name: continue
        p_name = str(p_name).strip()
        key = p_name
        if "SHAHNAWAZ" in p_name.upper():
            shahnawaz_count += 1
            key = f"SHAHNAWAZ/108033_{shahnawaz_count}"
        
        sal = wsDB.Cells(r, 2).Value
        ptype = wsDB.Cells(r, 3).Value
        pdate = wsDB.Cells(r, 4).Value
        pstatus = wsDB.Cells(r, 5).Value
        august_people[key] = (p_name, sal, ptype, pdate, pstatus)

    august_side = [
        # NAIM (idx 0)
        {'adv': (None, None, None, None),
         'urp': (datetime.date(2026, 9, 15), "M/P", 38320.0, "PAID"),
         'resd': (datetime.date(2026, 9, 15), "SHAMIMA", 4900.0, "PAID"),
         'lab': (datetime.date(2026, 9, 5), "WEEKLY", 32050.0, "PAID"),
         'dig': (None, "PASHA", 15000.0, "PAID"),
         'af': (datetime.date(2026, 9, 23), "BECHALI", 3200.0, "PAID"),
         'fuel': (None, "SHAMIMA", None, None),
         'med': (None, "SAM / DR", None, None),
         'rep': (None, None, "SHED / MISTRI", None, None),
         'rel': ("ZD", None)},

        # NIZAM (idx 1)
        {'adv': (None, None, None, None),
         'urp': (datetime.date(2026, 9, 15), "NAUSHAD ALAM", 92093.0, "PAID"),
         'resd': (datetime.date(2026, 9, 15), "KANIZ", 17919.0, "PAID"),
         'lab': (datetime.date(2026, 9, 12), "WEEKLY", 28200.0, "PAID"),
         'dig': (datetime.date(2026, 9, 12), "PASHA", 15000.0, "PAID"),
         'af': (datetime.date(2026, 9, 5), "CHOKHAR", 4800.0, "PAID"),
         'fuel': (None, "ZOE", None, None),
         'med': (None, "SAM/TEST", None, None),
         'rep': (None, None, "TIN ", None, None),
         'rel': ("SD", None)},

        # SHAFIQUE (idx 2)
        {'adv': (None, None, None, None),
         'urp': (datetime.date(2026, 9, 15), "SK BHALLU", 13990.0, "PAID"),
         'resd': (datetime.date(2026, 9, 15), "SULTANA", 19448.0, "PAID"),
         'lab': (datetime.date(2026, 9, 19), "WEEKLY", 29750.0, "PAID"),
         'dig': (None, "PASHA", None, None),
         'af': (None, None, None, None),
         'fuel': (None, "ZARA", None, None),
         'med': (None, "SAM /  MED", None, None),
         'rep': (datetime.date(2026, 9, 4), "PPL", "SHED", 1500.0, "PAID"),
         'rel': ("KD", None)},

        # KAMAL (idx 3)
        {'adv': (None, None, None, None),
         'urp': (datetime.date(2026, 9, 15), "SULTAN", 6020.0, "PAID"),
         'resd': (datetime.date(2026, 9, 15), "MASJID", 8740.0, "PAID"),
         'lab': (datetime.date(2026, 9, 26), "WEEKLY", 33800.0, "PAID"),
         'dig': (datetime.date(2026, 9, 26), "PASHA", 15000.0, "PAID"),
         'af': (None, None, None, None),
         'fuel': (None, "SULTAN", None, None),
         'med': (None, "HAKIM SAHEB", None, None),
         'rep': (datetime.date(2026, 9, 7), None, "KUDDUS", 24350.0, "PAID"),
         'rel': ("AT", None)},

        # NAZIR (idx 4)
        {'adv': (None, None, None, None),
         'urp': (datetime.date(2026, 9, 15), "JAVED", 100.0, "PAID"),
         'resd': (None, "TOTAL", 51007.0, None),
         'lab': (None, None, None, None),
         'dig': (None, None, None, None),
         'af': (None, None, None, None),
         'fuel': (datetime.date(2026, 9, 2), "PETROL / NIZAM", 1000.0, "PAID"),
         'med': (datetime.date(2026, 9, 7), "HAKIM SAHEB", 500.0, "PAID"),
         'rep': (datetime.date(2026, 9, 7), None, "MUNNA", 3000.0, "PAID"),
         'rel': ("MASJID", None)},

        # ZYED (idx 5)
        {'adv': (None, None, None, None),
         'urp': (None, "TOTAL", 150523.0, None),
         'resd': (None, None, None, None),
         'lab': (None, None, None, None),
         'dig': (None, None, None, None),
         'af': (None, None, None, None),
         'fuel': (datetime.date(2026, 9, 11), "PETROL / MEHRAB", 1000.0, None),
         'med': (datetime.date(2026, 9, 17), "BLOOD TEST SULTAN", 3500.0, "PAID"),
         'rep': (datetime.date(2026, 9, 7), None, "MEHRAB", 9600.0, "PAID"),
         'rel': ("OTHERS", None)},

        # ISHRAQUE (idx 6)
        {'adv': (None, None, None, None),
         'urp': (None, None, None, None),
         'resd': (None, None, None, None),
         'lab': (None, None, None, None),
         'dig': (None, None, None, None),
         'af': (None, None, None, None),
         'fuel': (None, None, None, None),
         'med': (datetime.date(2026, 9, 17), "HAKIM SAHEB", 10000.0, "PAID"),
         'rep': (datetime.date(2026, 9, 7), None, "ISHRAQ", 13450.0, "PAID"),
         'rel': (None, None)},

        # MERAB (idx 7)
        {'adv': (None, None, None, None),
         'urp': (None, None, None, None),
         'resd': (None, None, None, None),
         'lab': (None, None, None, None),
         'dig': (None, None, None, None),
         'af': (None, None, None, None),
         'fuel': (None, None, None, None),
         'med': (datetime.date(2026, 9, 24), "HAKIM SAHEB (TAJ)", 5000.0, "PAID"),
         'rep': (None, None, None, None, None),
         'rel': (None, None)}
    ]

    print("Data structures ready. Writing clean database and todays expenditure...")

    def write_month_block(ws, start_row, month_name, people_dict, side_data):
        ws.Cells(start_row, 1).Value = month_name
        for c in range(2, 45):
            ws.Cells(start_row, c).Value = None

        row_idx = start_row + 1
        for p_key in personnel_order:
            if p_key in people_dict:
                p_name, sal, ptype, pdate, pstatus = people_dict[p_key]
            elif p_key == "TOTAL":
                p_name = "TOTAL"
                sal = 298600.0 if "JULY" in month_name else 293600.0
                ptype, pdate, pstatus = None, None, None
            else:
                p_name = p_key
                sal, ptype, pdate, pstatus = None, None, None, None

            ws.Cells(row_idx, 1).Value = p_name
            ws.Cells(row_idx, 2).Value = sal
            ws.Cells(row_idx, 3).Value = ptype
            ws.Cells(row_idx, 4).Value = to_pywintypes_date(pdate)
            ws.Cells(row_idx, 5).Value = pstatus

            ws.Cells(row_idx, 1).NumberFormat = "General"
            ws.Cells(row_idx, 2).NumberFormat = "#,##0"
            ws.Cells(row_idx, 3).NumberFormat = "General"
            if pdate:
                ws.Cells(row_idx, 4).NumberFormat = "dd-mmm-yyyy"
            else:
                ws.Cells(row_idx, 4).NumberFormat = "General"
            ws.Cells(row_idx, 5).NumberFormat = "General"

            side_idx = row_idx - (start_row + 1)
            if side_idx < len(side_data):
                item = side_data[side_idx]
                ws.Cells(row_idx, 6).Value = to_pywintypes_date(item['adv'][0])
                ws.Cells(row_idx, 7).Value = item['adv'][1]
                ws.Cells(row_idx, 8).Value = item['adv'][2]
                ws.Cells(row_idx, 9).Value = item['adv'][3]

                ws.Cells(row_idx, 10).Value = to_pywintypes_date(item['urp'][0])
                ws.Cells(row_idx, 11).Value = item['urp'][1]
                ws.Cells(row_idx, 12).Value = item['urp'][2]
                ws.Cells(row_idx, 13).Value = item['urp'][3]

                ws.Cells(row_idx, 14).Value = to_pywintypes_date(item['resd'][0])
                ws.Cells(row_idx, 15).Value = item['resd'][1]
                ws.Cells(row_idx, 16).Value = item['resd'][2]
                ws.Cells(row_idx, 17).Value = item['resd'][3]

                ws.Cells(row_idx, 18).Value = to_pywintypes_date(item['lab'][0])
                ws.Cells(row_idx, 19).Value = item['lab'][1]
                ws.Cells(row_idx, 20).Value = item['lab'][2]
                ws.Cells(row_idx, 21).Value = item['lab'][3]

                ws.Cells(row_idx, 22).Value = to_pywintypes_date(item['dig'][0])
                ws.Cells(row_idx, 23).Value = item['dig'][1]
                ws.Cells(row_idx, 24).Value = item['dig'][2]
                ws.Cells(row_idx, 25).Value = item['dig'][3]

                ws.Cells(row_idx, 26).Value = to_pywintypes_date(item['af'][0])
                ws.Cells(row_idx, 27).Value = item['af'][1]
                ws.Cells(row_idx, 28).Value = item['af'][2]
                ws.Cells(row_idx, 29).Value = item['af'][3]

                ws.Cells(row_idx, 30).Value = to_pywintypes_date(item['fuel'][0])
                ws.Cells(row_idx, 31).Value = item['fuel'][1]
                ws.Cells(row_idx, 32).Value = item['fuel'][2]
                ws.Cells(row_idx, 33).Value = item['fuel'][3]

                ws.Cells(row_idx, 34).Value = to_pywintypes_date(item['med'][0])
                ws.Cells(row_idx, 35).Value = item['med'][1]
                ws.Cells(row_idx, 36).Value = item['med'][2]
                ws.Cells(row_idx, 37).Value = item['med'][3]

                ws.Cells(row_idx, 38).Value = to_pywintypes_date(item['rep'][0])
                ws.Cells(row_idx, 39).Value = item['rep'][1]
                ws.Cells(row_idx, 40).Value = item['rep'][2]
                ws.Cells(row_idx, 41).Value = item['rep'][3]
                ws.Cells(row_idx, 42).Value = item['rep'][4]

                ws.Cells(row_idx, 43).Value = item['rel'][0]
                ws.Cells(row_idx, 44).Value = item['rel'][1]
            else:
                for c in range(6, 45):
                    ws.Cells(row_idx, c).Value = None

            date_cols = [6, 10, 14, 18, 22, 26, 30, 34, 38]
            amt_cols = [8, 12, 16, 20, 24, 28, 32, 36, 41, 44]
            txt_cols = [7, 9, 11, 13, 15, 17, 19, 21, 23, 25, 27, 29, 31, 33, 35, 37, 39, 40, 42, 43]

            for c in date_cols:
                v = ws.Cells(row_idx, c).Value
                if v: ws.Cells(row_idx, c).NumberFormat = "dd-mmm-yyyy"
                else: ws.Cells(row_idx, c).NumberFormat = "General"

            for c in amt_cols:
                v = ws.Cells(row_idx, c).Value
                if v is not None and v != "": ws.Cells(row_idx, c).NumberFormat = "#,##0"
                else: ws.Cells(row_idx, c).NumberFormat = "General"

            for c in txt_cols:
                ws.Cells(row_idx, c).NumberFormat = "General"

            row_idx += 1
        return row_idx - 1

    print("Rebuilding database_expenditure...")
    wsDB.Range("A2:AZ100").Clear()

    for c, h in enumerate(headers, 1):
        wsDB.Cells(1, c).Value = h
        wsDB.Cells(1, c).NumberFormat = "General"

    end_july = write_month_block(wsDB, 2, "JULY_2026", july_people, july_side)
    print(f"DB July written: Rows 2 to {end_july}")

    end_august = write_month_block(wsDB, 34, "AUGUST_2026", august_people, august_side)
    print(f"DB August written: Rows 34 to {end_august}")

    print("Rebuilding TODAYS_EXPENDITURE (Active Month: JULY_2026)...")
    for c, h in enumerate(headers, 1):
        wsToday.Cells(1, c).Value = h
        wsToday.Cells(1, c).NumberFormat = "General"

    wsToday.Range("A2:AZ34").Clear()
    write_month_block(wsToday, 2, "JULY_2026", july_people, july_side)

    # Clean executive formatting
    rng_header = wsToday.Range("A1:AR1")
    rng_header.Interior.Color = 0x33251E
    rng_header.Font.Color = 0xFFFFFF
    rng_header.Font.Bold = True
    rng_header.Font.Size = 10
    rng_header.HorizontalAlignment = -4108

    rng_month = wsToday.Range("A2:AR2")
    rng_month.Interior.Color = 0x804000
    rng_month.Font.Color = 0xFFFFFF
    rng_month.Font.Bold = True
    rng_month.Font.Size = 11

    rng_total = wsToday.Range("A33:AR33")
    rng_total.Interior.Color = 0x1B1B99
    rng_total.Font.Color = 0xFFFFFF
    rng_total.Font.Bold = True
    rng_total.Font.Size = 11

    # Vertical section dividers
    separator_cols = [5, 9, 13, 17, 21, 25, 29, 33, 37, 42, 44]
    for c in separator_cols:
        col_rng = wsToday.Range(wsToday.Cells(1, c), wsToday.Cells(33, c))
        border = col_rng.Borders(10)
        border.LineStyle = 1
        border.Weight = 3
        border.Color = 0x473C33

    wb.Save()
    print("Workbook successfully saved!")
    wb.Close(False)
    xl.Quit()

if __name__ == '__main__':
    main()
