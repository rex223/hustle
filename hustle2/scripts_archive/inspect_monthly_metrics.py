import win32com.client
import os

xl = win32com.client.Dispatch('Excel.Application')
xl.DisplayAlerts = False
xl.Visible = False

wb = xl.Workbooks.Open(os.path.abspath('RENTAL DATASHEET 19 OCT.xlsm'), ReadOnly=True)
wsDB = wb.Sheets('database_expenditure')
wsPurch = wb.Sheets('DATABASE_PURCHASE')
wsRentMisc = wb.Sheets('RENT PERSONS MISCELLANEOUS')
wsMisc = wb.Sheets('MISCELLANEOUS')

def get_month_numbers(wsDB, start_r, end_r):
    # Salaries
    sal_paid = 0.0
    sal_due = 0.0
    for r in range(start_r + 1, end_r + 1):
        name = str(wsDB.Cells(r, 1).Value or "").strip().upper()
        if name == "TOTAL" or "TOTAL" in name: continue
        amt = wsDB.Cells(r, 2).Value
        status = str(wsDB.Cells(r, 5).Value or "").strip().upper()
        if amt:
            if status == "PAID": sal_paid += float(amt)
            else: sal_due += float(amt)

    # Side expenses
    exp_paid = 0.0
    exp_due = 0.0
    cat_splits = {'adv': 0.0, 'cesc': 0.0, 'labour': 0.0, 'digrum': 0.0, 'feed_fuel': 0.0, 'med_rep': 0.0}

    # Adv (8), URP (12), RESD (16), Lab (20), Dig (24), AF (28), Fuel (32), Med (36), Rep (41)
    amt_cols = [
        (8, 9, 'adv'), (12, 13, 'cesc'), (16, 17, 'cesc'), 
        (20, 21, 'labour'), (24, 25, 'digrum'), (28, 29, 'feed_fuel'), 
        (32, 33, 'feed_fuel'), (36, 37, 'med_rep'), (41, 42, 'med_rep')
    ]

    for r in range(start_r + 1, end_r + 1):
        name = str(wsDB.Cells(r, 1).Value or "").strip().upper()
        if name == "TOTAL" or "TOTAL" in name: continue

        for amt_c, paid_c, cat in amt_cols:
            v = wsDB.Cells(r, amt_c).Value
            desc = str(wsDB.Cells(r, amt_c - 1).Value or "").strip().upper() if amt_c > 1 else ""
            if "TOTAL" in desc: continue

            if v is not None and v != "":
                try:
                    val = float(v)
                    if val <= 0: continue
                    is_p = False
                    for pc in range(paid_c, min(paid_c + 3, 44)):
                        if str(wsDB.Cells(r, pc).Value or "").strip().upper() == "PAID":
                            is_p = True
                            break
                    if is_p:
                        exp_paid += val
                        cat_splits[cat] += val
                    else:
                        exp_due += val
                except ValueError:
                    pass

    return sal_paid, sal_due, exp_paid, exp_due, cat_splits

# July is Rows 2 to 33
j_sal_p, j_sal_d, j_exp_p, j_exp_d, j_cats = get_month_numbers(wsDB, 2, 33)
# August is Rows 34 to 65
a_sal_p, a_sal_d, a_exp_p, a_exp_d, a_cats = get_month_numbers(wsDB, 34, 65)

print("=== JULY 2026 EXPENDITURE ===")
print(f"Salaries Paid: {j_sal_p:,.2f} | Due: {j_sal_d:,.2f}")
print(f"Side Exp Paid: {j_exp_p:,.2f} | Due: {j_exp_d:,.2f}")
print(f"Total Operating Paid: {j_sal_p + j_exp_p:,.2f} | Total Due: {j_sal_d + j_exp_d:,.2f}")
print("Category splits:", j_cats)

print("\n=== AUGUST 2026 EXPENDITURE ===")
print(f"Salaries Paid: {a_sal_p:,.2f} | Due: {a_sal_d:,.2f}")
print(f"Side Exp Paid: {a_exp_p:,.2f} | Due: {a_exp_d:,.2f}")
print(f"Total Operating Paid: {a_sal_p + a_exp_p:,.2f} | Total Due: {a_sal_d + a_exp_d:,.2f}")
print("Category splits:", a_cats)

wb.Close(False)
xl.Quit()
