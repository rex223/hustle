import shutil
import os
import win32com.client

# Test on a copy
src = os.path.abspath('RENTAL DATASHEET 19 OCT.xlsm')
dst = os.path.abspath('test_exp_reorder.xlsm')
if os.path.exists(dst):
    try:
        os.remove(dst)
    except:
        pass
shutil.copyfile(src, dst)

xl = win32com.client.Dispatch('Excel.Application')
xl.DisplayAlerts = False
xl.Visible = False

wb = xl.Workbooks.Open(dst)
ws = wb.Sheets('TODAYS_EXPENDITURE')

# Current rows in TODAYS_EXPENDITURE:
# 3: NAIM, 4: NIZAM, 5: SHAFIQUE, 6: KAMAL, 7: NAZIR, 8: ZYED, 9: ISHRAQUE, 10: MERAB, 11: ENAYATULLA,
# 12: RASHEED, 13: MUNNA, 14: RAFIQUE, 15: SURAJ , 16: SECURITY, 17: SAM, 18: ZOE/ZARA,
# 19: EC, 20: ESI, 21: NARAYAN, 22: BAWA, 23-27: SHAHNAWAZ 1-5, 28: ADINUL, 29: PARVEZ, 30: AKHTAR, 31: CESC

# Read all rows into a dict by Name + Type
rows_data = {}
for r in range(3, 32):
    key = (str(ws.Cells(r, 1).Value).strip(), str(ws.Cells(r, 3).Value).strip())
    vals = [ws.Cells(r, c).Value for c in range(1, 6)]
    rows_data[key] = vals

print("Total items read:", len(rows_data))

# Define the new ordered keys according to client instructions:
# 1. BAWA (Top)
# 2. Staff (NAIM, NIZAM, SHAFIQUE, KAMAL, NAZIR, ZYED, ISHRAQUE, MERAB, ENAYATULLA, RASHEED, MUNNA, RAFIQUE)
# 3. Site/Security (SURAJ , SECURITY)
# 4. Residential (SAM, ZOE/ZARA)
# 5. Statutory (ESI, NARAYAN)
# 6. Weekly Maintenance (SHAHNAWAZ W1 to W5)
# 7. Utilities (CESC+RAJA+OTHERS)
# 8. ADINUL, AKHTAR (above Parvez)
# 9. PARVEZ (bottom)
# 10. EC (in the last of the list bottom)

ordered_keys = [
    ('BAWA', 'None'),
    ('NAIM', 'None'),
    ('NIZAM', 'GATE'),
    ('SHAFIQUE', 'None'),
    ('KAMAL', 'None'),
    ('NAZIR', 'BARI'),
    ('ZYED', 'None'),
    ('ISHRAQUE', 'None'),
    ('MERAB', 'MASJID'),
    ('ENAYATULLA', 'None'),
    ('RASHEED', 'URLP'),
    ('MUNNA', 'None'),
    ('RAFIQUE', '56.0'),
    ('SURAJ', 'PPL'),
    ('SECURITY', 'PPL'),
    ('SAM', 'RESD'),
    ('ZOE/ZARA', 'RESD'),
    ('ESI', 'None'),
    ('NARAYAN', 'None'),
    ('SHAHNAWAZ/108033', 'WEEK 1'),
    ('SHAHNAWAZ/108033', 'WEEK 2'),
    ('SHAHNAWAZ/108033', 'WEEK 3'),
    ('SHAHNAWAZ/108033', 'WEEK 4'),
    ('SHAHNAWAZ/108033', 'WEEK 5'),
    ('CESC+RAJA+OTHERS', 'PPL'),
    ('ADINUL', 'None'),
    ('AKHTAR', 'URLP'),
    ('PARVEZ', 'URLP'),
    ('EC', '50000.0')
]

print("Ordered keys count:", len(ordered_keys))

# Verify all keys exist in rows_data
for k in ordered_keys:
    found = False
    for rk in rows_data:
        if k[0] in rk[0] and (k[1] == 'None' or k[1] in rk[1]):
            found = True
            break
    if not found:
        print("MISSING KEY:", k)

wb.Close(False)
xl.Quit()

if os.path.exists('test_exp_reorder.xlsm'):
    os.remove('test_exp_reorder.xlsm')
print("Verification complete.")
