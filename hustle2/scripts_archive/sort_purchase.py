import win32com.client
import os
import re
import traceback

xl = win32com.client.Dispatch('Excel.Application')
xl.DisplayAlerts = False
wb = xl.Workbooks.Open(os.path.abspath('RENTAL DATASHEET 3 SEP.xlsm'))

try:
    sh_entry = wb.Sheets('PURCHASE SHEET')
    tbl_entry = sh_entry.ListObjects('EntryFormTable')
    
    data = []
    current_group = ""
    for i in range(1, tbl_entry.DataBodyRange.Rows.Count + 1):
        grp = tbl_entry.DataBodyRange.Cells(i, 1).Value
        if grp and str(grp).strip() != "":
            current_group = str(grp).strip()
            
        item = tbl_entry.DataBodyRange.Cells(i, 2).Value
        if item and str(item).strip() != "":
            qty_kg = tbl_entry.DataBodyRange.Cells(i, 3).Value
            qty_pcs = tbl_entry.DataBodyRange.Cells(i, 4).Value
            rate = tbl_entry.DataBodyRange.Cells(i, 5).Value
            trsp = tbl_entry.DataBodyRange.Cells(i, 6).Value
            total = tbl_entry.DataBodyRange.Cells(i, 7).Value
            
            data.append({
                'group': current_group,
                'item': str(item).strip(),
                'qty_kg': qty_kg,
                'qty_pcs': qty_pcs,
                'rate': rate,
                'trsp': trsp,
                'total': total
            })
            
    def sort_key(d):
        g = d['group'].upper()
        # Find number in group name
        m = re.search(r'\d+', g)
        grp_num = int(m.group(0)) if m else 999
        return (grp_num, g, d['item'].upper())
        
    data.sort(key=sort_key)
    
    # Write back
    for i in range(len(data)):
        row_idx = i + 1
        
        # Determine if we should show the group name
        if i == 0 or data[i]['group'] != data[i-1]['group']:
            tbl_entry.DataBodyRange.Cells(row_idx, 1).Value = data[i]['group']
        else:
            tbl_entry.DataBodyRange.Cells(row_idx, 1).Value = None
            
        tbl_entry.DataBodyRange.Cells(row_idx, 2).Value = data[i]['item']
        tbl_entry.DataBodyRange.Cells(row_idx, 3).Value = data[i]['qty_kg']
        tbl_entry.DataBodyRange.Cells(row_idx, 4).Value = data[i]['qty_pcs']
        tbl_entry.DataBodyRange.Cells(row_idx, 5).Value = data[i]['rate']
        tbl_entry.DataBodyRange.Cells(row_idx, 6).Value = data[i]['trsp']
        tbl_entry.DataBodyRange.Cells(row_idx, 7).Value = data[i]['total']
        
    # Clear any leftover rows if we skipped blank items
    for i in range(len(data) + 1, tbl_entry.DataBodyRange.Rows.Count + 1):
        tbl_entry.DataBodyRange.Cells(i, 1).Value = None
        tbl_entry.DataBodyRange.Cells(i, 2).Value = None
        tbl_entry.DataBodyRange.Cells(i, 3).Value = None
        tbl_entry.DataBodyRange.Cells(i, 4).Value = None
        tbl_entry.DataBodyRange.Cells(i, 5).Value = None
        tbl_entry.DataBodyRange.Cells(i, 6).Value = None
        tbl_entry.DataBodyRange.Cells(i, 7).Value = None

    # Remove extra blank rows from the table itself so it shrinks
    while tbl_entry.DataBodyRange.Rows.Count > len(data):
        tbl_entry.ListRows(tbl_entry.ListRows.Count).Delete()

    wb.Save()
    print('SUCCESS: Sorted and formatted PURCHASE SHEET')
except Exception as e:
    traceback.print_exc()
finally:
    wb.Close(False)
    xl.Quit()
