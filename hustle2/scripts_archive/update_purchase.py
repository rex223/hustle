import win32com.client
import os
import traceback

xl = win32com.client.Dispatch('Excel.Application')
xl.DisplayAlerts = False
wb = xl.Workbooks.Open(os.path.abspath('RENTAL DATASHEET 3 SEP.xlsm'))

try:
    # 1. ADD MISSING ITEMS TO PURCHASE SHEET
    sh_entry = wb.Sheets('PURCHASE SHEET')
    tbl_entry = sh_entry.ListObjects('EntryFormTable')
    
    items_to_add = [
        ('GROUP 2 ', 'TOMATO'), ('GROUP 2 ', 'BADHA GOBI'), ('GROUP 2 ', 'BANGON'), 
        ('GROUP 2 ', 'BARBATI'), ('GROUP 2 ', 'CHIMMI'), ('GROUP 2 ', 'DATTA'), 
        ('GROUP 2 ', 'FOOL GOBI'), ('GROUP 2 ', 'POTOL'), ('GROUP 2 ', 'SIM'), 
        ('GROUP 2 ', 'VINDI'), ('GROUP 2 ', 'ZINGHI'),
        ('GROUP 4', 'SUGAR'),
        ('GROUP 6', 'DALCHINI'), ('GROUP 6', 'ELIACHI'), ('GROUP 6', 'LONG'),
        ('GROUP 7', 'GRAPES'), ('GROUP 7', 'MANGO'), ('GROUP 7', 'MUSUMBI'), 
        ('GROUP 7', 'ORANGE'), ('GROUP 7', 'PINEAPPLE'), ('GROUP 7', 'WATER MELON'),
        ('GROUP 10', 'BADAM'), ('GROUP 10', 'KAJU'), ('GROUP 10', 'KISMIS'), 
        ('GROUP 10', 'BETTEL'), ('GROUP 10', 'COAL'), ('GROUP 10', 'DUB'), 
        ('GROUP 10', 'HONEY'), ('GROUP 10', 'PAAN'), ('GROUP 10', 'TEA')
    ]
    
    # Check if already added to avoid duplicates
    existing_items = []
    if tbl_entry.DataBodyRange is not None:
        for i in range(1, tbl_entry.DataBodyRange.Rows.Count + 1):
            v = tbl_entry.DataBodyRange.Cells(i, 2).Value
            if v: existing_items.append(str(v).strip().upper())
        
    for grp, item in items_to_add:
        if item.upper() not in existing_items:
            r = tbl_entry.ListRows.Add()
            r.Range(1, 1).Value = grp
            r.Range(1, 2).Value = item
            
    # 2. CREATE PIVOT TABLE ON THE SIDE OF DATABASE_PURCHASE
    sh_db = wb.Sheets('DATABASE_PURCHASE')
    
    # Check if Pivot Table already exists and delete it to recreate
    for pt in sh_db.PivotTables():
        pt.TableRange2.Clear()
        
    pc = wb.PivotCaches().Create(SourceType=1, SourceData="DailyPurchaseLedger", Version=6)
    # Put it in Col J
    pt = pc.CreatePivotTable(TableDestination=sh_db.Range("J2"), TableName="MonthlyTotals")
    
    date_field = pt.PivotFields("DATE")
    date_field.Orientation = 1 # xlRowField
    
    val_field = pt.PivotFields("TOTAL AMOUNT")
    val_field.Orientation = 4 # xlDataField
    val_field.Function = -4157 # xlSum
    val_field.NumberFormat = "#,##0"
    
    pt.TableStyle2 = "PivotStyleLight16"
    sh_db.Columns("J:K").AutoFit()
    
    # Try to group by Months and Years (Works in Excel if dates are valid)
    try:
        sh_db.Range("J3").Group(Periods=(False, False, False, False, True, False, True)) # Months and Years
    except:
        pass # Ignore if it fails (e.g., blank dates or already grouped)

    wb.Save()
    print('SUCCESS: Updated PURCHASE SHEET and added Monthly Totals Pivot Table to DATABASE_PURCHASE.')
except Exception as e:
    traceback.print_exc()
finally:
    wb.Close(False)
    xl.Quit()
