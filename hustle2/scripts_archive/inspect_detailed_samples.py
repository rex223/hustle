import win32com.client
import os

xl = win32com.client.Dispatch('Excel.Application')
xl.DisplayAlerts = False
xl.Visible = False

wb = xl.Workbooks.Open(os.path.abspath('RENTAL DATASHEET 19 OCT.xlsm'), ReadOnly=True)
wsDB = wb.Sheets('database_expenditure')
wsToday = wb.Sheets('TODAYS_EXPENDITURE')

def dump_block(ws, start_row, end_row, title):
    print(f"\n=================== {title} (Rows {start_row} to {end_row}) ===================")
    for r in range(start_row, end_row + 1):
        vals = {}
        for c in range(1, 46):
            v = ws.Cells(r, c).Value
            fmt = ws.Cells(r, c).NumberFormat
            if v is not None:
                vals[c] = (v, fmt)
        if vals:
            name = ws.Cells(r, 1).Value
            print(f"R{r:2d} [{name}]:")
            for c in sorted(vals.keys()):
                val, fmt = vals[c]
                print(f"   Col {c:2d}: {str(val):25s} (Format: {fmt})")

dump_block(wsDB, 2, 8, "DB JULY SAMPLE")
dump_block(wsDB, 34, 40, "DB AUGUST SAMPLE")
dump_block(wsToday, 2, 8, "TODAYS EXPENDITURE SAMPLE")

wb.Close(False)
xl.Quit()
