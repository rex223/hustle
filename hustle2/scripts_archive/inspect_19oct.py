import win32com.client
import os

xl = win32com.client.Dispatch('Excel.Application')
xl.DisplayAlerts = False
xl.Visible = False

file_path = os.path.abspath('RENTAL DATASHEET 19 OCT.xlsm')
print("Opening:", file_path)
wb = xl.Workbooks.Open(file_path, ReadOnly=True)

try:
    print("\n--- SHEETS ---")
    sheet_names = [s.Name for s in wb.Sheets]
    print(sheet_names)
    
    # Export VBA
    os.makedirs('vba_19oct', exist_ok=True)
    for comp in wb.VBProject.VBComponents:
        try:
            lines = comp.CodeModule.CountOfLines
            if lines > 0:
                code = comp.CodeModule.Lines(1, lines)
                with open(f'vba_19oct/{comp.Name}.txt', 'w', encoding='utf-8') as f:
                    f.write(code)
                print(f"Exported {comp.Name} ({lines} lines)")
        except Exception as e:
            print(f"Error {comp.Name}: {e}")

    # Inspect TODAYS_EXPENDITURE
    print("\n--- TODAYS_EXPENDITURE ---")
    ws_exp = wb.Sheets('TODAYS_EXPENDITURE')
    print("UsedRange:", ws_exp.UsedRange.Rows.Count, "rows,", ws_exp.UsedRange.Columns.Count, "cols")
    for r in range(1, min(35, ws_exp.UsedRange.Rows.Count + 1)):
        row_vals = [ws_exp.Cells(r, c).Value for c in range(1, 20)]
        if any(v is not None for v in row_vals):
            print(f"Row {r:2d}: {row_vals}")

    # Inspect RENT PERSONS MISCELLANEOUS
    print("\n--- RENT PERSONS MISCELLANEOUS ---")
    if 'RENT PERSONS MISCELLANEOUS' in sheet_names:
        ws_rpm = wb.Sheets('RENT PERSONS MISCELLANEOUS')
        for r in range(1, min(20, ws_rpm.UsedRange.Rows.Count + 1)):
            row_vals = [ws_rpm.Cells(r, c).Value for c in range(1, 15)]
            if any(v is not None for v in row_vals):
                print(f"Row {r:2d}: {row_vals}")

    # Check Kamal in todays_rent_data
    print("\n--- KAMAL IN TODAYS_RENT_DATA ---")
    ws_rent = wb.Sheets('todays_rent_data')
    for r in range(1, ws_rent.UsedRange.Rows.Count + 1):
        row_vals = [ws_rent.Cells(r, c).Value for c in range(1, 10)]
        if any('KAMAL' in str(v).upper() for v in row_vals if v):
            print(f"Row {r:2d}: {row_vals}")

finally:
    wb.Close(False)
    xl.Quit()
    print("Done inspect.")
