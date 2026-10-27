"""
update_month_labels.py
Updates existing month labels to be full uppercase names (e.g. JULY-2026).
Re-injects updated VBA.
Run with Excel CLOSED.
"""
import win32com.client
import os

WORKBOOK_PATH = os.path.abspath('hustle2/RENTAL DATASHEET 1ST AUG.xlsm')

def run():
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    try:
        wb = excel.Workbooks.Open(WORKBOOK_PATH)
        ws_entry = wb.Sheets("TODAYS EXPENDITURE")
        ws_db    = wb.Sheets("DATABASE_EXPENDITURE")

        # 1. Update DATABASE_EXPENDITURE labels
        last_row = ws_db.Cells(ws_db.Rows.Count, 1).End(-4162).Row # xlUp
        for r in range(2, last_row + 1):
            val = str(ws_db.Cells(r, 1).Value).strip()
            if val == "Jul-2026":
                ws_db.Cells(r, 1).NumberFormat = "@"
                ws_db.Cells(r, 1).Value = "JULY-2026"
            elif val == "--- Jul-2026 ---":
                ws_db.Cells(r, 1).NumberFormat = "@"
                ws_db.Cells(r, 1).Value = "--- JULY-2026 ---"
        
        print("  DATABASE_EXPENDITURE labels updated to full caps.")

        # 2. Update TODAYS EXPENDITURE C1
        ws_entry.Range("C1").NumberFormat = "@"
        ws_entry.Range("C1").Value = "AUGUST-2026"
        
        # Remove old strict validation just in case
        try:
            ws_entry.Range("C1").Validation.Delete()
        except:
            pass
            
        print("  TODAYS EXPENDITURE C1 updated to full caps.")

        # 3. Re-inject VBA
        bas_dir = os.path.normpath(
            os.path.join(os.path.dirname(WORKBOOK_PATH), 'scripts_archive')
        )
        sheet_bas = os.path.join(bas_dir, 'vba_sheet_expenditure.bas')
        with open(sheet_bas, 'r', encoding='utf-8') as f:
            vba_sheet = f.read()

        mod_entry = wb.VBProject.VBComponents(ws_entry.CodeName).CodeModule
        mod_entry.DeleteLines(1, mod_entry.CountOfLines)
        mod_entry.AddFromString(vba_sheet)
        print("  VBA re-injected with new rules.")

        # 4. Trigger the new VBA to update the header and validation
        ws_entry.Activate()
        
        wb.Save()
        print("SUCCESS")

    except Exception as e:
        import traceback
        traceback.print_exc()
        print(f"Error: {e}")
    finally:
        try:
            wb.Close(False)
        except Exception:
            pass
        excel.Quit()

if __name__ == '__main__':
    run()
