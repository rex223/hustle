"""
fix_db_and_labels.py
1. Fix DATABASE_EXPENDITURE: set col A as text format, re-write month labels, restore col headers
2. Fix TODAYS EXPENDITURE C1 to store text properly
3. Fix Worksheet_Activate to handle text labels
Run with Excel CLOSED.
"""
import win32com.client
import os

WORKBOOK_PATH = os.path.abspath('hustle2/RENTAL DATASHEET 1ST AUG.xlsm')
TOTAL_COLS    = 27
JULY_LABEL    = "Jul-2026"

GROUP_COLORS = [
    (range(1,  6),  (0xEA, 0xD8, 0xB1)),
    (range(6,  8),  (0xD4, 0xED, 0xDA)),
    (range(8,  10), (0xD1, 0xEC, 0xF1)),
    (range(10, 12), (0xFF, 0xF3, 0xCD)),
    (range(12, 14), (0xF8, 0xD7, 0xDA)),
    (range(14, 17), (0xE2, 0xD9, 0xF3)),
    (range(17, 20), (0xFF, 0xF9, 0xC4)),
    (range(20, 23), (0xFC, 0xE4, 0xEC)),
    (range(23, 26), (0xE8, 0xF5, 0xE9)),
    (range(26, 28), (0xE3, 0xF2, 0xFD)),
]

def excol(r, g, b):
    return b | (g << 8) | (r << 16)

def col_color(col_1indexed):
    for rng, (r, g, b) in GROUP_COLORS:
        if col_1indexed in rng:
            return excol(r, g, b)
    return -4142


def run():
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    try:
        wb    = excel.Workbooks.Open(WORKBOOK_PATH)
        ws_src   = wb.Sheets("EXPENDITURE")
        ws_entry = wb.Sheets("TODAYS EXPENDITURE")
        ws_db    = wb.Sheets("DATABASE_EXPENDITURE")

        # ============================================================
        # 1. FIX DATABASE_EXPENDITURE
        # ============================================================

        # 1a. Restore column headers row 1
        ws_db.Cells.Clear()

        # MONTH header in A1
        ws_db.Range("A1").NumberFormat = "@"
        ws_db.Range("A1").Value = "MONTH"
        ws_db.Range("A1").Font.Bold = True
        ws_db.Range("A1").Font.Size = 11
        ws_db.Range("A1").Interior.Color = excol(0x34, 0x3A, 0x40)
        ws_db.Range("A1").Font.Color = excol(0xFF, 0xFF, 0xFF)
        ws_db.Range("A1").HorizontalAlignment = -4108
        ws_db.Range("A1").VerticalAlignment   = -4108
        ws_db.Range("A1").Borders.LineStyle = 1

        # Copy EXPENDITURE headers (row 1) into DB row 1 cols B onward
        ws_src.Range(ws_src.Cells(1, 1), ws_src.Cells(1, TOTAL_COLS)).Copy(ws_db.Cells(1, 2))

        # Apply group colors to header cols
        for col in range(1, TOTAL_COLS + 1):
            ws_db.Cells(1, col + 1).Interior.Color = col_color(col)
            ws_db.Cells(1, col + 1).Borders.LineStyle = 1
            ws_db.Cells(1, col + 1).Font.Bold = True

        ws_db.Rows("1:1").RowHeight = 22
        ws_db.Columns(1).ColumnWidth = 12

        # 1b. Re-add Jul-2026 data block
        # Section header
        next_row = 2
        header_rng = ws_db.Range(ws_db.Cells(next_row, 1), ws_db.Cells(next_row, TOTAL_COLS + 1))
        header_rng.Merge()
        header_rng.Cells(1, 1).NumberFormat = "@"
        header_rng.Cells(1, 1).Value = f"--- {JULY_LABEL} ---"
        header_rng.Font.Bold = True
        header_rng.Font.Size = 11
        header_rng.Interior.Color = excol(0x49, 0x50, 0x57)
        header_rng.Font.Color = excol(0xFF, 0xFF, 0xFF)
        header_rng.HorizontalAlignment = -4108
        header_rng.VerticalAlignment   = -4108
        ws_db.Rows(next_row).RowHeight = 20
        next_row += 1

        # Copy all 22 data rows from EXPENDITURE
        # Turn off AutoFilter first
        if ws_src.AutoFilterMode:
            ws_src.AutoFilterMode = False

        for src_row in range(2, 24):  # rows 2-23 = 22 people
            # IMPORTANT: set text format on col A FIRST before value
            ws_db.Cells(next_row, 1).NumberFormat = "@"
            ws_db.Cells(next_row, 1).Value = JULY_LABEL
            ws_db.Cells(next_row, 1).Interior.Color = excol(0xDC, 0xDC, 0xDC)
            ws_db.Cells(next_row, 1).Borders.LineStyle = 1
            ws_db.Cells(next_row, 1).HorizontalAlignment = -4108

            for col in range(1, TOTAL_COLS + 1):
                val = ws_src.Cells(src_row, col).Value
                ws_db.Cells(next_row, col + 1).Value = val
                ws_db.Cells(next_row, col + 1).Borders.LineStyle = 1
                ws_db.Cells(next_row, col + 1).Interior.Color = col_color(col)

            next_row += 1

        # Auto-fit and freeze header
        ws_db.Columns("A:AB").AutoFit()
        ws_db.Activate()
        excel.ActiveWindow.FreezePanes = False
        excel.ActiveWindow.ScrollRow = 1
        excel.ActiveWindow.ScrollColumn = 1
        ws_db.Cells(2, 1).Select()
        excel.ActiveWindow.FreezePanes = True

        print(f"  DATABASE_EXPENDITURE rebuilt: headers + {JULY_LABEL} data ({next_row - 3} rows)")

        # ============================================================
        # 2. FIX TODAYS EXPENDITURE C1 to store text "Aug-2026"
        # ============================================================
        ws_entry.Range("C1").NumberFormat = "@"
        ws_entry.Range("C1").Value = "Aug-2026"
        print(f"  TODAYS EXPENDITURE C1 fixed to: {ws_entry.Range('C1').Value}")

        # Also fix the validation to allow manual typing (remove strict validation)
        ws_entry.Range("C1").Validation.Delete()

        # ============================================================
        # 3. FIX VBA: Worksheet_Activate — use text format for C1 cell
        # ============================================================
        bas_dir = os.path.normpath(
            os.path.join(os.path.dirname(WORKBOOK_PATH), 'scripts_archive')
        )
        sheet_bas = os.path.join(bas_dir, 'vba_sheet_expenditure.bas')
        with open(sheet_bas, 'r', encoding='utf-8') as f:
            vba_sheet = f.read()

        # Re-inject sheet VBA (same as before, it reads col A which is now proper text)
        mod_entry = wb.VBProject.VBComponents(ws_entry.CodeName).CodeModule
        mod_entry.DeleteLines(1, mod_entry.CountOfLines)
        mod_entry.AddFromString(vba_sheet)
        print("  VBA re-injected")

        # ============================================================
        # 4. SAVE
        # ============================================================
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
