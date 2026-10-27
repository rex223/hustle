"""
rebuild_expenditure_system.py
Full clean rebuild of TODAYS EXPENDITURE and DATABASE_EXPENDITURE.
Run with Excel CLOSED.
"""
import win32com.client
import os

WORKBOOK_PATH = os.path.abspath('hustle2/RENTAL DATASHEET 1ST AUG.xlsm')

DATA_LAST_ROW  = 23   # last row in EXPENDITURE (header=row1, data rows 2-23)
TOTAL_COLS     = 27   # A..AA

# Input columns (1-indexed): D=4, E=5, G=7, I=9, K=11, M=13, O=15, P=16, R=18, S=19, U=21, V=22, Y=25
INPUT_COLS = [4, 5, 7, 9, 11, 13, 15, 16, 18, 19, 21, 22, 25]

GROUP_COLORS = [
    (range(1,  6),  (0xEA, 0xD8, 0xB1)),   # Salary
    (range(6,  8),  (0xD4, 0xED, 0xDA)),   # CESC URP
    (range(8,  10), (0xD1, 0xEC, 0xF1)),   # CESC RESD
    (range(10, 12), (0xFF, 0xF3, 0xCD)),   # Labour
    (range(12, 14), (0xF8, 0xD7, 0xDA)),   # Digrum
    (range(14, 17), (0xE2, 0xD9, 0xF3)),   # Animal Feed
    (range(17, 20), (0xFF, 0xF9, 0xC4)),   # Fuel/Gas
    (range(20, 23), (0xFC, 0xE4, 0xEC)),   # Medical
    (range(23, 26), (0xE8, 0xF5, 0xE9)),   # Repair
    (range(26, 28), (0xE3, 0xF2, 0xFD)),   # Overall/Religion
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
        wb = excel.Workbooks.Open(WORKBOOK_PATH)
        ws_src = wb.Sheets("EXPENDITURE")

        # entry_data_last = last data row in TODAYS EXPENDITURE
        entry_data_last = DATA_LAST_ROW + 1   # row 2 = header copy, rows 3..24 = data

        # ============================================================
        # 1. REBUILD TODAYS EXPENDITURE
        # ============================================================
        try:
            ws_entry = wb.Sheets("TODAYS EXPENDITURE")
        except Exception:
            wb.Sheets.Add(After=wb.Sheets(wb.Sheets.Count))
            wb.ActiveSheet.Name = "TODAYS EXPENDITURE"
            ws_entry = wb.Sheets("TODAYS EXPENDITURE")

        ws_entry.Cells.Clear()
        for i in range(ws_entry.Shapes.Count, 0, -1):
            ws_entry.Shapes.Item(i).Delete()
        while ws_entry.ListObjects.Count > 0:
            ws_entry.ListObjects.Item(1).Delete()

        # Copy exact layout from EXPENDITURE into row 2 of entry sheet
        ws_src.Range(
            ws_src.Cells(1, 1), ws_src.Cells(DATA_LAST_ROW, TOTAL_COLS)
        ).Copy(ws_entry.Cells(2, 1))

        # Apply group colors to all copied rows
        for row in range(2, entry_data_last + 1):
            for col in range(1, TOTAL_COLS + 1):
                ws_entry.Cells(row, col).Interior.Color = col_color(col)
                ws_entry.Cells(row, col).Borders.LineStyle = 1

        # Clear input columns in data rows (3 to entry_data_last)
        for col in INPUT_COLS:
            ws_entry.Range(
                ws_entry.Cells(3, col),
                ws_entry.Cells(entry_data_last, col)
            ).ClearContents()

        # Row 1: Controls bar
        ws_entry.Rows("1:1").RowHeight = 28
        ws_entry.Range("A1").Value = "MONTH:"
        ws_entry.Range("A1").Font.Bold = True
        ws_entry.Range("A1").Font.Size = 11
        ws_entry.Range("A1").HorizontalAlignment = -4152   # xlRight
        ws_entry.Range("A1").VerticalAlignment   = -4108   # xlCenter

        # Yellow month selector
        ws_entry.Range("C1").Value = "Aug-2026"
        ws_entry.Range("C1").Font.Bold = True
        ws_entry.Range("C1").Font.Size = 11
        ws_entry.Range("C1").Interior.Color = excol(0xFF, 0xFF, 0x00)
        ws_entry.Range("C1").Borders.LineStyle = 1
        ws_entry.Range("C1").HorizontalAlignment = -4108
        ws_entry.Range("C1").VerticalAlignment   = -4108

        # Default month column header (col E row 2)
        ws_entry.Cells(2, 5).Value = "AUGUST"

        # SAVE button
        btn_save = ws_entry.Shapes.AddShape(1, 310, 3, 100, 22)
        btn_save.Fill.ForeColor.RGB = excol(0x1E, 0x7E, 0x34)
        btn_save.Line.Visible = False
        btn_save.TextFrame.Characters().Text = "SAVE MONTH"
        btn_save.TextFrame.Characters().Font.Bold = True
        btn_save.TextFrame.Characters().Font.Size = 9
        btn_save.TextFrame.Characters().Font.Color = excol(0xFF, 0xFF, 0xFF)
        btn_save.TextFrame.HorizontalAlignment = -4108
        btn_save.TextFrame.VerticalAlignment   = -4108
        btn_save.OnAction = "SaveMonthlyExpenditures"

        # LOAD button
        btn_load = ws_entry.Shapes.AddShape(1, 420, 3, 100, 22)
        btn_load.Fill.ForeColor.RGB = excol(0x00, 0x69, 0xD9)
        btn_load.Line.Visible = False
        btn_load.TextFrame.Characters().Text = "LOAD MONTH"
        btn_load.TextFrame.Characters().Font.Bold = True
        btn_load.TextFrame.Characters().Font.Size = 9
        btn_load.TextFrame.Characters().Font.Color = excol(0xFF, 0xFF, 0xFF)
        btn_load.TextFrame.HorizontalAlignment = -4108
        btn_load.TextFrame.VerticalAlignment   = -4108
        btn_load.OnAction = "LoadMonthlyExpenditures"

        # Freeze rows 1+2 always visible
        ws_entry.Activate()
        excel.ActiveWindow.FreezePanes = False
        ws_entry.Cells(3, 1).Select()
        excel.ActiveWindow.FreezePanes = True

        # ============================================================
        # 2. REBUILD DATABASE_EXPENDITURE
        # ============================================================
        try:
            ws_db = wb.Sheets("DATABASE_EXPENDITURE")
        except Exception:
            wb.Sheets.Add(After=wb.Sheets(wb.Sheets.Count))
            wb.ActiveSheet.Name = "DATABASE_EXPENDITURE"
            ws_db = wb.Sheets("DATABASE_EXPENDITURE")

        ws_db.Cells.Clear()
        for i in range(ws_db.Shapes.Count, 0, -1):
            ws_db.Shapes.Item(i).Delete()
        while ws_db.ListObjects.Count > 0:
            ws_db.ListObjects.Item(1).Delete()

        # Copy EXPENDITURE header row (row 1) into DB row 1, shifted right by 1 col
        ws_src.Range(
            ws_src.Cells(1, 1), ws_src.Cells(1, TOTAL_COLS)
        ).Copy(ws_db.Cells(1, 2))

        # MONTH header col A
        ws_db.Cells(1, 1).Value = "MONTH"
        ws_db.Cells(1, 1).Font.Bold = True
        ws_db.Cells(1, 1).Font.Size = 11
        ws_db.Cells(1, 1).Interior.Color = excol(0x34, 0x3A, 0x40)
        ws_db.Cells(1, 1).Font.Color = excol(0xFF, 0xFF, 0xFF)
        ws_db.Cells(1, 1).HorizontalAlignment = -4108
        ws_db.Cells(1, 1).VerticalAlignment   = -4108
        ws_db.Cells(1, 1).Borders.LineStyle = 1

        # Apply group colors to header (cols 2+)
        for col in range(1, TOTAL_COLS + 1):
            ws_db.Cells(1, col + 1).Interior.Color = col_color(col)
            ws_db.Cells(1, col + 1).Borders.LineStyle = 1
            ws_db.Cells(1, col + 1).Font.Bold = True

        ws_db.Rows("1:1").RowHeight = 22
        ws_db.Columns(1).ColumnWidth = 12

        # Freeze header
        ws_db.Activate()
        excel.ActiveWindow.FreezePanes = False
        ws_db.Cells(2, 1).Select()
        excel.ActiveWindow.FreezePanes = True

        # ============================================================
        # 3. INJECT VBA from external .bas files
        # ============================================================
        bas_dir = os.path.normpath(
            os.path.join(os.path.dirname(WORKBOOK_PATH), 'scripts_archive')
        )

        sheet_bas = os.path.join(bas_dir, 'vba_sheet_expenditure.bas')
        module_bas = os.path.join(bas_dir, 'vba_module_expenditure.bas')

        with open(sheet_bas, 'r', encoding='utf-8') as f:
            vba_sheet = f.read()

        with open(module_bas, 'r', encoding='utf-8') as f:
            vba_module = f.read()

        # Patch EXP_LAST and EXP_COLS constants
        vba_module = (vba_module
            .replace('Public Const EXP_LAST  As Long = 24',
                     f'Public Const EXP_LAST  As Long = {entry_data_last}')
            .replace('Public Const EXP_COLS  As Long = 27',
                     f'Public Const EXP_COLS  As Long = {TOTAL_COLS}')
        )

        # Sheet-level VBA
        mod_entry = wb.VBProject.VBComponents(ws_entry.CodeName).CodeModule
        mod_entry.DeleteLines(1, mod_entry.CountOfLines)
        mod_entry.AddFromString(vba_sheet)

        # Standard module
        mod_name = "ModExpenditureSystem"
        mod_found = False
        for comp in wb.VBProject.VBComponents:
            if comp.Name == mod_name:
                comp.CodeModule.DeleteLines(1, comp.CodeModule.CountOfLines)
                comp.CodeModule.AddFromString(vba_module)
                mod_found = True
                break
        if not mod_found:
            new_mod = wb.VBProject.VBComponents.Add(1)
            new_mod.Name = mod_name
            new_mod.CodeModule.AddFromString(vba_module)

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
