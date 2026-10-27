"""
fix_and_populate.py
1. Fix TODAYS EXPENDITURE: ensure all 22 rows + fix button positions
2. Populate DATABASE_EXPENDITURE with EXPENDITURE data tagged as Jul-2026
Run with Excel CLOSED.
"""
import win32com.client
import os

WORKBOOK_PATH = os.path.abspath('hustle2/RENTAL DATASHEET 1ST AUG.xlsm')

DATA_LAST_ROW  = 23   # last row in EXPENDITURE (row 1 = header, rows 2-23 = data = 22 people)
TOTAL_COLS     = 27   # A..AA
INPUT_COLS     = [4, 5, 7, 9, 11, 13, 15, 16, 18, 19, 21, 22, 25]

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
        wb = excel.Workbooks.Open(WORKBOOK_PATH)
        ws_src    = wb.Sheets("EXPENDITURE")
        ws_entry  = wb.Sheets("TODAYS EXPENDITURE")
        ws_db     = wb.Sheets("DATABASE_EXPENDITURE")

        entry_data_last = DATA_LAST_ROW + 1   # rows 3..24 in TODAYS EXPENDITURE

        # ============================================================
        # 1. FIX TODAYS EXPENDITURE
        # ============================================================
        ws_entry.Cells.Clear()
        # Remove all shapes (old buttons)
        for i in range(ws_entry.Shapes.Count, 0, -1):
            ws_entry.Shapes.Item(i).Delete()

        # Copy the FULL EXPENDITURE layout (rows 1–23) into entry sheet starting at row 2
        ws_src.Range(
            ws_src.Cells(1, 1), ws_src.Cells(DATA_LAST_ROW, TOTAL_COLS)
        ).Copy(ws_entry.Cells(2, 1))

        # Apply group colors to ALL copied rows (rows 2..24)
        for row in range(2, entry_data_last + 1):
            for col in range(1, TOTAL_COLS + 1):
                ws_entry.Cells(row, col).Interior.Color = col_color(col)
                ws_entry.Cells(row, col).Borders.LineStyle = 1

        # Clear input columns in data rows (3..24) — names/labels stay, amounts clear
        for col in INPUT_COLS:
            ws_entry.Range(
                ws_entry.Cells(3, col),
                ws_entry.Cells(entry_data_last, col)
            ).ClearContents()

        # --- Row 1: Controls bar ---
        ws_entry.Rows("1:1").RowHeight = 30

        ws_entry.Range("A1").Value = "MONTH:"
        ws_entry.Range("A1").Font.Bold = True
        ws_entry.Range("A1").Font.Size = 11
        ws_entry.Range("A1").HorizontalAlignment = -4152  # xlRight
        ws_entry.Range("A1").VerticalAlignment   = -4108  # xlCenter
        ws_entry.Range("A1").Interior.Color = excol(0xF8, 0xF9, 0xFA)

        ws_entry.Range("B1").Interior.Color = excol(0xF8, 0xF9, 0xFA)

        ws_entry.Range("C1").Value = "Aug-2026"
        ws_entry.Range("C1").Font.Bold = True
        ws_entry.Range("C1").Font.Size = 11
        ws_entry.Range("C1").Interior.Color = excol(0xFF, 0xFF, 0x00)
        ws_entry.Range("C1").Borders.LineStyle = 1
        ws_entry.Range("C1").HorizontalAlignment = -4108
        ws_entry.Range("C1").VerticalAlignment   = -4108

        ws_entry.Range("D1").Interior.Color = excol(0xF8, 0xF9, 0xFA)

        # Set default month column header to AUGUST
        ws_entry.Cells(2, 5).Value = "AUGUST"

        # SAVE MONTH button — placed at col E row 1 area using Left/Top from cell position
        # Use Cells(1,5) position to anchor button correctly
        cell_e1 = ws_entry.Range("E1")
        btn_left = cell_e1.Left
        btn_top  = cell_e1.Top + 4
        btn_save = ws_entry.Shapes.AddShape(1, btn_left, btn_top, 100, 22)
        btn_save.Fill.ForeColor.RGB = excol(0x1E, 0x7E, 0x34)
        btn_save.Line.Visible = False
        btn_save.TextFrame.Characters().Text = "SAVE MONTH"
        btn_save.TextFrame.Characters().Font.Bold = True
        btn_save.TextFrame.Characters().Font.Size = 9
        btn_save.TextFrame.Characters().Font.Color = excol(0xFF, 0xFF, 0xFF)
        btn_save.TextFrame.HorizontalAlignment = -4108
        btn_save.TextFrame.VerticalAlignment   = -4108
        btn_save.OnAction = "SaveMonthlyExpenditures"

        # LOAD MONTH button — next to save
        btn_left2 = cell_e1.Left + 108
        btn_load = ws_entry.Shapes.AddShape(1, btn_left2, btn_top, 100, 22)
        btn_load.Fill.ForeColor.RGB = excol(0x00, 0x69, 0xD9)
        btn_load.Line.Visible = False
        btn_load.TextFrame.Characters().Text = "LOAD MONTH"
        btn_load.TextFrame.Characters().Font.Bold = True
        btn_load.TextFrame.Characters().Font.Size = 9
        btn_load.TextFrame.Characters().Font.Color = excol(0xFF, 0xFF, 0xFF)
        btn_load.TextFrame.HorizontalAlignment = -4108
        btn_load.TextFrame.VerticalAlignment   = -4108
        btn_load.OnAction = "LoadMonthlyExpenditures"

        # Freeze rows 1+2
        ws_entry.Activate()
        excel.ActiveWindow.FreezePanes = False
        ws_entry.Cells(3, 1).Select()
        excel.ActiveWindow.FreezePanes = True

        print(f"  TODAYS EXPENDITURE rebuilt: rows 2-{entry_data_last}, buttons placed at E1")

        # ============================================================
        # 2. POPULATE DATABASE_EXPENDITURE WITH JULY DATA
        # ============================================================
        JULY_LABEL = "Jul-2026"

        # Remove any existing Jul-2026 entries
        last_db = ws_db.Cells(ws_db.Rows.Count, 1).End(-4162).Row  # xlUp
        for r in range(last_db, 1, -1):
            cv = str(ws_db.Cells(r, 1).Value).strip()
            if cv == JULY_LABEL or cv == f"--- {JULY_LABEL} ---":
                ws_db.Rows(r).Delete()

        # Find next empty row
        next_row = ws_db.Cells(ws_db.Rows.Count, 1).End(-4162).Row + 1
        if next_row < 2:
            next_row = 2

        # Add section header
        header_range = ws_db.Range(ws_db.Cells(next_row, 1), ws_db.Cells(next_row, TOTAL_COLS + 1))
        header_range.Merge()
        header_range.Cells(1, 1).Value = f"--- {JULY_LABEL} ---"
        header_range.Font.Bold = True
        header_range.Font.Size = 11
        header_range.Interior.Color = excol(0x73, 0x50, 0x49)  # dark warm color
        header_range.Font.Color = excol(0xFF, 0xFF, 0xFF)
        header_range.HorizontalAlignment = -4108
        header_range.VerticalAlignment   = -4108
        ws_db.Rows(next_row).RowHeight = 20
        next_row += 1

        # Copy all 22 data rows from EXPENDITURE (rows 2..23) into DATABASE_EXPENDITURE
        for src_row in range(2, DATA_LAST_ROW + 1):
            # Stamp month in col A
            ws_db.Cells(next_row, 1).Value = JULY_LABEL
            ws_db.Cells(next_row, 1).Interior.Color = excol(0xDC, 0xDC, 0xDC)
            ws_db.Cells(next_row, 1).Borders.LineStyle = 1
            ws_db.Cells(next_row, 1).HorizontalAlignment = -4108

            # Copy values cols A..AA from EXPENDITURE into cols B..AB in DB
            for col in range(1, TOTAL_COLS + 1):
                val = ws_src.Cells(src_row, col).Value
                ws_db.Cells(next_row, col + 1).Value = val
                ws_db.Cells(next_row, col + 1).Borders.LineStyle = 1
                # Apply same group color
                ws_db.Cells(next_row, col + 1).Interior.Color = col_color(col)

            next_row += 1

        ws_db.Columns("A:AB").AutoFit()
        print(f"  DATABASE_EXPENDITURE populated with {DATA_LAST_ROW - 1} rows for {JULY_LABEL}")

        # ============================================================
        # 3. SAVE
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
