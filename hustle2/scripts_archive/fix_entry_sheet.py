"""
fix_entry_sheet.py
- Turns off AutoFilter on EXPENDITURE before copy
- Copies all 22 rows into TODAYS EXPENDITURE correctly
- Places SAVE MONTH / LOAD MONTH buttons visibly in row 1
Run with Excel CLOSED.
"""
import win32com.client
import os

WORKBOOK_PATH = os.path.abspath('hustle2/RENTAL DATASHEET 1ST AUG.xlsm')

DATA_LAST_ROW = 23   # EXPENDITURE: header=row1, data=rows2-23 (22 people)
TOTAL_COLS    = 27   # A..AA
INPUT_COLS    = [4, 5, 7, 9, 11, 13, 15, 16, 18, 19, 21, 22, 25]

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
        ws_src   = wb.Sheets("EXPENDITURE")
        ws_entry = wb.Sheets("TODAYS EXPENDITURE")

        entry_data_last = DATA_LAST_ROW + 1  # rows 3..24 in TODAYS EXPENDITURE

        # ----------------------------------------------------------------
        # 0. Turn off AutoFilter on source sheet so all rows are copied
        # ----------------------------------------------------------------
        if ws_src.AutoFilterMode:
            ws_src.AutoFilterMode = False
            print("  AutoFilter turned off on EXPENDITURE")

        # ----------------------------------------------------------------
        # 1. Clear TODAYS EXPENDITURE fully
        # ----------------------------------------------------------------
        ws_entry.Cells.Clear()
        for i in range(ws_entry.Shapes.Count, 0, -1):
            ws_entry.Shapes.Item(i).Delete()
        while ws_entry.ListObjects.Count > 0:
            ws_entry.ListObjects.Item(1).Delete()

        # Turn off freeze panes first
        ws_entry.Activate()
        excel.ActiveWindow.FreezePanes = False

        # ----------------------------------------------------------------
        # 2. Copy ALL rows from EXPENDITURE into entry sheet starting row 2
        # ----------------------------------------------------------------
        src_range = ws_src.Range(
            ws_src.Cells(1, 1),
            ws_src.Cells(DATA_LAST_ROW, TOTAL_COLS)
        )
        src_range.Copy(ws_entry.Cells(2, 1))
        print(f"  Copied rows 1-{DATA_LAST_ROW} from EXPENDITURE into TODAYS EXPENDITURE rows 2-{entry_data_last}")

        # ----------------------------------------------------------------
        # 3. Apply group colors to ALL copied rows (rows 2..24)
        # ----------------------------------------------------------------
        for row in range(2, entry_data_last + 1):
            for col in range(1, TOTAL_COLS + 1):
                ws_entry.Cells(row, col).Interior.Color = col_color(col)
                ws_entry.Cells(row, col).Borders.LineStyle = 1

        # ----------------------------------------------------------------
        # 4. Clear input columns in data rows (3..24) — keep static labels
        # ----------------------------------------------------------------
        for col in INPUT_COLS:
            ws_entry.Range(
                ws_entry.Cells(3, col),
                ws_entry.Cells(entry_data_last, col)
            ).ClearContents()

        # Rename col E header to AUGUST (current month)
        ws_entry.Cells(2, 5).Value = "AUGUST"

        # ----------------------------------------------------------------
        # 5. Setup Row 1 — Controls bar
        # ----------------------------------------------------------------
        ws_entry.Rows("1:1").RowHeight = 35  # Tall enough for button

        # Label in A1
        ws_entry.Range("A1").Value = "MONTH:"
        ws_entry.Range("A1").Font.Bold = True
        ws_entry.Range("A1").Font.Size = 12
        ws_entry.Range("A1").HorizontalAlignment = -4152  # xlRight
        ws_entry.Range("A1").VerticalAlignment   = -4108  # xlCenter
        ws_entry.Range("A1").Interior.Color = excol(0xF0, 0xF0, 0xF0)

        ws_entry.Range("B1").Interior.Color = excol(0xF0, 0xF0, 0xF0)

        # Yellow month selector C1
        ws_entry.Range("C1").Value = "Aug-2026"
        ws_entry.Range("C1").Font.Bold = True
        ws_entry.Range("C1").Font.Size = 12
        ws_entry.Range("C1").Interior.Color = excol(0xFF, 0xFF, 0x00)
        ws_entry.Range("C1").Borders.LineStyle = 1
        ws_entry.Range("C1").HorizontalAlignment = -4108
        ws_entry.Range("C1").VerticalAlignment   = -4108

        ws_entry.Range("D1").Interior.Color = excol(0xF0, 0xF0, 0xF0)

        # ----------------------------------------------------------------
        # 6. Add SAVE + LOAD buttons anchored to cell positions in row 1
        # ----------------------------------------------------------------
        # Use cell Left/Top from E1 and F1 to place buttons precisely
        cell_e1  = ws_entry.Range("E1")
        cell_g1  = ws_entry.Range("G1")

        btn_top  = cell_e1.Top  + 5
        btn_h    = ws_entry.Rows("1:1").RowHeight - 10

        # SAVE button at E1
        btn_save = ws_entry.Shapes.AddShape(1, cell_e1.Left, btn_top, cell_e1.Width * 2, btn_h)
        btn_save.Fill.ForeColor.RGB = excol(0x1E, 0x7E, 0x34)
        btn_save.Line.Visible = False
        btn_save.TextFrame.Characters().Text = "SAVE MONTH"
        btn_save.TextFrame.Characters().Font.Bold = True
        btn_save.TextFrame.Characters().Font.Size = 10
        btn_save.TextFrame.Characters().Font.Color = excol(0xFF, 0xFF, 0xFF)
        btn_save.TextFrame.HorizontalAlignment = -4108
        btn_save.TextFrame.VerticalAlignment   = -4108
        btn_save.OnAction = "SaveMonthlyExpenditures"

        # LOAD button at G1
        btn_load = ws_entry.Shapes.AddShape(1, cell_g1.Left, btn_top, cell_g1.Width * 2, btn_h)
        btn_load.Fill.ForeColor.RGB = excol(0x00, 0x69, 0xD9)
        btn_load.Line.Visible = False
        btn_load.TextFrame.Characters().Text = "LOAD MONTH"
        btn_load.TextFrame.Characters().Font.Bold = True
        btn_load.TextFrame.Characters().Font.Size = 10
        btn_load.TextFrame.Characters().Font.Color = excol(0xFF, 0xFF, 0xFF)
        btn_load.TextFrame.HorizontalAlignment = -4108
        btn_load.TextFrame.VerticalAlignment   = -4108
        btn_load.OnAction = "LoadMonthlyExpenditures"

        print(f"  Buttons placed: SAVE at E1 (left={cell_e1.Left}), LOAD at G1 (left={cell_g1.Left})")

        # ----------------------------------------------------------------
        # 7. Freeze rows 1 and 2 (controls + column headers always visible)
        # ----------------------------------------------------------------
        ws_entry.Activate()
        ws_entry.Cells(3, 1).Select()
        excel.ActiveWindow.FreezePanes = True
        print("  Freeze panes set at row 3")

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
