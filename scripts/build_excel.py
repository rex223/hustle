"""
build_excel.py
--------------
Builds / rebuilds forms_v3.xlsm from extracted_data.csv.
- Merges existing live rows from the workbook (by CODE) — never overwrites manual entries.
- Adds 4 new transport columns (Nearest Railway Station, Train from Sealdah, Auto Stand, Main Bazaar).
- Raabta Centre is positioned right after NAME OF MADRASAH.
- Password protects the workbook with the supplied password.
- Injects all VBA macros: SubmitToDatabase, SearchMadrasa, FindNearestTransport.

Usage:
    python scripts/build_excel.py
"""

import os, sys, winreg, subprocess
import pandas as pd
from pathlib import Path

# ── Paths ────────────────────────────────────────────────────────────────────
ROOT       = Path(__file__).parent.parent
CSV_PATH   = ROOT / "extracted_data.csv"
OUTPUT     = ROOT / "sheets" / "forms_v4.xlsm"
WB_PASSWORD = "adinul"

# ── Colours (Excel BGR integers) ─────────────────────────────────────────────
def rgb(r, g, b): return r + g*256 + b*65536

# Madrasa Theme (Greens)
C_DARK_GREEN = rgb(2, 44, 34)        # Emerald 950 — headers / banners
C_EMERALD    = rgb(6, 95, 70)        # Emerald 800 — title banner Forms
C_GREEN_BG   = rgb(240, 253, 244)    # Emerald 50  — Dashboard background
C_GREEN_L1   = rgb(209, 250, 229)    # Emerald 100 — Default input cell
C_GREEN_L2   = rgb(167, 243, 208)    # Emerald 200 — Dashboard input cell

C_RED_600    = rgb(220, 38, 38)      # Special col header
C_RED_50     = rgb(254, 242, 242)    # Special col input
C_ORANGE_500 = rgb(249, 115, 22)     # Transport col header
C_ORANGE_50  = rgb(255, 247, 237)    # Transport col input
C_BLUE_600   = rgb(37, 99, 235)      # Auto/Bazaar header (changed from green to blue to contrast with theme)
C_BLUE_50    = rgb(239, 246, 255)    # Auto/Bazaar input
C_GREEN_600  = rgb(22, 163, 74)      # Auto/Bazaar header
C_GREEN_50   = rgb(240, 253, 244)    # Auto/Bazaar input
C_WHITE      = rgb(255, 255, 255)

# Columns that get the red "special" highlight
SPECIAL_COLS   = {"Entry Date", "Raabta Centre", "Municipality/Block-Panchayat"}
# Transport columns — orange for rail, green for auto/bazaar
ORANGE_COLS    = {"Nearest Railway Station", "Train from Sealdah"}
GREEN_COLS     = {"Auto Stand", "Main Bazaar"}


# ── Registry helper ──────────────────────────────────────────────────────────
def enable_vbom():
    try:
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER,
                             r"Software\Microsoft\Office\16.0\Excel\Security",
                             0, winreg.KEY_SET_VALUE)
        winreg.SetValueEx(key, "AccessVBOM", 0, winreg.REG_DWORD, 1)
        winreg.CloseKey(key)
        print("VBA project access enabled.")
        return True
    except Exception as e:
        print(f"Could not set VBOM registry key: {e}")
        return False


# ── Data helpers ─────────────────────────────────────────────────────────────
def load_and_merge(csv_path: Path, existing_wb=None) -> pd.DataFrame:
    """Load CSV and merge in any live rows from the existing workbook."""
    df = pd.read_csv(csv_path, dtype=str).fillna("") if csv_path.exists() else pd.DataFrame()

    # If the workbook already exists, read live rows from the Database sheet
    if existing_wb and Path(existing_wb).exists():
        try:
            live = pd.read_excel(existing_wb, sheet_name="Database",
                                  dtype=str, engine="openpyxl").fillna("")
            # Align columns
            for col in df.columns:
                if col not in live.columns:
                    live[col] = ""
            live = live[df.columns]
            # Append only rows whose (NAME OF MADRASAH, PIN) is not already in CSV
            csv_keys = set(zip(df['NAME OF MADRASAH'].str.strip().str.upper(), df['PIN'].str.strip()))
            def is_duplicate(row):
                name = str(row.get('NAME OF MADRASAH', '')).strip().upper()
                pin = str(row.get('PIN', '')).strip()
                return (name, pin) in csv_keys
            
            new_live = live[~live.apply(is_duplicate, axis=1)]
            if not new_live.empty:
                print(f"Merging {len(new_live)} live row(s) from existing workbook.")
                df = pd.concat([df, new_live], ignore_index=True)
        except Exception as e:
            print(f"Could not read existing workbook for merge: {e}")

    return df


# ── Main builder ─────────────────────────────────────────────────────────────
def build():
    import win32com.client as win32

    if not enable_vbom():
        return

    df = load_and_merge(CSV_PATH, OUTPUT)
    if df.empty:
        print("No data found. Run populate_data.py first.")
        return

    # Add the 3 new columns beside CODE
    if "CODE" in df.columns:
        code_idx = df.columns.get_loc("CODE")
        if "Raabta Code" not in df.columns:
            df.insert(code_idx + 1, "Raabta Code", "")
        if "Total No of Students Involved" not in df.columns:
            df.insert(code_idx + 2, "Total No of Students Involved", "")
        if "Madaris Involved" not in df.columns:
            df.insert(code_idx + 3, "Madaris Involved", "")
            
    # Sort the dataframe so common Raabta codes are adjacent
    if "Raabta Code" in df.columns:
        df = df.sort_values(by=["Raabta Code", "CODE"], na_position='last')


    headers = list(df.columns)
    n_cols  = len(headers)

    # Column index (1-based) lookups for VBA
    def col1(name): return headers.index(name) + 1
    pin_col  = col1('PIN')
    dist_col = col1('District')
    date_col = col1('Entry Date')   # always 1

    print(f"Building workbook with {n_cols} columns and {len(df)} data rows...")

    excel = win32.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    try:
        wb = excel.Workbooks.Add()

        # ── Create sheets ────────────────────────────────────────────────────
        ws_forms = wb.Sheets(1);  ws_forms.Name = "Forms"
        ws_db    = wb.Sheets.Add(After=ws_forms); ws_db.Name = "Database"
        ws_dash  = wb.Sheets.Add(After=ws_db);   ws_dash.Name = "Dashboard"
        
        ws_student_db = wb.Sheets.Add(After=ws_dash); ws_student_db.Name = "Student_DB"
        ws_teacher_db = wb.Sheets.Add(After=ws_student_db); ws_teacher_db.Name = "Teacher_DB"
        ws_madaris_db = wb.Sheets.Add(After=ws_teacher_db); ws_madaris_db.Name = "Madaris_DB"

        # ════════════════════════════════════════════════════════════════════
        # FORMS SHEET
        # ════════════════════════════════════════════════════════════════════
        ws_forms.Activate()
        excel.ActiveWindow.DisplayGridlines = True

        # Title banner rows 1-3
        ws_forms.Range(ws_forms.Cells(1,1), ws_forms.Cells(3, n_cols + 20)).Interior.Color = C_EMERALD
        ws_forms.Cells(2,2).Value = "MADRASA DATA ENTRY PORTAL"
        ws_forms.Cells(2,2).Font.Size  = 22
        ws_forms.Cells(2,2).Font.Bold  = True
        ws_forms.Cells(2,2).Font.Color = C_WHITE
        ws_forms.Cells(2,2).Font.Name  = "Segoe UI"

        # Column headers — row 5; input cells — row 6
        for i, hdr in enumerate(headers):
            col_num = i + 2   # offset by 1 (leave col A empty)
            hcell = ws_forms.Cells(5, col_num)
            icell = ws_forms.Cells(6, col_num)

            hcell.Value       = hdr
            hcell.Font.Bold   = True
            hcell.Font.Color  = C_WHITE
            hcell.Font.Name   = "Segoe UI"
            hcell.Borders.Weight = 2
            icell.Borders.Weight = 2

            if hdr in SPECIAL_COLS:
                hcell.Interior.Color = C_RED_600
                icell.Interior.Color = C_RED_50
            elif hdr in ORANGE_COLS:
                hcell.Interior.Color = C_ORANGE_500
                icell.Interior.Color = C_ORANGE_50
            elif hdr in GREEN_COLS:
                hcell.Interior.Color = C_BLUE_600
                icell.Interior.Color = C_BLUE_50
            else:
                hcell.Interior.Color = C_DARK_GREEN
                icell.Interior.Color = C_GREEN_L1

            # Entry Date: pre-fill with today, formatted dd/mm/yyyy, read-only look
            if hdr == "Entry Date":
                icell.Formula        = "=TODAY()"
                icell.NumberFormat   = "dd/mm/yyyy"
                icell.Font.Color     = rgb(100, 116, 139)

            # PIN: text format so leading zeros survive
            if hdr == "PIN":
                icell.NumberFormat = "@"

        # Hint text row 8
        ws_forms.Cells(8,2).Value       = "Fill the form above and click SUBMIT DATA. Entry Date is auto-filled."
        ws_forms.Cells(8,2).Font.Italic = True
        ws_forms.Cells(8,2).Font.Color  = rgb(100, 116, 139)

        # ── Add 3 Smart Forms (Student, Teacher, Madaris) ──
        form_start_col = n_cols + 4
        
        # Helper to draw a form section
        def draw_form(ws, title, start_col, color, fields):
            ws.Cells(5, start_col).Value = title
            ws.Range(ws.Cells(5, start_col), ws.Cells(5, start_col + 1)).Merge()
            ws.Cells(5, start_col).Font.Size = 14
            ws.Cells(5, start_col).Font.Bold = True
            ws.Cells(5, start_col).Font.Color = C_WHITE
            ws.Cells(5, start_col).Interior.Color = color
            ws.Cells(5, start_col).HorizontalAlignment = -4108 # xlCenter
            
            for idx, field in enumerate(fields):
                r = 7 + idx
                ws.Cells(r, start_col).Value = field
                ws.Cells(r, start_col).Font.Bold = True
                ws.Cells(r, start_col).Interior.Color = rgb(241, 245, 249) # Slate 100
                ws.Cells(r, start_col).Borders.Weight = 2
                
                icell = ws.Cells(r, start_col + 1)
                icell.Interior.Color = C_WHITE
                icell.Borders.Weight = 2
                
            # Add a basic SmartArt/Chart placeholder below the fields
            shape_top = ws.Cells(7 + len(fields) + 2, start_col).Top
            shape_left = ws.Cells(7 + len(fields) + 2, start_col).Left
            shape = ws.Shapes.AddShape(1, shape_left, shape_top, 150, 80) # 1=msoShapeRectangle
            shape.TextFrame2.TextRange.Text = f"{title} Data Viz\\n(SmartArt/Chart Placeholder)"
            shape.TextFrame2.TextRange.Font.Size = 10
            shape.Fill.ForeColor.RGB = rgb(226, 232, 240) # Slate 200
            shape.Line.ForeColor.RGB = color
            shape.Line.Weight = 2
            
            # Submit button for this form
            btn_top = ws.Cells(7 + len(fields) + 2 + 6, start_col).Top
            btn = ws.Buttons().Add(shape_left, btn_top, 100, 30)
            btn.Caption = f"Submit {title.split()[0]}"
            btn.OnAction = f"Submit{title.split()[0]}"

        student_fields = ["Student Name", "Age", "Class / Grade", "Gender", "Notes"]
        draw_form(ws_forms, "Student Information", form_start_col, rgb(37, 99, 235), student_fields) # Blue
        
        teacher_fields = ["Teacher Name", "Qualification", "Subject Taught", "Experience (Yrs)", "Contact"]
        draw_form(ws_forms, "Teacher Information", form_start_col + 4, rgb(147, 51, 234), teacher_fields) # Purple
        
        madaris_fields = ["Facilities", "Infrastructure", "Boarding Details", "Extracurriculars", "Needs"]
        draw_form(ws_forms, "Madaris Information", form_start_col + 8, rgb(217, 119, 6), madaris_fields) # Amber

        ws_forms.Columns.AutoFit()

        # ════════════════════════════════════════════════════════════════════
        # DATABASE SHEET
        # ════════════════════════════════════════════════════════════════════
        ws_db.Activate()
        excel.ActiveWindow.DisplayGridlines = True

        ws_db.Range(ws_db.Cells(1, 1), ws_db.Cells(1, n_cols)).Value = headers
        if "PIN" in headers:
            pin_idx = headers.index("PIN") + 1
            ws_db.Columns(pin_idx).NumberFormat = "@"
            
        if "Entry Date" in headers:
            date_idx = headers.index("Entry Date") + 1
            ws_db.Columns(date_idx).NumberFormat = "dd/mm/yyyy"

        if not df.empty:
            data_body = ws_db.Range(ws_db.Cells(2, 1), ws_db.Cells(len(df) + 1, n_cols))
            data_body.Value = df.values.tolist()

        # Format as named table
        end_cell  = ws_db.Cells(len(df)+1, n_cols)
        db_range  = ws_db.Range(ws_db.Cells(1,1), end_cell)
        tbl       = ws_db.ListObjects.Add(1, db_range, 1, 1)
        tbl.Name  = "tblDatabase"
        tbl.TableStyle = "TableStyleDark10"

        # Apply formulas to the new columns
        if "Total No of Students Involved" in headers:
            idx = headers.index("Total No of Students Involved") + 1
            if len(df) > 0:
                ws_db.Cells(2, idx).Formula = "=SUMIF([Raabta Code], [@[Raabta Code]], [TOTAL S])"
                
        if "Madaris Involved" in headers:
            idx = headers.index("Madaris Involved") + 1
            if len(df) > 0:
                ws_db.Cells(2, idx).Formula = "=COUNTIF([Raabta Code], [@[Raabta Code]])"

        ws_db.Columns.AutoFit()

        # ════════════════════════════════════════════════════════════════════
        # DASHBOARD SHEET
        # ════════════════════════════════════════════════════════════════════
        ws_dash.Activate()
        excel.ActiveWindow.DisplayGridlines = True
        ws_dash.Cells.Interior.Color = C_GREEN_BG

        # Title banner
        ws_dash.Range("A1:Z3").Interior.Color = C_DARK_GREEN
        ws_dash.Cells(2,2).Value      = "MADRASA SEARCH DASHBOARD"
        ws_dash.Cells(2,2).Font.Size  = 24
        ws_dash.Cells(2,2).Font.Bold  = True
        ws_dash.Cells(2,2).Font.Color = C_WHITE
        ws_dash.Cells(2,2).Font.Name  = "Segoe UI"

        # Search panel background
        ws_dash.Range("B5:F9").Interior.Color = C_WHITE
        ws_dash.Range("B5:F9").BorderAround(Weight=3, Color=rgb(203,213,225))

        # PIN search row
        ws_dash.Cells(6,3).Value       = "Search by PIN Code:"
        ws_dash.Cells(6,3).Font.Bold   = True
        ws_dash.Cells(6,3).Font.Size   = 11
        ws_dash.Cells(6,4).Interior.Color   = C_GREEN_L2
        ws_dash.Cells(6,4).Borders.Weight   = 2
        ws_dash.Cells(6,4).NumberFormat     = "@"

        # District dropdown row
        ws_dash.Cells(7,3).Value       = "Search by District:"
        ws_dash.Cells(7,3).Font.Bold   = True
        ws_dash.Cells(7,3).Font.Size   = 11
        ws_dash.Cells(7,4).Interior.Color   = C_GREEN_L2
        ws_dash.Cells(7,4).Borders.Weight   = 2

        # District validation list
        districts = df["District"].dropna().unique().tolist()
        districts = [d for d in districts if d.strip()]
        if districts:
            dv = ws_dash.Cells(7,4).Validation
            dv.Delete()
            dv.Add(Type=3, AlertStyle=1, Operator=1, Formula1=",".join(districts))
            dv.IgnoreBlank   = True
            dv.InCellDropdown = True
            dv.ShowError     = False

        # Results header
        ws_dash.Cells(10,2).Value      = "SEARCH RESULTS:"
        ws_dash.Cells(10,2).Font.Size  = 14
        ws_dash.Cells(10,2).Font.Bold  = True
        ws_dash.Cells(10,2).Font.Color = rgb(30,41,59)

        ws_dash.Columns.AutoFit()
        
        # ════════════════════════════════════════════════════════════════════
        # FORM DATABASE SHEETS (Student, Teacher, Madaris)
        # ════════════════════════════════════════════════════════════════════
        def init_db_sheet(ws, title, fields):
            ws.Activate()
            excel.ActiveWindow.DisplayGridlines = True
            headers = fields + ["Timestamp"]
            ws.Range(ws.Cells(1, 1), ws.Cells(1, len(headers))).Value = headers
            # Format as named table
            tbl_range = ws.Range(ws.Cells(1, 1), ws.Cells(2, len(headers)))
            tbl = ws.ListObjects.Add(1, tbl_range, 1, 1)
            tbl.Name = f"tbl{title.split()[0]}"
            tbl.TableStyle = "TableStyleDark10"
            ws.Columns.AutoFit()
            
        init_db_sheet(ws_student_db, "Student Information", student_fields)
        init_db_sheet(ws_teacher_db, "Teacher Information", teacher_fields)
        init_db_sheet(ws_madaris_db, "Madaris Information", madaris_fields)


        # ════════════════════════════════════════════════════════════════════
        # VBA MACROS
        # ════════════════════════════════════════════════════════════════════
        vba = f"""
Sub SubmitToDatabase()
    Dim wsForm As Worksheet
    Dim wsDb   As Worksheet
    Dim tbl    As ListObject
    Dim newRow As ListRow
    Dim c      As Long
    Dim r      As Long: r = 6

    Set wsForm = ThisWorkbook.Sheets("Forms")
    Set wsDb   = ThisWorkbook.Sheets("Database")
    Set tbl    = wsDb.ListObjects("tblDatabase")

    ' Validate mandatory fields
    Dim missing As String
    missing = ""
    
    If Trim(wsForm.Cells(r, 3).Value) = "" Then missing = missing & "- CODE" & vbCrLf
    If Trim(wsForm.Cells(r, 4).Value) = "" Then missing = missing & "- NAME OF MADRASAH" & vbCrLf
    If Trim(wsForm.Cells(r, 7).Value) = "" Then missing = missing & "- VILLAGE" & vbCrLf
    If Trim(wsForm.Cells(r, 10).Value) = "" Then missing = missing & "- PIN Code" & vbCrLf
    If Trim(wsForm.Cells(r, 14).Value) = "" Then missing = missing & "- District" & vbCrLf
    
    If missing <> "" Then
        MsgBox "Please fill the following mandatory fields before submitting:" & vbCrLf & vbCrLf & missing, vbExclamation, "Missing Information"
        Exit Sub
    End If

    Set newRow = tbl.ListRows.Add

    For c = 1 To {n_cols}
        If c = {date_col} Then
            ' Auto-stamp today's date
            newRow.Range(c).Value         = Date
            newRow.Range(c).NumberFormat  = "dd/mm/yyyy"
        Else
            newRow.Range(c).Value = wsForm.Cells(r, c + 1).Value
        End If
    Next c

    ' Clear form (skip Entry Date cell — it stays =TODAY())
    Dim clr As Long
    For clr = 2 To {n_cols + 1}
        If clr <> {date_col + 1} Then wsForm.Cells(r, clr).ClearContents
    Next clr

    MsgBox "Record submitted successfully with today's date!", vbInformation, "Success"
End Sub


Sub SearchMadrasa()
    Dim wsDash As Worksheet
    Dim wsDb   As Worksheet
    Dim tbl    As ListObject
    Dim pinVal As String
    Dim distVal As String
    Dim res    As Range

    Set wsDash = ThisWorkbook.Sheets("Dashboard")
    Set wsDb   = ThisWorkbook.Sheets("Database")
    Set tbl    = wsDb.ListObjects("tblDatabase")

    pinVal  = Trim(wsDash.Cells(6, 4).Text)
    distVal = Trim(wsDash.Cells(7, 4).Text)

    ' Mutual exclusivity
    If pinVal <> "" And distVal <> "" Then
        MsgBox "Please search by EITHER PIN Code OR District, not both.", vbExclamation, "Search Error"
        wsDash.Cells(6, 4).ClearContents
        wsDash.Cells(7, 4).ClearContents
        Exit Sub
    End If

    If pinVal = "" And distVal = "" Then
        MsgBox "Please enter a PIN Code or select a District first.", vbExclamation, "Empty Search"
        Exit Sub
    End If

    ' Clear previous results
    wsDash.Range("B12:Z2000").ClearContents
    wsDash.Range("B12:Z2000").Interior.ColorIndex = xlNone
    wsDash.Range("B12:Z2000").Borders.LineStyle = xlNone

    ' Remove any existing filter
    If tbl.ShowAutoFilter Then
        If tbl.AutoFilter.FilterMode Then tbl.AutoFilter.ShowAllData
    End If

    ' Apply filter
    If pinVal <> "" Then
        tbl.Range.AutoFilter Field:={pin_col}, Criteria1:="=" & pinVal
    ElseIf distVal <> "" Then
        tbl.Range.AutoFilter Field:={dist_col}, Criteria1:="*" & distVal & "*"
    End If

    ' Copy visible rows
    On Error Resume Next
    Set res = tbl.Range.SpecialCells(xlCellTypeVisible)
    On Error GoTo 0

    If Not res Is Nothing Then
        res.Copy Destination:=wsDash.Range("B12")
        ' Style the header row
        With wsDash.Range("B12").EntireRow
            .Interior.Color = 32896 ' C_EMERALD color equivalent in VBA
            .Font.Color     = 16777215 ' white
            .Font.Bold      = True
        End With
        wsDash.Columns.AutoFit
    Else
        With wsDash.Cells(12, 2)
            .Value      = "No results found for your search."
            .Font.Italic = True
            .Font.Color  = vbRed
        End With
    End If

    ' Reset filter so Database looks normal
    If tbl.ShowAutoFilter Then
        If tbl.AutoFilter.FilterMode Then tbl.AutoFilter.ShowAllData
    End If
End Sub

Sub SubmitStudent()
    SubmitSmartForm "Student", {form_start_col}
End Sub

Sub SubmitTeacher()
    SubmitSmartForm "Teacher", {form_start_col + 4}
End Sub

Sub SubmitMadaris()
    SubmitSmartForm "Madaris", {form_start_col + 8}
End Sub

Sub SubmitSmartForm(formName As String, startCol As Long)
    Dim wsForm As Worksheet, wsDb As Worksheet
    Dim tbl As ListObject
    Dim newRow As ListRow
    Dim r As Long: r = 7
    Dim c As Long
    
    Set wsForm = ThisWorkbook.Sheets("Forms")
    Set wsDb = ThisWorkbook.Sheets(formName & "_DB")
    Set tbl = wsDb.ListObjects("tbl" & formName)
    
    ' Basic check - first field not empty
    If Trim(wsForm.Cells(r, startCol + 1).Value) = "" Then
        MsgBox "Please fill out the first field before submitting.", vbExclamation
        Exit Sub
    End If
    
    Set newRow = tbl.ListRows.Add
    
    ' Read 5 fields
    For c = 1 To 5
        newRow.Range(c).Value = wsForm.Cells(r + c - 1, startCol + 1).Value
        wsForm.Cells(r + c - 1, startCol + 1).ClearContents
    Next c
    
    ' Timestamp
    newRow.Range(6).Value = Now
    
    MsgBox formName & " record submitted successfully!", vbInformation, "Success"
End Sub
"""
        mod = wb.VBProject.VBComponents.Add(1)
        mod.CodeModule.AddFromString(vba)

        # ── Buttons ──────────────────────────────────────────────────────────
        # Forms sheet — Submit button
        btn_sub = ws_forms.Buttons().Add(500, 145, 160, 35)
        btn_sub.Caption  = "  SUBMIT DATA  "
        btn_sub.OnAction = "SubmitToDatabase"

        # Dashboard — Search button
        # Moved button down to row 9 so it doesn't overlay the search boxes
        btn_sch = ws_dash.Buttons().Add(250, 150, 130, 32)
        btn_sch.Caption  = "  SEARCH  "
        btn_sch.OnAction = "SearchMadrasa"

        # ── Password protect & save ─────────────────────────────────────────
        wb.SaveAs(
            str(OUTPUT.resolve()),
            FileFormat=52,          # xlOpenXMLWorkbookMacroEnabled
            Password=WB_PASSWORD
        )
        wb.SaveAs(str(OUTPUT), FileFormat=52)
        print(f"\nWorkbook saved: {OUTPUT}")
        print(f"Password: {WB_PASSWORD}")

    except Exception as e:
        print(f"ERROR: {e}")
        import traceback; traceback.print_exc()
    finally:
        excel.Quit()


if __name__ == "__main__":
    build()
