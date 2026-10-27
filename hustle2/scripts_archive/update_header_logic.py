"""
update_header_logic.py
Updates the UpdateMonthHeader macro to robustly handle both Date values and Text strings.
Run with Excel CLOSED.
"""
import win32com.client
import os

WORKBOOK_PATH = os.path.abspath('hustle2/RENTAL DATASHEET 1ST AUG.xlsm')
BAS_PATH = os.path.abspath('hustle2/scripts_archive/vba_sheet_expenditure.bas')

def fix_header_logic():
    # We will completely replace the UpdateMonthHeader sub in the bas file
    with open(BAS_PATH, 'r', encoding='utf-8') as f:
        vba_code = f.read()
        
    old_sub_start = vba_code.find("Private Sub UpdateMonthHeader()")
    if old_sub_start != -1:
        vba_code = vba_code[:old_sub_start]
        
    new_sub = """Private Sub UpdateMonthHeader()
    Dim val As Variant
    val = Me.Range("C1").Value
    If IsEmpty(val) Or CStr(val) = "" Then Exit Sub
    
    Dim mName As String
    mName = ""
    
    If IsDate(val) Then
        mName = UCase(Format(CDate(val), "MMMM"))
    Else
        Dim lbl As String
        lbl = Trim(CStr(val))
        If InStr(lbl, "-") > 0 Then
            Dim parts() As String
            parts = Split(lbl, "-")
            mName = UCase(Trim(parts(0)))
        Else
            mName = UCase(lbl)
        End If
    End If
    
    If mName <> "" Then
        Me.Cells(2, 5).Value = mName
    End If
End Sub
"""
    vba_code += new_sub
    
    with open(BAS_PATH, 'w', encoding='utf-8') as f:
        f.write(vba_code)

    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    
    try:
        wb = excel.Workbooks.Open(WORKBOOK_PATH)
        ws = wb.Sheets('TODAYS EXPENDITURE')
        mod = wb.VBProject.VBComponents(ws.CodeName).CodeModule
        
        # Replace the entire code module with the new robust code
        mod.DeleteLines(1, mod.CountOfLines)
        mod.AddFromString(vba_code)
        
        # Trigger it manually so it updates right away
        ws.Activate()
        
        wb.Save()
        print("SUCCESS: UpdateMonthHeader macro updated and made robust.")
        
    except Exception as e:
        import traceback
        traceback.print_exc()
        print(f"Error: {e}")
    finally:
        try:
            wb.Close(False)
        except:
            pass
        excel.Quit()

if __name__ == '__main__':
    fix_header_logic()
