"""
fix_purchase_validation.py
Changes the data validation on PURCHASE SHEET C1 to allow manual entry of new dates.
Run with Excel CLOSED.
"""
import win32com.client
import os

WORKBOOK_PATH = os.path.abspath('hustle2/RENTAL DATASHEET 1ST AUG.xlsm')

def fix_validation():
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    
    try:
        wb = excel.Workbooks.Open(WORKBOOK_PATH)
        ws = wb.Sheets('PURCHASE SHEET')
        mod = wb.VBProject.VBComponents(ws.CodeName).CodeModule
        
        # Read existing code
        code = mod.Lines(1, mod.CountOfLines)
        
        # Replace the strict validation with relaxed validation
        old_validation = ".Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:=datesStr"
        new_validation = ".Add Type:=xlValidateList, AlertStyle:=xlValidAlertInformation, Formula1:=datesStr\n            .ShowError = False"
        
        if old_validation in code:
            code = code.replace(old_validation, new_validation)
            
            # Clear old code and write new code
            mod.DeleteLines(1, mod.CountOfLines)
            mod.AddFromString(code)
            
            # Delete any existing strict validation on the sheet immediately
            try:
                ws.Range("C1").Validation.Delete()
            except:
                pass
            
            # Trigger Worksheet_Activate to re-apply the new relaxed validation
            ws.Activate()
            
            wb.Save()
            print("SUCCESS: PURCHASE SHEET validation fixed.")
        else:
            print("Target validation string not found in VBA code.")
        
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
    fix_validation()
