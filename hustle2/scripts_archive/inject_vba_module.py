"""
inject_vba_module.py
Injects the updated vba_module_expenditure.bas into the Excel file.
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

        bas_dir = os.path.normpath(
            os.path.join(os.path.dirname(WORKBOOK_PATH), 'scripts_archive')
        )
        module_bas = os.path.join(bas_dir, 'vba_module_expenditure.bas')
        with open(module_bas, 'r', encoding='utf-8') as f:
            vba_module = f.read()

        mod_name = "ModExpenditureSystem"
        mod_found = False
        for comp in wb.VBProject.VBComponents:
            if comp.Name == mod_name:
                comp.CodeModule.DeleteLines(1, comp.CodeModule.CountOfLines)
                comp.CodeModule.AddFromString(vba_module)
                mod_found = True
                break
        
        if not mod_found:
            print(f"Error: {mod_name} not found.")

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
