"""
fix_vba_duplicates.py
Removes the duplicated AutoPaySalaries macros from the Excel file and the .bas file.
Run with Excel CLOSED.
"""
import win32com.client
import os

WORKBOOK_PATH = os.path.abspath('hustle2/RENTAL DATASHEET 1ST AUG.xlsm')
BAS_PATH = os.path.abspath('hustle2/scripts_archive/vba_module_expenditure.bas')

def fix_duplicates():
    # Read the .bas file and keep only up to line 231
    with open(BAS_PATH, 'r', encoding='utf-8') as f:
        lines = f.readlines()
        
    cleaned_lines = lines[:231]
    
    with open(BAS_PATH, 'w', encoding='utf-8') as f:
        f.writelines(cleaned_lines)
        
    vba_code = "".join(cleaned_lines)

    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    
    try:
        wb = excel.Workbooks.Open(WORKBOOK_PATH)
        
        mod_name = "ModExpenditureSystem"
        for comp in wb.VBProject.VBComponents:
            if comp.Name == mod_name:
                comp.CodeModule.DeleteLines(1, comp.CodeModule.CountOfLines)
                comp.CodeModule.AddFromString(vba_code)
                break
                
        wb.Save()
        print("SUCCESS: Duplicates removed and macro updated.")
        
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
    fix_duplicates()
