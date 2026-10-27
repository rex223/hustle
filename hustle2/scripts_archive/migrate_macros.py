"""
migrate_macros.py
Copies all VBA code from 10TH AUG to 1ST AUG, overwriting existing modules with the same name.
Run with Excel CLOSED.
"""
import win32com.client
import os
import time

FILE_10TH = os.path.abspath('hustle2/RENTAL DATASHEET 10TH AUG.xlsm')
FILE_1ST  = os.path.abspath('hustle2/RENTAL DATASHEET 1ST AUG.xlsm')

def migrate():
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    
    try:
        wb10 = excel.Workbooks.Open(FILE_10TH, ReadOnly=True)
        
        # 1. Extract all macros from 10th Aug
        macros_to_copy = {}
        for comp in wb10.VBProject.VBComponents:
            lines_count = comp.CodeModule.CountOfLines
            if lines_count > 0:
                code = comp.CodeModule.Lines(1, lines_count)
                macros_to_copy[comp.Name] = code
                
        wb10.Close(False)
        
        # 2. Inject into 1st Aug
        wb1 = excel.Workbooks.Open(FILE_1ST)
        
        for comp in wb1.VBProject.VBComponents:
            if comp.Name in macros_to_copy:
                comp.CodeModule.DeleteLines(1, comp.CodeModule.CountOfLines)
                comp.CodeModule.AddFromString(macros_to_copy[comp.Name])
                print(f"Updated {comp.Name} from 10th Aug.")
                
        wb1.Save()
        print("SUCCESS: Macros migrated from 10th Aug to 1st Aug.")
        
    except Exception as e:
        import traceback
        traceback.print_exc()
        print(f"Error: {e}")
    finally:
        try:
            wb1.Close(False)
        except:
            pass
        try:
            wb10.Close(False)
        except:
            pass
        excel.Quit()

if __name__ == '__main__':
    migrate()
