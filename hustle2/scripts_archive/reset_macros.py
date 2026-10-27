import win32com.client
import os

excel = win32com.client.Dispatch('Excel.Application')
excel.Visible = False
file_path = os.path.abspath('RENTAL DATASHEET 12TH AUG.xlsm')
export_dir = os.path.abspath('vba_export')

try:
    wb = excel.Workbooks.Open(file_path)
    
    for comp in wb.VBProject.VBComponents:
        export_file = os.path.join(export_dir, f'{comp.Name}.txt')
        if os.path.exists(export_file):
            try:
                with open(export_file, 'r') as f:
                    code = f.read()
                
                comp.CodeModule.DeleteLines(1, comp.CodeModule.CountOfLines)
                
                if code.strip():
                    comp.CodeModule.AddFromString(code)
                print(f"Reset {comp.Name} successfully.")
            except Exception as e:
                print(f"Failed to reset {comp.Name}: {e}")
                
    wb.Save()
    print("SUCCESS: All macros have been reset to baseline.")
except Exception as e:
    import traceback
    traceback.print_exc()
finally:
    try:
        wb.Close(False)
    except:
        pass
    excel.Quit()
