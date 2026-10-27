import win32com.client
import os
import shutil

def extract_macros(file_name, output_dir):
    excel = win32com.client.Dispatch('Excel.Application')
    excel.Visible = False
    file_path = os.path.abspath(file_name)
    output_path = os.path.abspath(output_dir)
    
    if os.path.exists(output_path):
        shutil.rmtree(output_path)
    os.makedirs(output_path)
    
    try:
        wb = excel.Workbooks.Open(file_path)
        for comp in wb.VBProject.VBComponents:
            lines = comp.CodeModule.CountOfLines
            if lines > 0:
                code = comp.CodeModule.Lines(1, lines)
                with open(os.path.join(output_path, f'{comp.Name}.txt'), 'w', encoding='utf-8') as f:
                    f.write(code)
        print(f"Extracted macros from {file_name} to {output_dir}")
    except Exception as e:
        print(f"Failed to extract from {file_name}: {e}")
    finally:
        try:
            wb.Close(False)
        except:
            pass
        excel.Quit()

if __name__ == '__main__':
    extract_macros('RENTAL DATASHEET 26TH AUG.xlsm', 'vba_26_aug')
    extract_macros('RENTAL DATASHEET 3 SEP.xlsm', 'vba_3_sep')
