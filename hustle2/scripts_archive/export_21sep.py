import win32com.client
import os

xl = win32com.client.Dispatch('Excel.Application')
xl.DisplayAlerts = False
wb = xl.Workbooks.Open(os.path.abspath('RENTAL DATASHEET 21 SEP.xlsm'))

os.makedirs('vba_21sep', exist_ok=True)

for comp in wb.VBProject.VBComponents:
    try:
        lines = comp.CodeModule.CountOfLines
        if lines > 0:
            code = comp.CodeModule.Lines(1, lines)
            with open(f'vba_21sep/{comp.Name}.txt', 'w') as f:
                f.write(code)
            print(f'Exported {comp.Name}')
    except Exception as e:
        print(f'Skipped {comp.Name}: {e}')

wb.Close(False)
xl.Quit()
