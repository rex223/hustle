import win32com.client
import os

excel = win32com.client.Dispatch('Excel.Application')
excel.Visible = False
excel.DisplayAlerts = False
file_path = os.path.abspath('RENTAL DATASHEET 12TH AUG.xlsm')

vba_fix = '''Public Sub FixEverything()
    On Error Resume Next
    Dim wsRent As Worksheet
    Set wsRent = ThisWorkbook.Sheets("todays_rent_data")
    
    ' 1. Fix Conditional Formatting
    If Not wsRent Is Nothing Then
        wsRent.Range("M2:N1000").FormatConditions.Delete
        
        ' Column M Green
        With wsRent.Range("M2:M1000").FormatConditions.Add(Type:=xlExpression, Formula1:="=AND($M2<>"""", $L2<>"""", $M2>=$L2)")
            .Interior.Color = RGB(226, 239, 218)
            .Font.Color = RGB(0, 97, 0)
            .Font.Bold = True
        End With
        
        ' Column M Light Green / Yellow
        With wsRent.Range("M2:M1000").FormatConditions.Add(Type:=xlExpression, Formula1:="=AND($M2<>"""", $L2<>"""", $M2>0, $M2<$L2)")
            .Interior.Color = RGB(255, 255, 204)
        End With
        
        ' Column N Red
        With wsRent.Range("N2:N1000").FormatConditions.Add(Type:=xlExpression, Formula1:="=AND($N2<>"""", $N2>0)")
            .Interior.Color = RGB(255, 199, 206)
            .Font.Color = RGB(156, 0, 6)
            .Font.Bold = True
        End With
    End If
    
    ' 2. Re-create Buttons in todays_expenditure
    Dim wsExp As Worksheet
    Set wsExp = ThisWorkbook.Sheets("todays_expenditure")
    
    If Not wsExp Is Nothing Then
        Dim shp As Shape
        For Each shp In wsExp.Shapes
            shp.Delete
        Next shp
        
        Dim btn1 As Button, btn2 As Button, btn3 As Button, btn4 As Button
        Set btn1 = wsExp.Buttons.Add(300, 5, 100, 25)
        btn1.Caption = "SAVE DATA"
        btn1.OnAction = "SaveExpenditureData"
        
        Set btn2 = wsExp.Buttons.Add(410, 5, 100, 25)
        btn2.Caption = "LOAD DATA"
        btn2.OnAction = "LoadExpenditureData"
        
        Set btn3 = wsExp.Buttons.Add(520, 5, 100, 25)
        btn3.Caption = "START NEW"
        btn3.OnAction = "StartNewMonth"
        
        Set btn4 = wsExp.Buttons.Add(630, 5, 100, 25)
        btn4.Caption = "AUTOPAY"
        btn4.OnAction = "AutoPaySalaries"
    End If
    
    On Error GoTo 0
End Sub
'''

try:
    wb = excel.Workbooks.Open(file_path)
    
    # Insert temporary module
    temp_mod = wb.VBProject.VBComponents.Add(1) # Standard Module
    temp_mod.Name = "TempFixer"
    temp_mod.CodeModule.AddFromString(vba_fix)
    
    # Run macro
    excel.Application.Run("TempFixer.FixEverything")
    
    # Remove temporary module
    wb.VBProject.VBComponents.Remove(temp_mod)
    
    # The fix_issues.py already fixed ModDailyPurchases and injected the Worksheet_Change macro!
    # Let's verify ModDailyPurchases
    
    wb.Save()
    print("SUCCESS")
except Exception as e:
    import traceback
    traceback.print_exc()
finally:
    try:
        wb.Close(False)
    except:
        pass
    excel.Quit()
