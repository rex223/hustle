$ErrorActionPreference = "Stop"

Write-Host "Closing running Excel processes to prevent file lock issues..."
Get-Process -Name "excel" -ErrorAction SilentlyContinue | Stop-Process -Force

$filePath = "d:\CLLG\hustle\hustle2\RENTAL DATASHEET 10TH AUG.xlsm"

Write-Host "Opening Excel..."
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

Try {
    $wb = $excel.Workbooks.Open($filePath)
    
    # Find the sheet (might already be renamed if crashed halfway)
    $ws = $null
    Try { $ws = $wb.Sheets.Item("EXPENDITURE") } Catch { }
    if (-not $ws) {
        Try { $ws = $wb.Sheets.Item("todays_expenditure") } Catch { }
    }
    
    if ($ws) {
        $ws.Name = "todays_expenditure"
        $ws.Cells.Item(1, 5).Value2 = "MONTH"
        
        # Check if A2 is already JULY_2026
        $a2Val = $ws.Cells.Item(2, 1).Value2
        if ($a2Val -ne "JULY_2026") {
            # Clear bottom rows to avoid the "can't push cells off the sheet" error
            $ws.Rows.Item("1000:1048576").Clear()
            
            $ws.Rows.Item(2).Insert()
            $ws.Cells.Item(2, 1).Value2 = "JULY_2026"
        }
        
        # Create database_expenditure if it doesn't exist
        $dbWs = $null
        Try { $dbWs = $wb.Sheets.Item("database_expenditure") } Catch { }
        
        if (-not $dbWs) {
            $ws.Copy([Type]::Missing, $ws)
            $newWs = $wb.ActiveSheet
            $newWs.Name = "database_expenditure"
        }
        
        Write-Host "Success! Sheets prepared."
    } else {
        Write-Host "Could not find EXPENDITURE sheet."
    }
    
    $wb.Save()
} Catch {
    Write-Error $_
} Finally {
    if ($wb) { $wb.Close($false) }
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
}
