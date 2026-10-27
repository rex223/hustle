$ErrorActionPreference = "Stop"

Write-Host "Closing running Excel processes to prevent file lock issues..."
Get-Process -Name "excel" -ErrorAction SilentlyContinue | Stop-Process -Force

Start-Sleep -Seconds 2

$filePath = "d:\CLLG\hustle\hustle2\RENTAL DATASHEET 10TH AUG.xlsm"

Write-Host "Opening Excel..."
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

Try {
    $wb = $excel.Workbooks.Open($filePath)
    
    $ws = $null
    Try { $ws = $wb.Sheets.Item("EXPENDITURE") } Catch { }
    if (-not $ws) {
        Try { $ws = $wb.Sheets.Item("TODAYS_EXPENDITURE") } Catch { }
    }
    
    if ($ws) {
        $ws.Name = "TODAYS_EXPENDITURE"
        $ws.Cells.Item(1, 5).Value2 = "MONTH"
        
        $a2Val = $ws.Cells.Item(2, 1).Value2
        if ($a2Val -ne "JULY_2026") {
            $ws.Rows.Item("1000:1048576").Clear()
            $ws.Rows.Item(2).Insert()
            $ws.Cells.Item(2, 1).Value2 = "JULY_2026"
        }
        
        $dbWs = $null
        Try { $dbWs = $wb.Sheets.Item("DATABASE_EXPENDITURE") } Catch { }
        
        if (-not $dbWs) {
            $ws.Copy([Type]::Missing, $ws)
            $newWs = $wb.ActiveSheet
            $newWs.Name = "DATABASE_EXPENDITURE"
        } else {
            $dbWs.Cells.Item(1, 5).Value2 = "MONTH"
            $dbA2 = $dbWs.Cells.Item(2, 1).Value2
            if ($dbA2 -ne "JULY_2026") {
                $dbWs.Rows.Item("1000:1048576").Clear()
                $dbWs.Rows.Item(2).Insert()
                $dbWs.Cells.Item(2, 1).Value2 = "JULY_2026"
            }
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
