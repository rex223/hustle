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
    Try { $ws = $wb.Sheets.Item("RENT DATA") } Catch { }
    if (-not $ws) {
        Try { $ws = $wb.Sheets.Item("todays_rent_data") } Catch { }
    }
    
    if ($ws) {
        $ws.Name = "todays_rent_data"
        
        # Setup the Month Marker safely in I1 and I2 to avoid macro collisions
        $ws.Cells.Item(1, 9).Value2 = "CURRENT MONTH"
        $ws.Cells.Item(2, 9).Value2 = "JULY_2026"
        
        $markerHeader = $ws.Range("I1")
        $markerHeader.Interior.Color = 12157225 # RGB(41, 128, 185)
        $markerHeader.Font.Color = 16777215     # White
        $markerHeader.Font.Bold = $true
        $markerHeader.HorizontalAlignment = -4108
        
        $markerCell = $ws.Range("I2")
        $markerCell.Interior.Color = 14391348 # RGB(52, 152, 219)
        $markerCell.Font.Color = 16777215     # White
        $markerCell.Font.Bold = $true
        $markerCell.Font.Size = 14
        $markerCell.HorizontalAlignment = -4108
        
        # AutoFit Column I
        $ws.Columns("I").AutoFit()
        
        $dbWs = $null
        Try { $dbWs = $wb.Sheets.Item("database_rent_data") } Catch { }
        
        if (-not $dbWs) {
            $ws.Copy([Type]::Missing, $ws)
            $newWs = $wb.ActiveSheet
            $newWs.Name = "database_rent_data"
            $newWs.Range("J2:N1000").ClearContents() # Database doesn't need the summary table calculated yet
        }
        
        Write-Host "Success! Rent Sheets prepared safely."
    } else {
        Write-Host "Could not find RENT DATA sheet."
    }
    
    $wb.Save()
} Catch {
    Write-Error $_
} Finally {
    if ($wb) { $wb.Close($false) }
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
}
