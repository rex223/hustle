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
    
    foreach ($sheetName in @("todays_expenditure", "database_expenditure")) {
        $ws = $null
        Try { $ws = $wb.Sheets.Item($sheetName) } Catch { }
        
        if ($ws) {
            Write-Host "Formatting $sheetName ..."
            
            # 1. Format Headers (A1:AC1) - Sleek Dark Blue
            $headerRange = $ws.Range("A1:AC1")
            $headerRange.Interior.Color = 12157225 # RGB(41, 128, 185)
            $headerRange.Font.Color = 16777215     # White
            $headerRange.Font.Bold = $true
            $headerRange.HorizontalAlignment = -4108 # xlCenter
            $headerRange.VerticalAlignment = -4108   # xlCenter
            
            # 2. Format Month Marker Row (A2:AC2) - Vivid Blue Separator
            $markerRange = $ws.Range("A2:AC2")
            $markerRange.Interior.Color = 14391348 # RGB(52, 152, 219)
            $markerRange.Font.Color = 16777215     # White
            $markerRange.Font.Bold = $true
            $markerRange.Font.Size = 14
            $markerRange.VerticalAlignment = -4108
            
            # 3. Merge A2:C2 so the month name is clearly visible across 3 columns
            # Using Center Across Selection is safer, but merge is fine for visuals.
            $ws.Range("A2:C2").Merge()
            $ws.Range("A2").HorizontalAlignment = -4108 # xlCenter
            
            # 4. Add some gridlines/borders to make it look clean
            $fullRange = $ws.Range("A1:AC30")
            $fullRange.Borders.LineStyle = 1 # xlContinuous
            $fullRange.Borders.Color = 12632256 # Light Gray RGB(192, 192, 192)
            
            # Re-apply thick border around headers and marker
            $headerRange.Borders.Color = 0 # Black
            $markerRange.Borders.Color = 0 # Black
        }
    }
    
    $wb.Save()
    Write-Host "Done formatting!"
} Catch {
    Write-Error $_
} Finally {
    if ($wb) { $wb.Close($false) }
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
}
