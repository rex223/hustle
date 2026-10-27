$ErrorActionPreference = "Stop"

Get-Process -Name "excel" -ErrorAction SilentlyContinue | Stop-Process -Force
Start-Sleep -Seconds 2

$filePath = "d:\CLLG\hustle\hustle2\RENTAL DATASHEET 10TH AUG.xlsm"

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

Try {
    $wb = $excel.Workbooks.Open($filePath)
    $ws = $wb.Sheets.Item('RENT DATA')
    $ws.Range('J7:N1000').ClearContents()
    $wb.Save()
    Write-Host 'Cleared!'
} Catch {
    Write-Error $_
} Finally {
    if ($wb) { $wb.Close($false) }
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
}
