$ErrorActionPreference = "Stop"
$excel = [System.Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')
if ($excel) {
    $wb = $excel.ActiveWorkbook
    $ws = $wb.Sheets.Item('RENT DATA')
    $ws.Range('J7:N1000').ClearContents()
    $wb.Save()
    Write-Host 'Cleared!'
}
