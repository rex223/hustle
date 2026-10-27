$ErrorActionPreference = "Stop"

$filePath = "d:\CLLG\hustle\hustle2\RENTAL DATASHEET 10TH AUG.xlsm"

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

Try {
    $wb = $excel.Workbooks.Open($filePath)
    
    $ws = $null
    Try { $ws = $wb.Sheets.Item("ANAS") } Catch { }
    
    if ($ws) {
        Write-Host "--- ANAS SHEET DATA ---"
        For ($i = 1; $i -le 5; $i++) {
            $rowStr = ""
            For ($j = 1; $j -le 8; $j++) {
                $val = $ws.Cells.Item($i, $j).Value2
                if ($null -eq $val) { $val = "<EMPTY>" }
                $rowStr += "$val`t"
            }
            Write-Host "Row $i`: $rowStr"
        }
    } else {
        Write-Host "ANAS sheet not found!"
    }
    
    $wb.Close($false)
} Catch {
    Write-Error $_
} Finally {
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
}
