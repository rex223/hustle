$ErrorActionPreference = "Stop"
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
try {
    $wb = $excel.Workbooks.Open("d:\CLLG\hustle\hustle2\RENTAL DATASHEET 10TH AUG.xlsm")
    $exportPath = "d:\CLLG\hustle\hustle2\vba_export"
    if (!(Test-Path $exportPath)) { New-Item -ItemType Directory -Path $exportPath | Out-Null }
    foreach ($comp in $wb.VBProject.VBComponents) {
        $exportFile = Join-Path -Path $exportPath -ChildPath ($comp.Name + ".txt")
        $comp.Export($exportFile)
    }
} finally {
    if ($wb) { $wb.Close($false) }
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
}
