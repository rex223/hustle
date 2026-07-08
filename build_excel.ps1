# build_excel.ps1
# Automates Excel to build forms.xlsm with Forms, Database, Dashboard, Slicers, and VBA macro.

$ErrorActionPreference = "Stop"

function Get-ExcelColor($R, $G, $B) {
    return [int]($R + ($G * 256) + ($B * 65536))
}

$savePath = "d:\CLLG\hustle\forms.xlsm"

# If the file already exists, let's delete it to overwrite cleanly
if (Test-Path $savePath) {
    Remove-Item $savePath -Force
    Write-Host "Existing forms.xlsm deleted for clean generation."
}

Write-Host "Initializing Excel Application..."
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

try {
    # Get Excel version to modify the registry
    $excelVersion = $excel.Version
    Write-Host "Excel Version: $excelVersion"
    
    $registryPath = "HKCU:\Software\Microsoft\Office\$excelVersion\Excel\Security"
    $regValueName = "AccessVBOM"
    $originalValue = $null
    
    # Enable AccessVBOM temporarily to allow script injection of VBA macro
    if (Test-Path $registryPath) {
        $originalValue = (Get-ItemProperty -Path $registryPath -Name $regValueName -ErrorAction SilentlyContinue).$regValueName
    } else {
        New-Item -Path $registryPath -Force | Out-Null
    }
    Set-ItemProperty -Path $registryPath -Name $regValueName -Value 1 -Type DWord -Force | Out-Null
    Write-Host "VBA Project Access temporarily enabled in registry."

    Write-Host "Creating Workbook..."
    $wb = $excel.Workbooks.Add()
    $sheets = $wb.Sheets
    
    # Ensure exactly 3 sheets
    while ($sheets.Count -lt 3) {
        $sheets.Add() | Out-Null
    }
    while ($sheets.Count -gt 3) {
        $sheets.Item(1).Delete() | Out-Null
    }
    
    $wsForm = $sheets.Item(1)
    $wsForm.Name = "Forms"
    
    $wsDb = $sheets.Item(2)
    $wsDb.Name = "Database"
    
    $wsDashboard = $sheets.Item(3)
    $wsDashboard.Name = "Dashboard"

    # Define the 22 preloaded records extracted from images
    $data = @(
        [PSCustomObject]@{Code="S24PGS 453"; Name="DARUL ULOOM ASRARIA SANTOSH PUR"; Estd=2010; Village="FAKIRPARA"; POPS="SANTOSH PUR - MAHESHTALA"; PIN="700142"; Syllabus="HIFZ"; Expenditure=10520038; Words="ONE CRORE FIVE LAKH TWENTY THOUSAND THIRTY EIGHT"; Head="MAULANA NAUSHER AHMAD"; Total=389; Boarding=345},
        [PSCustomObject]@{Code="S24PGS 744"; Name="MADRASA INAYATUL ISLAM"; Estd=1985; Village="BARANDARY"; POPS="B.BAGARIA - KULPI"; PIN="743332"; Syllabus="HIFZ TO ARBI AOWAL"; Expenditure=3500000; Words="THIRTY FIVE LAKH"; Head="AIJUR R.QASEMI"; Total=158; Boarding=158},
        [PSCustomObject]@{Code="S24PGS 745"; Name="BAZE ORANCHA JUNIOR MADRASA BAITUL ULOOM MOHAMMADIA"; Estd=1993; Village="BAZE ORANCHA"; POPS="SITAKUNDU - BARUI PUR"; PIN="743387"; Syllabus="HIFZ TO ARBI SOM"; Expenditure=970664; Words="NINE LAKH SEVENTY THOUSAND SIX HUNDRED SIXTY FOUR"; Head="MD HOSSAIN ALI MOLLA"; Total=93; Boarding=46},
        [PSCustomObject]@{Code="S24PGS 746"; Name="DIGBERIA MADRASATUNNABABI"; Estd=1985; Village="DIGBERIA"; POPS="DIGBERIA - MANDIR BAZAR"; PIN="743395"; Syllabus="HIFZ"; Expenditure=1023256; Words="TEN LAKH TWENTY THREE THOUSAND TWO HUNDRED VISY SIX"; Head="HAFIZ NURUL HUDA HALDAR"; Total=105; Boarding=75},
        [PSCustomObject]@{Code="S24PGS 747"; Name="KALASH MADRASAHA MAZAHIRUL ULOOM"; Estd=1989; Village="KALAS"; POPS="KALASH - MAGRAHAT"; PIN="743609"; Syllabus="HIFZ TO ARBI PANJUM"; Expenditure=8817333; Words="THIRTY FIVE LAKH"; Head="M. AYNUL HAQUE SAHEB"; Total=694; Boarding=578},
        [PSCustomObject]@{Code="S24PGS 748"; Name="SUNDARBON MADRASA MAZAHIRUL ULOOM TRUST"; Estd=1952; Village="KALIDANGA"; POPS="BASANTI"; PIN="743312"; Syllabus="HIFZ TO ARBI DOM"; Expenditure=1272509; Words="TWELVE LAKH SEVENTY TWO THOUSAND FIVE HUNDRED NINE"; Head="EMDADULLAH KHAN"; Total=116; Boarding=80},
        [PSCustomObject]@{Code="S24PGS 749"; Name="MADRASAH MUJADDIDUL ULOOM PANCHA GRAM"; Estd=2003; Village="PANCHA GRAM"; POPS="PANCHAGRAM - DIAMONDHARBAR"; PIN="743368"; Syllabus="HIFA"; Expenditure=2472620; Words="TWENTY FOUR LAKH SEVENTY TWO THOUSAND SIX HUNDRED TWENTY"; Head="M SHAHABUDDIN GHORAMI"; Total=111; Boarding=90},
        [PSCustomObject]@{Code="S24PGS 750"; Name="SUNDARBAN ISLAMIA MADRASA"; Estd=1995; Village="TENTUL TALI"; POPS="PATHANKHALI - GOSABA"; PIN="743611"; Syllabus="HIFZ TO ARBI DOM"; Expenditure=7569314; Words="SEVENTY FIVE LAKH SIXTY NINE THOUSAND THREE HUNDRED FOURTEEN"; Head="MAULANA ABUL WAJED"; Total=376; Boarding=348},
        [PSCustomObject]@{Code="S24PGS 752"; Name="MADRASA JAMIA ISLAMIA NAJMUL HUDA TRUST"; Estd=1972; Village="UTTAR GRAN BOSE"; POPS="BHARAT GARH - BASANTI"; PIN="743312"; Syllabus="HIFZ"; Expenditure=630000; Words="SIX LAKH THIRTY THOUSAND"; Head="NASIM SAGARKH"; Total=42; Boarding=36},
        [PSCustomObject]@{Code="S24PGS 753"; Name="SUNDARBAN ASHRTIFUL ULOOM MADRASHA"; Estd=1966; Village="SERHANGAM PUR"; POPS="NIMPITH - BAKULTALA"; PIN="743338"; Syllabus="HIFZ"; Expenditure=1743817; Words="SEVENTEEN LAKH FORTY THREE THOUSAND EIGHT HUNDRED SEVENTEEN"; Head="ABDUR RAHIM MOLLA"; Total=75; Boarding=75},
        [PSCustomObject]@{Code="S24PGS 754"; Name="AMRATALA MADRASA TABLIGUL ULOOM"; Estd=1962; Village="AMRATALA"; POPS="HARANANDAPUR - MAGRAHAT"; PIN="743503"; Syllabus="HIFZ"; Expenditure=850000; Words="EIGHT LAKH FIFTY THOUSAND"; Head="MD ZAKARIA MOLLA"; Total=70; Boarding=50},
        [PSCustomObject]@{Code="S24PGS 755"; Name="JAMIA RAHAMANIA DARUL ULOOM BANTRA"; Estd=1946; Village="BANTRA"; POPS="BANTRA - JAYNAGAR"; PIN="743337"; Syllabus="HIFZ TO ARBI SOM"; Expenditure=8549205; Words="EIGHTY FIVE LAKH FORTY NINE THOUSAND TWO HUNDRED FIVE"; Head="MAWLANA ABDUL KAHHAR MAZAHARI"; Total=280; Boarding=280},
        [PSCustomObject]@{Code="S24PGS 756"; Name="JAMIA ISLAMIA BAHRUL ULOOM"; Estd=1965; Village="BAKULTALA"; POPS="BAKULTALA HAT"; PIN="743338"; Syllabus="HIFZ TO DAWRA E HADIS"; Expenditure=11990903; Words="ONE CRORE NINETEEN LAKH NINETY THOUSAND NINE HUNDRED THREE"; Head="MAULANA SAFIULLAH"; Total=570; Boarding=518},
        [PSCustomObject]@{Code="S24PGS 757"; Name="BANKIPUR MADRASA DARUL ULOOM ILIYASIA"; Estd=1970; Village="BANKIPUR"; POPS="MAGRAHAT"; PIN="743355"; Syllabus="HIFZ TO DAWRA E HADIS"; Expenditure=11095217; Words="ONE CRORE TEN LAKH NINETY FIVE THOUSAND TWO HUNDRED SEVENTEEN"; Head="HAFIZ ABDUL KHABIR"; Total=793; Boarding=576},
        [PSCustomObject]@{Code="S24PGS 758"; Name="MADRASAH DARUS SALAM IDGHA"; Estd=1961; Village="CHALTABERIA"; POPS="CHALTABERIA - JAYNAGAR"; PIN="743337"; Syllabus="HIFZ TO KAFIA"; Expenditure=4296011; Words="FORTY TWO LAKH NINETY SIX THOUSAND ELEVEN"; Head="MUFTI AMINUDDIN QASIMI"; Total=235; Boarding=229},
        [PSCustomObject]@{Code="S24PGS 759"; Name="HAQOANIA MAHMOODIA FAYZUL ULOOM"; Estd=1982; Village="CHAKARBERIA"; POPS="KUNDARALI - BARUIPUR"; PIN="743610"; Syllabus="HIFZ TO ARBI DOM"; Expenditure=4866020; Words="FORTY EIGHT LAKH SIXTY SIX THOUSAND TWENTY"; Head="MUFTI ABDUL QAIYUM"; Total=177; Boarding=177},
        [PSCustomObject]@{Code="S24PGS 760"; Name="MADRASA DARUL ULOOM BAISHATA"; Estd=1968; Village="BAISHATA"; POPS="BAISHATA - BAKULTALA"; PIN="743338"; Syllabus="HIFZ TO KAFIA"; Expenditure=2615360; Words="TWENTY SIX LAKH FIFTEEN THOUSAND THREE HUNDRED SIXTY"; Head="MUFTI SEKENDAR ALI MOLLA"; Total=165; Boarding=155},
        [PSCustomObject]@{Code="N24PGS 1378/26"; Name="MADRASA QURANIA ARABIA"; Estd=2019; Village="JOARIA"; POPS="PARULIA - DEGANGA"; PIN="743423"; Syllabus="HIFZ"; Expenditure=710000; Words="SEVEN LAKH TEN THOUSAND"; Head="HAQUE"; Total=79; Boarding=79},
        [PSCustomObject]@{Code="N24PGS 1379/26"; Name="MADRASA ANWARUL ULOOM MIRZANAGAR"; Estd=2022; Village="GOLABARI CHANDPUR"; POPS="GOURIBHOJE - MATIA"; PIN="743445"; Syllabus="HIFZ"; Expenditure=1499170; Words="FOURTEEN LAKH NINETY NINE THOUSAND ONE HUNDRED SEVENTY"; Head="MOULANA ALKAMA MAJAHIRI"; Total=100; Boarding=100},
        [PSCustomObject]@{Code="N24PGS 1385/26"; Name="JAMIA NUMANIA"; Estd=2020; Village="FOTURATI"; POPS="KOTHUR - BASIRHAT"; PIN="743422"; Syllabus="HIFZ TO ARBI SOM"; Expenditure=6601561; Words="SIXTY SIX LAKH ONE THOUSAND FIVE HUNDRED SIXTY ONE"; Head="MUFTI NASIRUDDIN QASIMI"; Total=173; Boarding=173},
        [PSCustomObject]@{Code="N24PGS 1386/26"; Name="JAMIATUL QURAN WAS SUNNAH"; Estd=2023; Village="NIMDARIA"; POPS="TAKI - BASIRHAT"; PIN="743429"; Syllabus="HIFZ TO ARBI CHAHARUM"; Expenditure=6356000; Words="SIXTY THREE LAKH FIFTY SIX THOUSAND"; Head="SHAH MOHAMMAD ABDULLAH"; Total=242; Boarding=170},
        [PSCustomObject]@{Code="N24PGS 1387/26"; Name="MADRASA BAITUL HIKMAH"; Estd=2024; Village="BHABANIPUR"; POPS="BHABANIPUR - HAROA"; PIN="743445"; Syllabus="HIFZ"; Expenditure=993996; Words="NINE LAKH NINETY THREE THOUSAND NINE HUNDRED NINETY SIX"; Head="MOULANA NURUL ISLAM"; Total=42; Boarding=41}
    )

    Write-Host "Writing Data to Database Sheet..."
    $headers = @(
        "CODE", "NAME OF MADRASAH", "ESTD.", "VILLAGE", "P.O - P.S", "PIN",
        "TALEEM CLASS/JAMAT UPTO", "EXPENDITURE IN ACADEMIC YEAR OF 2025-26 Rs.",
        "EXPENDITURE IN ACADEMIC YEAR OF 2025-26 IN WORDS Rs.",
        "HEAD MAULANA / SECRETARY", "TOTAL STUDENTS", "BOARDING STUDENTS"
    )
    for ($i = 0; $i -lt $headers.Count; $i++) {
        $wsDb.Cells.Item(1, $i + 1) = $headers[$i]
    }
    
    $row = 2
    foreach ($item in $data) {
        $wsDb.Cells.Item($row, 1) = $item.Code
        $wsDb.Cells.Item($row, 2) = $item.Name
        $wsDb.Cells.Item($row, 3) = $item.Estd
        $wsDb.Cells.Item($row, 4) = $item.Village
        $wsDb.Cells.Item($row, 5) = $item.POPS
        $wsDb.Cells.Item($row, 6) = $item.PIN
        $wsDb.Cells.Item($row, 7) = $item.Syllabus
        $wsDb.Cells.Item($row, 8) = $item.Expenditure
        $wsDb.Cells.Item($row, 9) = $item.Words
        $wsDb.Cells.Item($row, 10) = $item.Head
        $wsDb.Cells.Item($row, 11) = $item.Total
        $wsDb.Cells.Item($row, 12) = $item.Boarding
        $row++
    }
    $lastRow = $row - 1
    
    Write-Host "Formatting Database Sheet as Excel Table..."
    $tableRange = $wsDb.Range("A1:L$lastRow")
    $tbl = $wsDb.ListObjects.Add(1, $tableRange, $null, 1) # 1 = xlSrcRange, 1 = xlYes
    $tbl.Name = "tblDatabase"
    $tbl.TableStyle = "TableStyleMedium9" # Nice blue style
    
    # Add a District Column to the Table
    Write-Host "Adding District column to database..."
    $distCol = $tbl.ListColumns.Add()
    $distCol.Name = "DISTRICT"
    # Auto-populate formula
    $distCol.DataBodyRange.Formula = "=IF(LEFT([@CODE],6)=""S24PGS"",""South 24 Parganas"",""North 24 Parganas"")"
    
    # Format numbers in Table - using DataBodyRange to prevent header formatting issues
    $tbl.ListColumns.Item(8).DataBodyRange.NumberFormat = "#,##0"
    $tbl.ListColumns.Item(11).DataBodyRange.NumberFormat = "#,##0"
    $tbl.ListColumns.Item(12).DataBodyRange.NumberFormat = "#,##0"
    
    # Align Database table cells
    $wsDb.Range("A2:A$lastRow").HorizontalAlignment = -4108 # Center
    $wsDb.Range("C2:C$lastRow").HorizontalAlignment = -4108 # Center
    $wsDb.Range("F2:F$lastRow").HorizontalAlignment = -4108 # Center
    $wsDb.Range("M2:M$lastRow").HorizontalAlignment = -4108 # Center (District)

    Write-Host "Creating Forms Sheet..."
    # Title Block
    $wsForm.Cells.Item(1, 1) = "MADRASA DATA ENTRY FORM"
    $titleRange = $wsForm.Range("A1:L1")
    $titleRange.Merge()
    $titleRange.Font.Bold = $true
    $titleRange.Font.Size = 16
    $titleRange.Font.Color = Get-ExcelColor 255 255 255
    $titleRange.Interior.Color = Get-ExcelColor 30 115 70 # Dark Excel Green
    $titleRange.HorizontalAlignment = -4108
    $titleRange.VerticalAlignment = -4108
    $titleRange.RowHeight = 35
    
    # Instructions
    $wsForm.Cells.Item(2, 1) = "Instructions: Fill in the rows below to manually insert records. Fields with * are mandatory. Press the 'Submit to Database' button to append and refresh charts."
    $instRange = $wsForm.Range("A2:L2")
    $instRange.Merge()
    $instRange.Font.Italic = $true
    $instRange.Font.Size = 10
    $instRange.Font.Color = Get-ExcelColor 80 80 80
    $instRange.RowHeight = 25
    $instRange.VerticalAlignment = -4108

    # Set Row 3 height to accommodate the button cleanly
    $wsForm.Rows.Item(3).RowHeight = 26

    # Form headers at Row 4
    $formHeaders = @(
        "CODE*", "NAME OF MADRASAH*", "ESTD.", "VILLAGE", "P.O - P.S", "PIN*",
        "TALEEM CLASS/JAMAT UPTO", "EXPENDITURE (Rs.)*", "EXPENDITURE IN WORDS",
        "HEAD MAULANA / SECRETARY", "TOTAL STUDENTS*", "BOARDING STUDENTS*"
    )
    for ($i = 0; $i -lt $formHeaders.Count; $i++) {
        $wsForm.Cells.Item(4, $i + 1) = $formHeaders[$i]
    }
    
    $headerRange = $wsForm.Range("A4:L4")
    $headerRange.Font.Bold = $true
    $headerRange.Font.Size = 10
    $headerRange.Font.Color = Get-ExcelColor 255 255 255
    $headerRange.Interior.Color = Get-ExcelColor 46 139 87 # Medium Sea Green
    $headerRange.HorizontalAlignment = -4108
    $headerRange.VerticalAlignment = -4108
    $headerRange.RowHeight = 24

    # Setup blank manual input fields (Rows 5 to 14, 10 blank rows)
    $inputRange = $wsForm.Range("A5:L14")
    $inputRange.Borders.LineStyle = 1 # xlContinuous
    $inputRange.Borders.Weight = 2    # xlThin
    $inputRange.Borders.Color = Get-ExcelColor 200 200 200
    $inputRange.RowHeight = 20
    
    # Format manual entry cells - using safe formatting
    $wsForm.Range("H5:H14").NumberFormat = "#,##0"
    $wsForm.Range("K5:K14").NumberFormat = "#,##0"
    $wsForm.Range("L5:L14").NumberFormat = "#,##0"
    $wsForm.Range("A5:A14").HorizontalAlignment = -4108
    $wsForm.Range("C5:C14").HorizontalAlignment = -4108
    $wsForm.Range("F5:F14").HorizontalAlignment = -4108

    # Add Submit Button on Forms - placed dynamically on cell I3 so it is clearly visible
    Write-Host "Adding Submit button..."
    $btnCell = $wsForm.Cells.Item(3, 9) # Column I, Row 3
    $btn = $wsForm.Buttons().Add($btnCell.Left, $btnCell.Top, 180, 22)
    $btn.Caption = "Submit to Database"
    $btn.OnAction = "SubmitToDatabase"
    $btn.Font.Bold = $true
    $btn.Font.Size = 10

    Write-Host "Creating Dashboard Visualizations..."
    # Dashboard Sheet setup
    $wsDashboard.Cells.Item(1, 1) = "MADRASA PERFORMANCE & ANALYTICS DASHBOARD"
    $dashTitle = $wsDashboard.Range("A1:N1")
    $dashTitle.Merge()
    $dashTitle.Font.Bold = $true
    $dashTitle.Font.Size = 18
    $dashTitle.Font.Color = Get-ExcelColor 255 255 255
    $dashTitle.Interior.Color = Get-ExcelColor 41 128 185 # Slate Blue
    $dashTitle.HorizontalAlignment = -4108
    $dashTitle.VerticalAlignment = -4108
    $dashTitle.RowHeight = 35
    
    # Chart 1: Enrollment comparison
    Write-Host "Creating Student Enrollment Chart..."
    $chartObj1 = $wsDashboard.ChartObjects().Add(20, 50, 600, 320)
    $chart1 = $chartObj1.Chart
    $chart1.ChartType = 51 # xlColumnClustered
    
    $chartRangeString1 = "Database!B1:B$lastRow,Database!K1:L$lastRow"
    $chartRange1 = $wsDb.Range($chartRangeString1)
    $chart1.SetSourceData($chartRange1)
    $chart1.HasTitle = $true
    $chart1.ChartTitle.Text = "Student Enrollment (Total vs. Boarding)"

    # Chart 2: Expenditure
    Write-Host "Creating Expenditure Chart..."
    $chartObj2 = $wsDashboard.ChartObjects().Add(640, 50, 600, 320)
    $chart2 = $chartObj2.Chart
    $chart2.ChartType = 51 # xlColumnClustered
    
    $chartRangeString2 = "Database!B1:B$lastRow,Database!H1:H$lastRow"
    $chartRange2 = $wsDb.Range($chartRangeString2)
    $chart2.SetSourceData($chartRange2)
    $chart2.HasTitle = $true
    $chart2.ChartTitle.Text = "Annual Expenditure (Rs.)"
    $chart2.Axes(2).TickLabels.NumberFormat = "#,##0"

    # Add Slicers for District and Estd Year
    Write-Host "Adding Filters (Slicers)..."
    $slicerCaches = $wb.SlicerCaches
    
    try {
        $slicerCache1 = $slicerCaches.Add2($tbl, "DISTRICT")
        # Add(SlicerDestination, Level, Name, Caption, Top, Left, Width, Height)
        $slicer1 = $slicerCache1.Slicers.Add($wsDashboard, $null, "DISTRICT", "Filter by District", 390, 20, 180, 110)
        $slicer1.Caption = "District Filter"
    } catch {
        Write-Host "Slicer for DISTRICT failed: $_"
    }

    try {
        $slicerCache2 = $slicerCaches.Add2($tbl, "ESTD.")
        $slicer2 = $slicerCache2.Slicers.Add($wsDashboard, $null, "ESTD.", "Filter by Estd. Year", 390, 220, 180, 200)
        $slicer2.Caption = "Estd Year Filter"
    } catch {
        Write-Host "Slicer for ESTD. failed: $_"
    }

    # Injecting the VBA Automation code
    Write-Host "Injecting VBA macro code..."
    $project = $wb.VBProject
    $module = $project.VBComponents.Add(1) # vbext_ct_StdModule
    $module.Name = "SubmitModule"
    
    $vbaCode = @"
Sub SubmitToDatabase()
    Dim wsForm As Worksheet
    Dim wsDb As Worksheet
    Dim tbl As ListObject
    Dim newRow As ListRow
    Dim r As Long
    Dim recordCount As Long
    Dim c As Long
    Dim valCode As String
    Dim valName As String
    
    Set wsForm = ThisWorkbook.Sheets("Forms")
    Set wsDb = ThisWorkbook.Sheets("Database")
    
    On Error Resume Next
    Set tbl = wsDb.ListObjects("tblDatabase")
    On Error GoTo 0
    
    If tbl Is Nothing Then
        MsgBox "Critical Error: Database table 'tblDatabase' was not found.", vbCritical, "System Error"
        Exit Sub
    End If
    
    r = 5 ' Form data starts at row 5
    recordCount = 0
    
    ' Loop to process and validate rows
    Do While True
        valCode = Trim(wsForm.Cells(r, 1).Value)
        valName = Trim(wsForm.Cells(r, 2).Value)
        
        ' If both fields are empty, we stop
        If valCode = "" And valName = "" Then
            Exit Do
        End If
        
        ' Validation checks
        If valCode = "" Or valName = "" Then
            MsgBox "Error in Row " & (r - 4) & ": Both 'CODE' and 'NAME OF MADRASAH' are mandatory fields.", vbCritical, "Validation Error"
            Exit Sub
        End If
        
        ' Add to Database Table
        Set newRow = tbl.ListRows.Add
        For c = 1 To 12
            newRow.Range(c).Value = wsForm.Cells(r, c).Value
        Next c
        
        ' Clear data from Form sheet for this row
        wsForm.Range(wsForm.Cells(r, 1), wsForm.Cells(r, 12)).ClearContents
        
        r = r + 1
        recordCount = recordCount + 1
    Loop
    
    If recordCount > 0 Then
        ' Refresh all charts and pivots linked to the Table
        ThisWorkbook.RefreshAll
        MsgBox "Success! " & recordCount & " record(s) submitted to Database. Dashboard updated.", vbInformation, "Submission Complete"
    Else
        MsgBox "No records found. Please enter data in the Form starting at Row 5.", vbExclamation, "No Data"
    End If
End Sub

Sub CreateDashboardSlicers()
    Dim wsDb As Worksheet
    Dim wsDash As Worksheet
    Dim tbl As ListObject
    Dim sc As SlicerCache
    Dim sl As Slicer
    
    Set wsDb = ThisWorkbook.Sheets("Database")
    Set wsDash = ThisWorkbook.Sheets("Dashboard")
    
    On Error Resume Next
    Set tbl = wsDb.ListObjects("tblDatabase")
    
    ' Add District Slicer
    Set sc = ThisWorkbook.SlicerCaches.Add2(tbl, "DISTRICT")
    Set sl = sc.Slicers.Add(wsDash, , "DISTRICT", "District Filter", 390, 20, 180, 110)
    
    ' Add Estd Slicer
    Set sc = ThisWorkbook.SlicerCaches.Add2(tbl, "ESTD.")
    Set sl = sc.Slicers.Add(wsDash, , "ESTD.", "Estd Year Filter", 390, 220, 180, 200)
    On Error GoTo 0
End Sub
"@
    $module.CodeModule.AddFromString($vbaCode)
    Write-Host "VBA macro successfully injected."

    # Run the slicer creator macro native to Excel to bypass PowerShell COM casting issue
    Write-Host "Running native Excel Slicer creator..."
    $excel.Run("CreateDashboardSlicers")

    # Autofit columns
    Write-Host "Autofitting Columns..."
    $wsForm.Columns.AutoFit() | Out-Null
    $wsDb.Columns.AutoFit() | Out-Null
    
    # Adjust specific column widths for a premium look
    $wsForm.Columns.Item(2).ColumnWidth = 40 # Name of Madrasah
    $wsForm.Columns.Item(9).ColumnWidth = 50 # Expenditure in Words
    
    $wsDb.Columns.Item(2).ColumnWidth = 40 # Name
    $wsDb.Columns.Item(9).ColumnWidth = 50 # Expenditure in Words
    $wsDb.Columns.Item(13).ColumnWidth = 20 # District

    # Enable gridlines visually on all sheets
    Write-Host "Enabling Sheet Gridlines..."
    $wsForm.Activate()
    $excel.ActiveWindow.DisplayGridlines = $true
    $wsDb.Activate()
    $excel.ActiveWindow.DisplayGridlines = $true
    $wsDashboard.Activate()
    $excel.ActiveWindow.DisplayGridlines = $true
    
    # Set active sheet back to Forms for a good first-open experience
    $wsForm.Activate()

    # Save Workbook as macro-enabled xlsm
    Write-Host "Saving Workbook to $savePath..."
    $wb.SaveAs($savePath, 52) # 52 = xlOpenXMLWorkbookMacroEnabled
    $wb.Close($true)
    Write-Host "Workbook successfully generated!"
    
} catch {
    Write-Host "An error occurred during build: $_"
    if ($wb) { $wb.Close($false) }
} finally {
    # Restore AccessVBOM setting
    if ($originalValue -ne $null) {
        Set-ItemProperty -Path $registryPath -Name $regValueName -Value $originalValue -Type DWord -Force | Out-Null
    } else {
        Remove-ItemProperty -Path $registryPath -Name $regValueName -Force -ErrorAction SilentlyContinue | Out-Null
    }
    Write-Host "VBA Project Access setting restored in registry."
    
    # Quit Excel and release object COM
    if ($excel) {
        $excel.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    }
    Write-Host "Excel processes terminated."
}
