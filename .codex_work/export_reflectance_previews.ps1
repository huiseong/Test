$ErrorActionPreference = 'Stop'

$xlsxPath = 'C:\Users\lunarianmetal\Desktop\Test\outputs\reflectance-thickness-template-20261007\반사율_극값_두께_계산_템플릿.xlsx'
$previewDir = 'C:\Users\lunarianmetal\Desktop\Test\.codex_work\reflectance_excel_previews'
New-Item -ItemType Directory -Force -Path $previewDir | Out-Null

$excel = $null
$workbook = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.ScreenUpdating = $false
    $workbook = $excel.Workbooks.Open($xlsxPath, 0, $true)
    $excel.CalculateFullRebuild()

    $resultsForTest = $workbook.Worksheets.Item('결과')
    $nbForTest = $workbook.Worksheets.Item('Nb2O5 입력')
    $beforeTest = [double]$resultsForTest.Range('G12').Value2
    $originalTestInput = $nbForTest.Range('Z6').Value2
    $nbForTest.Range('Z6').Value2 = 1
    $excel.CalculateFullRebuild()
    $afterTest = [double]$resultsForTest.Range('G12').Value2
    $nbForTest.Range('Z6').Value2 = $originalTestInput
    $excel.CalculateFullRebuild()
    $restoredTest = [double]$resultsForTest.Range('G12').Value2
    if ([Math]::Abs($beforeTest - $afterTest) -lt 0.01 -or [Math]::Abs($beforeTest - $restoredTest) -gt 0.000001) {
        throw "Excel recalculation test failed: before=$beforeTest after=$afterTest restored=$restoredTest"
    }
    Write-Output "EXCEL_RECALC_TEST`tbefore=$beforeTest`tafter=$afterTest`trestored=$restoredTest"
    [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($resultsForTest) | Out-Null
    [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($nbForTest) | Out-Null

    $settingsForTest = $workbook.Worksheets.Item('설정')
    $resultsForTmmTest = $workbook.Worksheets.Item('결과')
    $beforeTmmTest = [double]$resultsForTmmTest.Range('N12').Value2
    $originalNbThickness = $settingsForTest.Range('C16').Value2
    $settingsForTest.Range('C16').Value2 = 8
    $excel.CalculateFullRebuild()
    $afterTmmTest = [double]$resultsForTmmTest.Range('N12').Value2
    $settingsForTest.Range('C16').Value2 = $originalNbThickness
    $excel.CalculateFullRebuild()
    $restoredTmmTest = [double]$resultsForTmmTest.Range('N12').Value2
    if ([Math]::Abs($beforeTmmTest - $afterTmmTest) -lt 0.1 -or [Math]::Abs($beforeTmmTest - $restoredTmmTest) -gt 0.000001) {
        throw "Excel TMM recalculation test failed: before=$beforeTmmTest after=$afterTmmTest restored=$restoredTmmTest"
    }
    Write-Output "EXCEL_TMM_RECALC_TEST`tbefore=$beforeTmmTest`tafter=$afterTmmTest`trestored=$restoredTmmTest"
    [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($settingsForTest) | Out-Null
    [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($resultsForTmmTest) | Out-Null

    $pageSpecs = @(
        @{ Sheet = '결과'; Area = '$A$1:$O$46'; Orientation = 2; Wide = 1; Tall = 2 },
        @{ Sheet = '설정'; Area = '$A$1:$G$38'; Orientation = 2; Wide = 1; Tall = 1 },
        @{ Sheet = 'Nb2O5 입력'; Area = '$A$1:$BD$20'; Orientation = 2; Wide = 1; Tall = 1 },
        @{ Sheet = '2층 입력'; Area = '$A$1:$BD$20'; Orientation = 2; Wide = 1; Tall = 1 },
        @{ Sheet = '굴절률'; Area = '$A$1:$D$46'; Orientation = 1; Wide = 1; Tall = 1 },
        @{ Sheet = 'TMM 계산표'; Area = '$A$1:$I$25'; Orientation = 2; Wide = 1; Tall = 1 }
    )

    foreach ($spec in $pageSpecs) {
        $worksheet = $workbook.Worksheets.Item($spec.Sheet)
        $worksheet.PageSetup.PrintArea = $spec.Area
        $worksheet.PageSetup.Orientation = $spec.Orientation
        $worksheet.PageSetup.Zoom = $false
        $worksheet.PageSetup.FitToPagesWide = $spec.Wide
        $worksheet.PageSetup.FitToPagesTall = $spec.Tall
        $worksheet.PageSetup.LeftMargin = $excel.InchesToPoints(0.25)
        $worksheet.PageSetup.RightMargin = $excel.InchesToPoints(0.25)
        $worksheet.PageSetup.TopMargin = $excel.InchesToPoints(0.35)
        $worksheet.PageSetup.BottomMargin = $excel.InchesToPoints(0.35)
        [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($worksheet) | Out-Null
    }

    $pdfTarget = Join-Path $previewDir 'workbook_preview.pdf'
    $workbook.ExportAsFixedFormat(0, $pdfTarget, 0, $true, $false)
    Write-Output "PDF`t$pdfTarget"

    $resultsSheet = $workbook.Worksheets.Item('결과')
    $thicknessChart = $resultsSheet.ChartObjects().Item(1).Chart
    $chartTarget = Join-Path $previewDir 'thickness_chart.png'
    if (-not $thicknessChart.Export($chartTarget, 'PNG')) {
        throw "Failed to export $chartTarget"
    }
    Write-Output "PREVIEW`t$chartTarget"
    [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($thicknessChart) | Out-Null
    [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($resultsSheet) | Out-Null
}
finally {
    if ($null -ne $workbook) {
        $workbook.Close($false)
        [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($workbook) | Out-Null
    }
    if ($null -ne $excel) {
        $excel.Quit()
        [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) | Out-Null
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
