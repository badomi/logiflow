param(
    [string]$WorkbookPath = 'C:\dev\logiflow-main\output\sample-quote\샘플견적서_상하이-부산_HMM_20GP.xlsx'
)
$ErrorActionPreference = 'Stop'
$resolvedWorkbook = (Resolve-Path -LiteralPath $WorkbookPath).Path
$pdfPath = [System.IO.Path]::ChangeExtension($resolvedWorkbook, '.pdf')
$excelApp = $null
$quoteWorkbook = $null
try {
    $excelApp = New-Object -ComObject Excel.Application
    $excelApp.Visible = $false
    $excelApp.DisplayAlerts = $false
    $excelApp.AskToUpdateLinks = $false
    $excelApp.AutomationSecurity = 3
    $quoteWorkbook = $excelApp.Workbooks.Open($resolvedWorkbook, 0, $false)
    $excelApp.CalculateFull()
    $quoteSheet = $quoteWorkbook.Worksheets.Item(1)
    $quoteSheet.PageSetup.Zoom = $false
    $quoteSheet.PageSetup.FitToPagesWide = 1
    $quoteSheet.PageSetup.FitToPagesTall = 1
    $quoteWorkbook.Save()
    $quoteSheet.ExportAsFixedFormat(0, $pdfPath)
    $proofSheet = $quoteWorkbook.Worksheets.Item(2)
    $proofSheet.ExportAsFixedFormat(0, 'C:\dev\logiflow-main\output\quote-preview\proof.pdf')
    Write-Output $pdfPath
    Write-Output ('TOTAL USD: ' + $quoteSheet.Range('C35').Value2)
} finally {
    if ($null -ne $quoteWorkbook) { $quoteWorkbook.Close($false) }
    if ($null -ne $excelApp) { $excelApp.Quit() }
    if ($null -ne $quoteWorkbook) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($quoteWorkbook) }
    if ($null -ne $excelApp) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excelApp) }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
