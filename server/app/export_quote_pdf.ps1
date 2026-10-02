param(
    [Parameter(Mandatory=$true)][string]$WorkbookPath,
    [Parameter(Mandatory=$true)][string]$PdfPath
)
$ErrorActionPreference = 'Stop'
$excelApp = $null
$quoteWorkbook = $null
$quoteSheet = $null
try {
    $excelApp = New-Object -ComObject Excel.Application
    $excelApp.Visible = $false
    $excelApp.DisplayAlerts = $false
    $excelApp.AskToUpdateLinks = $false
    $excelApp.AutomationSecurity = 3
    $quoteWorkbook = $excelApp.Workbooks.Open($WorkbookPath, 0, $true)
    $quoteSheet = $quoteWorkbook.Worksheets.Item(1)
    $quoteSheet.PageSetup.Zoom = $false
    $quoteSheet.PageSetup.FitToPagesWide = 1
    $quoteSheet.PageSetup.FitToPagesTall = 1
    $quoteSheet.ExportAsFixedFormat(0, $PdfPath)
} finally {
    if ($null -ne $quoteWorkbook) { $quoteWorkbook.Close($false) }
    if ($null -ne $excelApp) { $excelApp.Quit() }
    foreach ($comObject in @($quoteSheet, $quoteWorkbook, $excelApp)) {
        if ($null -ne $comObject) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($comObject) }
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
