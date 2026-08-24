param(
    [Parameter(Mandatory = $true)]
    [string]$InputDocx,

    [Parameter(Mandatory = $true)]
    [string]$OutputPdf,

    [string]$UpdatedDocx = ''
)

$ErrorActionPreference = 'Stop'

$inputPath = (Resolve-Path -LiteralPath $InputDocx).Path
$outputPath = [System.IO.Path]::GetFullPath($OutputPdf)
$outputDir = [System.IO.Path]::GetDirectoryName($outputPath)
$updatedPath = if ($UpdatedDocx) { [System.IO.Path]::GetFullPath($UpdatedDocx) } else { '' }

if (-not (Test-Path -LiteralPath $outputDir)) {
    New-Item -ItemType Directory -Path $outputDir | Out-Null
}

$word = $null
$document = $null

try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0

    Write-Output "STEP=OPEN_START"
    # OpenAndRepair helps Word recover harmless package-level inconsistencies
    # produced by third-party DOCX generators while preserving editable content.
    $document = $word.Documents.Open(
        $inputPath,
        $false,
        $false,
        $false,
        '',
        '',
        $false,
        '',
        '',
        0,
        0,
        $false,
        $true,
        0,
        $true
    )
    Write-Output "STEP=OPEN_DONE"

    foreach ($toc in $document.TablesOfContents) {
        Write-Output "STEP=TOC_UPDATE_START"
        $toc.Update() | Out-Null
        Write-Output "STEP=TOC_UPDATE_DONE"
    }

    Write-Output "STEP=FIELDS_UPDATE_START"
    $document.Fields.Update() | Out-Null
    Write-Output "STEP=FIELDS_UPDATE_DONE"
    Write-Output "STEP=REPAGINATE_START"
    $document.Repaginate()
    Write-Output "STEP=REPAGINATE_DONE"
    $pageCount = $document.ComputeStatistics(2)
    if ($updatedPath) {
        Write-Output "STEP=SAVE_AS_START"
        $document.SaveAs2($updatedPath, 16)
        Write-Output "STEP=SAVE_AS_DONE"
    }
    Write-Output "STEP=EXPORT_START"
    $document.ExportAsFixedFormat($outputPath, 17)
    Write-Output "STEP=EXPORT_DONE"

    Write-Output "PAGE_COUNT=$pageCount"
    Write-Output "PDF_PATH=$outputPath"
}
finally {
    if ($null -ne $document) {
        $document.Close($false)
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($document)
    }
    if ($null -ne $word) {
        $word.Quit()
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($word)
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
