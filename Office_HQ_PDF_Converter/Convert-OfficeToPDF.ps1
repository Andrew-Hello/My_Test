[CmdletBinding()]
param(
    [Parameter(Position = 0, ValueFromRemainingArguments = $true)]
    [string[]]$InputFiles
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# -----------------------------------------------------------------------------
# High-quality Office -> PDF converter
# Requires desktop Microsoft Office on Windows.
# Supported: .docx .doc .xlsx .xls .pptx .ppt
# The source document is opened read-only and is never saved/modified.
# -----------------------------------------------------------------------------

$SupportedExtensions = @('.docx', '.doc', '.xlsx', '.xls', '.pptx', '.ppt')

function Write-Info([string]$Message) {
    Write-Host "[INFO] $Message" -ForegroundColor Cyan
}

function Write-Ok([string]$Message) {
    Write-Host "[ OK ] $Message" -ForegroundColor Green
}

function Write-WarnText([string]$Message) {
    Write-Host "[WARN] $Message" -ForegroundColor Yellow
}

function Write-Fail([string]$Message) {
    Write-Host "[FAIL] $Message" -ForegroundColor Red
}

function Release-ComObject($Object) {
    if ($null -ne $Object) {
        try {
            if ([System.Runtime.InteropServices.Marshal]::IsComObject($Object)) {
                [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($Object)
            }
        } catch {
            # Ignore cleanup-only failures.
        }
    }
}

function Force-ComCleanup {
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

function Get-OutputPdfPath([string]$SourcePath) {
    $directory = [System.IO.Path]::GetDirectoryName($SourcePath)
    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($SourcePath)

    $candidate = Join-Path $directory ($baseName + '.pdf')
    if (-not (Test-Path -LiteralPath $candidate)) {
        return $candidate
    }

    $candidate = Join-Path $directory ($baseName + '_HQ.pdf')
    if (-not (Test-Path -LiteralPath $candidate)) {
        return $candidate
    }

    $index = 2
    while ($true) {
        $candidate = Join-Path $directory ("{0}_HQ_{1}.pdf" -f $baseName, $index)
        if (-not (Test-Path -LiteralPath $candidate)) {
            return $candidate
        }
        $index++
    }
}

function Disable-OfficeMacros($Application) {
    # msoAutomationSecurityForceDisable = 3
    try { $Application.AutomationSecurity = 3 } catch { }
}

function Convert-WordToPdf([string]$SourcePath, [string]$OutputPath) {
    $word = $null
    $doc = $null

    try {
        Write-Info 'Starting Microsoft Word...'
        $word = New-Object -ComObject Word.Application
        $word.Visible = $false
        $word.DisplayAlerts = 0
        Disable-OfficeMacros $word

        # FileName, ConfirmConversions=False, ReadOnly=True, AddToRecentFiles=False
        $doc = $word.Documents.Open($SourcePath, $false, $true, $false)
        try { $doc.Repaginate() } catch { }

        # Constants:
        # wdExportFormatPDF = 17
        # wdExportOptimizeForPrint = 0
        # wdExportAllDocument = 0
        # wdExportDocumentContent = 0
        # wdExportCreateHeadingBookmarks = 1
        # BitmapMissingFonts = True preserves appearance when fonts cannot be embedded.
        # UseISO19005_1 = False avoids PDF/A restrictions that can introduce artifacts.
        # OptimizeForImageQuality = True explicitly requests original image quality.

        $exported = $false

        try {
            # Newest Word versions: includes OptimizeForImageQuality and improved tagging.
            $doc.ExportAsFixedFormat3(
                $OutputPath, 17, $false, 0, 0, 1, 1, 0,
                $true, $false, 1, $true, $true, $false,
                $true, $false, $null
            )
            $exported = $true
            Write-Info 'Word ExportAsFixedFormat3 used (OptimizeForImageQuality=True).'
        } catch {
            Write-WarnText 'ExportAsFixedFormat3 is unavailable; trying ExportAsFixedFormat2.'
        }

        if (-not $exported) {
            try {
                $doc.ExportAsFixedFormat2(
                    $OutputPath, 17, $false, 0, 0, 1, 1, 0,
                    $true, $false, 1, $true, $true, $false,
                    $true, $null
                )
                $exported = $true
                Write-Info 'Word ExportAsFixedFormat2 used (OptimizeForImageQuality=True).'
            } catch {
                Write-WarnText 'ExportAsFixedFormat2 is unavailable; using classic ExportAsFixedFormat.'
            }
        }

        if (-not $exported) {
            $doc.ExportAsFixedFormat(
                $OutputPath, 17, $false, 0, 0, 1, 1, 0,
                $true, $false, 1, $true, $true, $false
            )
            Write-Info 'Classic Word ExportAsFixedFormat used at print quality.'
        }
    }
    finally {
        if ($null -ne $doc) {
            try { $doc.Close(0) } catch { }
        }
        Release-ComObject $doc

        if ($null -ne $word) {
            try { $word.Quit() } catch { }
        }
        Release-ComObject $word
        Force-ComCleanup
    }
}

function Convert-ExcelToPdf([string]$SourcePath, [string]$OutputPath) {
    $excel = $null
    $workbook = $null

    try {
        Write-Info 'Starting Microsoft Excel...'
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        try { $excel.AskToUpdateLinks = $false } catch { }
        Disable-OfficeMacros $excel

        # FileName, UpdateLinks=0, ReadOnly=True
        $workbook = $excel.Workbooks.Open($SourcePath, 0, $true)

        # Constants:
        # xlTypePDF = 0
        # xlQualityStandard = 0  (highest Excel fixed-format quality option)
        # IgnorePrintAreas = True so content/images outside a saved print area are not lost.
        $workbook.ExportAsFixedFormat(0, $OutputPath, 0, $true, $true)
        Write-Info 'Excel exported at xlQualityStandard with saved print areas ignored.'
    }
    finally {
        if ($null -ne $workbook) {
            try { $workbook.Close($false) } catch { }
        }
        Release-ComObject $workbook

        if ($null -ne $excel) {
            try { $excel.Quit() } catch { }
        }
        Release-ComObject $excel
        Force-ComCleanup
    }
}

function Convert-PowerPointToPdf([string]$SourcePath, [string]$OutputPath) {
    $powerPoint = $null
    $presentation = $null

    try {
        Write-Info 'Starting Microsoft PowerPoint...'
        $powerPoint = New-Object -ComObject PowerPoint.Application
        Disable-OfficeMacros $powerPoint

        # FileName, ReadOnly=True, Untitled=False, WithWindow=False
        # MsoTriState: True=-1, False=0
        $presentation = $powerPoint.Presentations.Open($SourcePath, -1, 0, 0)

        # Constants:
        # ppFixedFormatTypePDF = 2
        # ppFixedFormatIntentPrint = 2
        # FrameSlides = False
        # ppPrintHandoutVerticalFirst = 1 (ignored for slide output)
        # ppPrintOutputSlides = 1
        # PrintHiddenSlides = False
        # PrintRange = Nothing
        # ppPrintAll = 1
        # IncludeDocProperties = True
        # KeepIRMSettings = True
        # DocStructureTags = True
        # BitmapMissingFonts = True
        # UseISO19005_1 = False
        try {
            $presentation.ExportAsFixedFormat(
                $OutputPath, 2, 2, 0, 1, 1, 0,
                $null, 1, '', $true, $true, $true, $true, $false
            )
            Write-Info 'PowerPoint ExportAsFixedFormat used with Print intent.'
        } catch {
            Write-WarnText 'PowerPoint ExportAsFixedFormat failed; falling back to SaveAs PDF.'
            # ppSaveAsPDF = 32
            $presentation.SaveAs($OutputPath, 32)
        }
    }
    finally {
        if ($null -ne $presentation) {
            try { $presentation.Close() } catch { }
        }
        Release-ComObject $presentation

        if ($null -ne $powerPoint) {
            try { $powerPoint.Quit() } catch { }
        }
        Release-ComObject $powerPoint
        Force-ComCleanup
    }
}

function Convert-OneFile([string]$InputPath) {
    $resolved = (Resolve-Path -LiteralPath $InputPath).Path

    if ((Get-Item -LiteralPath $resolved).PSIsContainer) {
        throw "Folders are not supported: $resolved"
    }

    $extension = [System.IO.Path]::GetExtension($resolved).ToLowerInvariant()
    if ($SupportedExtensions -notcontains $extension) {
        throw "Unsupported file type '$extension'. Supported: $($SupportedExtensions -join ', ')"
    }

    $outputPath = Get-OutputPdfPath $resolved

    Write-Host ''
    Write-Host ('=' * 78) -ForegroundColor DarkGray
    Write-Info ("Source : {0}" -f $resolved)
    Write-Info ("Output : {0}" -f $outputPath)

    switch ($extension) {
        '.docx' { Convert-WordToPdf $resolved $outputPath }
        '.doc'  { Convert-WordToPdf $resolved $outputPath }
        '.xlsx' { Convert-ExcelToPdf $resolved $outputPath }
        '.xls'  { Convert-ExcelToPdf $resolved $outputPath }
        '.pptx' { Convert-PowerPointToPdf $resolved $outputPath }
        '.ppt'  { Convert-PowerPointToPdf $resolved $outputPath }
    }

    if (-not (Test-Path -LiteralPath $outputPath)) {
        throw 'Office returned without an error, but the PDF file was not created.'
    }

    $pdf = Get-Item -LiteralPath $outputPath
    if ($pdf.Length -le 0) {
        throw 'The generated PDF is empty.'
    }

    Write-Ok ("Created: {0} ({1:N2} MB)" -f $pdf.FullName, ($pdf.Length / 1MB))
}

if ($null -eq $InputFiles -or $InputFiles.Count -eq 0) {
    Write-Host ''
    Write-Host 'High Quality Office -> PDF Converter' -ForegroundColor White
    Write-Host ''
    Write-Host 'Drag one or more Office files onto DragDrop_Office_to_PDF.cmd.'
    Write-Host 'Supported: DOCX, DOC, XLSX, XLS, PPTX, PPT'
    Write-Host ''
    exit 1
}

$successCount = 0
$failureCount = 0

foreach ($file in $InputFiles) {
    try {
        Convert-OneFile $file
        $successCount++
    } catch {
        $failureCount++
        Write-Fail ("{0}: {1}" -f $file, $_.Exception.Message)
    }
}

Write-Host ''
Write-Host ('=' * 78) -ForegroundColor DarkGray
Write-Host ("Finished. Success: {0}   Failed: {1}" -f $successCount, $failureCount) -ForegroundColor White

if ($failureCount -gt 0) {
    exit 2
}
exit 0
