[CmdletBinding()]
param(
    [Parameter(Position = 0, ValueFromRemainingArguments = $true)]
    [string[]]$InputFiles
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# -----------------------------------------------------------------------------
# High-quality Office -> PDF converter + diagnostics
# Requires desktop Microsoft Office on Windows.
# Supported: .docx .doc .xlsx .xls .pptx .ppt
# The source document is opened read-only and is never saved/modified.
# -----------------------------------------------------------------------------

$SupportedExtensions = @('.docx', '.doc', '.xlsx', '.xls', '.pptx', '.ppt')
$ScriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$DiagnosticsDirectory = Join-Path $ScriptRoot 'diagnostics'
New-Item -ItemType Directory -Path $DiagnosticsDirectory -Force | Out-Null

function Write-Info([string]$Message) { Write-Host "[INFO] $Message" -ForegroundColor Cyan }
function Write-Ok([string]$Message) { Write-Host "[ OK ] $Message" -ForegroundColor Green }
function Write-WarnText([string]$Message) { Write-Host "[WARN] $Message" -ForegroundColor Yellow }
function Write-Fail([string]$Message) { Write-Host "[FAIL] $Message" -ForegroundColor Red }

function Release-ComObject($Object) {
    if ($null -ne $Object) {
        try {
            if ([System.Runtime.InteropServices.Marshal]::IsComObject($Object)) {
                [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($Object)
            }
        } catch { }
    }
}

function Force-ComCleanup {
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

function Get-SafeFileName([string]$Name) {
    if ([string]::IsNullOrWhiteSpace($Name)) { return 'unknown' }
    $invalid = [System.IO.Path]::GetInvalidFileNameChars()
    $safe = $Name
    foreach ($char in $invalid) { $safe = $safe.Replace([string]$char, '_') }
    return $safe
}

function Get-SystemDiagnostics {
    $osCaption = $null
    $osVersion = [Environment]::OSVersion.VersionString
    $osBuild = $null
    try {
        $os = Get-CimInstance Win32_OperatingSystem -ErrorAction Stop
        $osCaption = $os.Caption
        $osVersion = $os.Version
        $osBuild = $os.BuildNumber
    } catch { }

    return [ordered]@{
        os_caption = $osCaption
        os_version = $osVersion
        os_build = $osBuild
        process_architecture = [System.Runtime.InteropServices.RuntimeInformation]::ProcessArchitecture.ToString()
        powershell_version = $PSVersionTable.PSVersion.ToString()
        dotnet_runtime = [System.Runtime.InteropServices.RuntimeInformation]::FrameworkDescription
    }
}

function Get-OpenXmlMediaStats([string]$Path) {
    $ext = [System.IO.Path]::GetExtension($Path).ToLowerInvariant()
    if ($ext -notin @('.docx', '.xlsx', '.pptx')) { return $null }

    try {
        Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction SilentlyContinue
        $zip = [System.IO.Compression.ZipFile]::OpenRead($Path)
        try {
            $count = 0
            [int64]$totalBytes = 0
            $types = @{}
            foreach ($entry in $zip.Entries) {
                if ($entry.FullName -match '/media/') {
                    $count++
                    $totalBytes += $entry.Length
                    $mediaExt = [System.IO.Path]::GetExtension($entry.FullName).ToLowerInvariant()
                    if ([string]::IsNullOrWhiteSpace($mediaExt)) { $mediaExt = '(none)' }
                    if (-not $types.ContainsKey($mediaExt)) { $types[$mediaExt] = 0 }
                    $types[$mediaExt]++
                }
            }
            return [ordered]@{
                embedded_media_count = $count
                embedded_media_uncompressed_bytes = $totalBytes
                embedded_media_uncompressed_mb = [math]::Round($totalBytes / 1MB, 3)
                media_extensions = $types
            }
        }
        finally {
            if ($null -ne $zip) { $zip.Dispose() }
        }
    }
    catch {
        return [ordered]@{
            embedded_media_scan_error = $_.Exception.Message
        }
    }
}

function Get-PdfPageCountHeuristic([string]$PdfPath) {
    try {
        $file = Get-Item -LiteralPath $PdfPath
        if ($file.Length -gt 200MB) {
            return [ordered]@{ count = $null; method = 'skipped_over_200MB' }
        }
        $bytes = [System.IO.File]::ReadAllBytes($PdfPath)
        $text = [System.Text.Encoding]::ASCII.GetString($bytes)
        $matches = [regex]::Matches($text, '/Type\s*/Page(?!s)\b')
        return [ordered]@{ count = $matches.Count; method = 'pdf_dictionary_heuristic' }
    }
    catch {
        return [ordered]@{ count = $null; method = 'unavailable'; error = $_.Exception.Message }
    }
}

function Get-OutputPdfPath([string]$SourcePath) {
    $directory = [System.IO.Path]::GetDirectoryName($SourcePath)
    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($SourcePath)

    $candidate = Join-Path $directory ($baseName + '.pdf')
    if (-not (Test-Path -LiteralPath $candidate)) { return $candidate }

    $candidate = Join-Path $directory ($baseName + '_HQ.pdf')
    if (-not (Test-Path -LiteralPath $candidate)) { return $candidate }

    $index = 2
    while ($true) {
        $candidate = Join-Path $directory ("{0}_HQ_{1}.pdf" -f $baseName, $index)
        if (-not (Test-Path -LiteralPath $candidate)) { return $candidate }
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
    $meta = [ordered]@{
        application = 'Microsoft Word'
        office_version = $null
        office_build = $null
        source_pages = $null
        inline_shapes = $null
        floating_shapes = $null
        tables = $null
        export_engine = $null
        optimize_for_image_quality = $false
    }

    try {
        Write-Info 'Starting Microsoft Word...'
        $word = New-Object -ComObject Word.Application
        $word.Visible = $false
        $word.DisplayAlerts = 0
        Disable-OfficeMacros $word
        try { $meta.office_version = [string]$word.Version } catch { }
        try { $meta.office_build = [string]$word.Build } catch { }

        $doc = $word.Documents.Open($SourcePath, $false, $true, $false)
        try { [void]$doc.Repaginate() } catch { }
        try { $meta.source_pages = [int]$doc.ComputeStatistics(2) } catch { }
        try { $meta.inline_shapes = [int]$doc.InlineShapes.Count } catch { }
        try { $meta.floating_shapes = [int]$doc.Shapes.Count } catch { }
        try { $meta.tables = [int]$doc.Tables.Count } catch { }

        $exported = $false
        try {
            [void]$doc.ExportAsFixedFormat3(
                $OutputPath, 17, $false, 0, 0, 1, 1, 0,
                $true, $false, 1, $true, $true, $false,
                $true, $false, $null
            )
            $exported = $true
            $meta.export_engine = 'ExportAsFixedFormat3'
            $meta.optimize_for_image_quality = $true
            Write-Info 'Word ExportAsFixedFormat3 used (OptimizeForImageQuality=True).'
        } catch {
            Write-WarnText 'ExportAsFixedFormat3 is unavailable; trying ExportAsFixedFormat2.'
        }

        if (-not $exported) {
            try {
                [void]$doc.ExportAsFixedFormat2(
                    $OutputPath, 17, $false, 0, 0, 1, 1, 0,
                    $true, $false, 1, $true, $true, $false,
                    $true, $null
                )
                $exported = $true
                $meta.export_engine = 'ExportAsFixedFormat2'
                $meta.optimize_for_image_quality = $true
                Write-Info 'Word ExportAsFixedFormat2 used (OptimizeForImageQuality=True).'
            } catch {
                Write-WarnText 'ExportAsFixedFormat2 is unavailable; using classic ExportAsFixedFormat.'
            }
        }

        if (-not $exported) {
            [void]$doc.ExportAsFixedFormat(
                $OutputPath, 17, $false, 0, 0, 1, 1, 0,
                $true, $false, 1, $true, $true, $false
            )
            $meta.export_engine = 'ExportAsFixedFormat'
            Write-Info 'Classic Word ExportAsFixedFormat used at print quality.'
        }

        return $meta
    }
    finally {
        if ($null -ne $doc) { try { $doc.Close(0) } catch { } }
        Release-ComObject $doc
        if ($null -ne $word) { try { $word.Quit() } catch { } }
        Release-ComObject $word
        Force-ComCleanup
    }
}

function Convert-ExcelToPdf([string]$SourcePath, [string]$OutputPath) {
    $excel = $null
    $workbook = $null
    $meta = [ordered]@{
        application = 'Microsoft Excel'
        office_version = $null
        office_build = $null
        worksheet_count = $null
        total_shapes = 0
        picture_shapes = 0
        chart_objects = 0
        sheets_with_saved_print_area = 0
        export_engine = 'Workbook.ExportAsFixedFormat'
        export_quality = 'xlQualityStandard'
        ignore_saved_print_areas = $true
    }

    try {
        Write-Info 'Starting Microsoft Excel...'
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        try { $excel.AskToUpdateLinks = $false } catch { }
        Disable-OfficeMacros $excel
        try { $meta.office_version = [string]$excel.Version } catch { }
        try { $meta.office_build = [string]$excel.Build } catch { }

        $workbook = $excel.Workbooks.Open($SourcePath, 0, $true)
        $worksheetCount = [int]$workbook.Worksheets.Count
        $meta.worksheet_count = $worksheetCount

        for ($i = 1; $i -le $worksheetCount; $i++) {
            $sheet = $null
            try {
                $sheet = $workbook.Worksheets.Item($i)
                try { $meta.total_shapes += [int]$sheet.Shapes.Count } catch { }
                try { $meta.chart_objects += [int]$sheet.ChartObjects().Count } catch { }
                try {
                    if (-not [string]::IsNullOrWhiteSpace([string]$sheet.PageSetup.PrintArea)) {
                        $meta.sheets_with_saved_print_area++
                    }
                } catch { }
                try {
                    $shapeCount = [int]$sheet.Shapes.Count
                    for ($s = 1; $s -le $shapeCount; $s++) {
                        $shape = $null
                        try {
                            $shape = $sheet.Shapes.Item($s)
                            # msoLinkedPicture=11, msoPicture=13
                            if ([int]$shape.Type -in @(11, 13)) { $meta.picture_shapes++ }
                        } catch { }
                        finally { Release-ComObject $shape }
                    }
                } catch { }
            }
            finally { Release-ComObject $sheet }
        }

        [void]$workbook.ExportAsFixedFormat(0, $OutputPath, 0, $true, $true)
        Write-Info 'Excel exported at xlQualityStandard with saved print areas ignored.'
        return $meta
    }
    finally {
        if ($null -ne $workbook) { try { $workbook.Close($false) } catch { } }
        Release-ComObject $workbook
        if ($null -ne $excel) { try { $excel.Quit() } catch { } }
        Release-ComObject $excel
        Force-ComCleanup
    }
}

function Convert-PowerPointToPdf([string]$SourcePath, [string]$OutputPath) {
    $powerPoint = $null
    $presentation = $null
    $meta = [ordered]@{
        application = 'Microsoft PowerPoint'
        office_version = $null
        office_build = $null
        slide_count = $null
        total_shapes = 0
        picture_shapes = 0
        chart_shapes = 0
        slide_width_points = $null
        slide_height_points = $null
        export_engine = $null
        export_intent = 'Print'
    }

    try {
        Write-Info 'Starting Microsoft PowerPoint...'
        $powerPoint = New-Object -ComObject PowerPoint.Application
        Disable-OfficeMacros $powerPoint
        try { $meta.office_version = [string]$powerPoint.Version } catch { }
        try { $meta.office_build = [string]$powerPoint.Build } catch { }

        $presentation = $powerPoint.Presentations.Open($SourcePath, -1, 0, 0)
        $slideCount = [int]$presentation.Slides.Count
        $meta.slide_count = $slideCount
        try { $meta.slide_width_points = [double]$presentation.PageSetup.SlideWidth } catch { }
        try { $meta.slide_height_points = [double]$presentation.PageSetup.SlideHeight } catch { }

        for ($i = 1; $i -le $slideCount; $i++) {
            $slide = $null
            try {
                $slide = $presentation.Slides.Item($i)
                $shapeCount = [int]$slide.Shapes.Count
                $meta.total_shapes += $shapeCount
                for ($s = 1; $s -le $shapeCount; $s++) {
                    $shape = $null
                    try {
                        $shape = $slide.Shapes.Item($s)
                        $shapeType = [int]$shape.Type
                        if ($shapeType -in @(11, 13)) { $meta.picture_shapes++ }
                        try { if ($shape.HasChart -eq -1) { $meta.chart_shapes++ } } catch { }
                    } catch { }
                    finally { Release-ComObject $shape }
                }
            }
            finally { Release-ComObject $slide }
        }

        try {
            [void]$presentation.ExportAsFixedFormat(
                $OutputPath, 2, 2, 0, 1, 1, 0,
                $null, 1, '', $true, $true, $true, $true, $false
            )
            $meta.export_engine = 'ExportAsFixedFormat'
            Write-Info 'PowerPoint ExportAsFixedFormat used with Print intent.'
        } catch {
            Write-WarnText 'PowerPoint ExportAsFixedFormat failed; falling back to SaveAs PDF.'
            [void]$presentation.SaveAs($OutputPath, 32)
            $meta.export_engine = 'SaveAs PDF fallback'
        }

        return $meta
    }
    finally {
        if ($null -ne $presentation) { try { $presentation.Close() } catch { } }
        Release-ComObject $presentation
        if ($null -ne $powerPoint) { try { $powerPoint.Quit() } catch { } }
        Release-ComObject $powerPoint
        Force-ComCleanup
    }
}

function Save-DiagnosticRecord($Record, [string]$BaseName) {
    $stamp = Get-Date -Format 'yyyyMMdd_HHmmss_fff'
    $safe = Get-SafeFileName $BaseName
    $path = Join-Path $DiagnosticsDirectory ("{0}_{1}.json" -f $stamp, $safe)
    $Record | ConvertTo-Json -Depth 12 | Set-Content -LiteralPath $path -Encoding UTF8
    return $path
}

function Convert-OneFile([string]$InputPath) {
    $started = Get-Date
    $resolved = (Resolve-Path -LiteralPath $InputPath).Path
    $sourceItem = Get-Item -LiteralPath $resolved

    if ($sourceItem.PSIsContainer) { throw "Folders are not supported: $resolved" }

    $extension = [System.IO.Path]::GetExtension($resolved).ToLowerInvariant()
    if ($SupportedExtensions -notcontains $extension) {
        throw "Unsupported file type '$extension'. Supported: $($SupportedExtensions -join ', ')"
    }

    $outputPath = Get-OutputPdfPath $resolved
    $sourceHash = $null
    try { $sourceHash = (Get-FileHash -LiteralPath $resolved -Algorithm SHA256).Hash } catch { }

    $record = [ordered]@{
        schema_version = 1
        timestamp_local = (Get-Date).ToString('yyyy-MM-ddTHH:mm:ss.fffK')
        status = 'running'
        source = [ordered]@{
            file_name = $sourceItem.Name
            extension = $extension
            size_bytes = [int64]$sourceItem.Length
            size_mb = [math]::Round($sourceItem.Length / 1MB, 3)
            sha256 = $sourceHash
            full_local_path_recorded = $false
            openxml_media = Get-OpenXmlMediaStats $resolved
        }
        output = [ordered]@{
            file_name = [System.IO.Path]::GetFileName($outputPath)
            size_bytes = $null
            size_mb = $null
            sha256 = $null
            pdf_page_count = $null
        }
        environment = Get-SystemDiagnostics
        office = $null
        conversion = [ordered]@{
            started_local = $started.ToString('yyyy-MM-ddTHH:mm:ss.fffK')
            finished_local = $null
            duration_seconds = $null
            source_opened_read_only = $true
            macro_automation_security = 'ForceDisable when Office COM supports it'
            success = $false
            error = $null
        }
    }

    Write-Host ''
    Write-Host ('=' * 78) -ForegroundColor DarkGray
    Write-Info ("Source : {0}" -f $resolved)
    Write-Info ("Output : {0}" -f $outputPath)

    try {
        switch ($extension) {
            '.docx' { $record.office = Convert-WordToPdf $resolved $outputPath }
            '.doc'  { $record.office = Convert-WordToPdf $resolved $outputPath }
            '.xlsx' { $record.office = Convert-ExcelToPdf $resolved $outputPath }
            '.xls'  { $record.office = Convert-ExcelToPdf $resolved $outputPath }
            '.pptx' { $record.office = Convert-PowerPointToPdf $resolved $outputPath }
            '.ppt'  { $record.office = Convert-PowerPointToPdf $resolved $outputPath }
        }

        if (-not (Test-Path -LiteralPath $outputPath)) {
            throw 'Office returned without an error, but the PDF file was not created.'
        }

        $pdf = Get-Item -LiteralPath $outputPath
        if ($pdf.Length -le 0) { throw 'The generated PDF is empty.' }

        $record.output.size_bytes = [int64]$pdf.Length
        $record.output.size_mb = [math]::Round($pdf.Length / 1MB, 3)
        try { $record.output.sha256 = (Get-FileHash -LiteralPath $outputPath -Algorithm SHA256).Hash } catch { }
        $record.output.pdf_page_count = Get-PdfPageCountHeuristic $outputPath

        $finished = Get-Date
        $record.status = 'success'
        $record.conversion.success = $true
        $record.conversion.finished_local = $finished.ToString('yyyy-MM-ddTHH:mm:ss.fffK')
        $record.conversion.duration_seconds = [math]::Round(($finished - $started).TotalSeconds, 3)

        $diagPath = Save-DiagnosticRecord $record $sourceItem.BaseName
        Write-Ok ("Created: {0} ({1:N2} MB)" -f $pdf.FullName, ($pdf.Length / 1MB))
        Write-Info ("Diagnostic: {0}" -f $diagPath)
        return $record
    }
    catch {
        $finished = Get-Date
        $record.status = 'failure'
        $record.conversion.success = $false
        $record.conversion.error = $_.Exception.Message
        $record.conversion.finished_local = $finished.ToString('yyyy-MM-ddTHH:mm:ss.fffK')
        $record.conversion.duration_seconds = [math]::Round(($finished - $started).TotalSeconds, 3)
        $diagPath = Save-DiagnosticRecord $record $sourceItem.BaseName
        Write-Fail ("{0}: {1}" -f $sourceItem.Name, $_.Exception.Message)
        Write-Info ("Failure diagnostic: {0}" -f $diagPath)
        throw
    }
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
$batchRecords = @()
$batchStarted = Get-Date

foreach ($file in $InputFiles) {
    try {
        $result = Convert-OneFile $file
        $batchRecords += $result
        $successCount++
    }
    catch {
        $failureCount++
    }
}

$batchFinished = Get-Date
$batchSummary = [ordered]@{
    schema_version = 1
    type = 'batch_summary'
    started_local = $batchStarted.ToString('yyyy-MM-ddTHH:mm:ss.fffK')
    finished_local = $batchFinished.ToString('yyyy-MM-ddTHH:mm:ss.fffK')
    duration_seconds = [math]::Round(($batchFinished - $batchStarted).TotalSeconds, 3)
    success_count = $successCount
    failure_count = $failureCount
    total_count = $InputFiles.Count
    successful_records = $batchRecords
}
$batchPath = Save-DiagnosticRecord $batchSummary 'batch_summary'

Write-Host ''
Write-Host ('=' * 78) -ForegroundColor DarkGray
Write-Host ("Finished. Success: {0}   Failed: {1}" -f $successCount, $failureCount) -ForegroundColor White
Write-Info ("Batch diagnostic: {0}" -f $batchPath)

if ($failureCount -gt 0) { exit 2 }
exit 0
