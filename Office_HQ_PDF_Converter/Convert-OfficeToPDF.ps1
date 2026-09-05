[CmdletBinding()]
param(
    [Parameter(Position = 0, ValueFromRemainingArguments = $true)]
    [string[]]$InputFiles
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$SupportedExtensions = @('.docx', '.doc', '.xlsx', '.xls', '.pptx', '.ppt')
$ModernExtensions = @('.docx', '.xlsx', '.pptx')
$LegacyExtensions = @('.doc', '.xls', '.ppt')
$ScriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$LogsDirectory = Join-Path $ScriptRoot 'logs'
$DiagnosticsDirectory = Join-Path $ScriptRoot 'diagnostics'
$TempDirectory = Join-Path $ScriptRoot 'temp'
$RebuilderScript = Join-Path $ScriptRoot 'src\HQImageRebuilder.py'

foreach ($dir in @($LogsDirectory, $DiagnosticsDirectory, $TempDirectory)) {
    New-Item -ItemType Directory -Path $dir -Force | Out-Null
}

$SessionStamp = Get-Date -Format 'yyyyMMdd_HHmmss_fff'
$Script:LogPath = Join-Path $LogsDirectory ("{0}_Office_HQ_PDF.log" -f $SessionStamp)

function Get-ExceptionText($Exception) {
    if ($null -eq $Exception) { return '' }
    $hr = 'n/a'
    try { $hr = ('0x{0:X8}' -f ($Exception.HResult -band 0xffffffff)) } catch { }
    return ("Type={0}; HRESULT={1}; Message={2}" -f $Exception.GetType().FullName, $hr, $Exception.Message)
}

function Write-Log([string]$Level, [string]$Message, $Exception = $null) {
    $ts = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss.fff zzz')
    $line = "[{0}] [{1}] {2}" -f $ts, $Level.ToUpperInvariant(), $Message
    if ($null -ne $Exception) { $line += ' | ' + (Get-ExceptionText $Exception) }
    Add-Content -LiteralPath $Script:LogPath -Value $line -Encoding UTF8
    switch ($Level.ToUpperInvariant()) {
        'OK'    { Write-Host $line -ForegroundColor Green }
        'WARN'  { Write-Host $line -ForegroundColor Yellow }
        'ERROR' { Write-Host $line -ForegroundColor Red }
        'DEBUG' { Write-Host $line -ForegroundColor DarkGray }
        default { Write-Host $line -ForegroundColor Cyan }
    }
}

function Release-ComObject($Object) {
    if ($null -ne $Object) {
        try {
            if ([Runtime.InteropServices.Marshal]::IsComObject($Object)) {
                [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($Object)
            }
        } catch { }
    }
}

function Force-ComCleanup {
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}

function Disable-OfficeMacros($Application) {
    try { $Application.AutomationSecurity = 3 } catch { }
}

function Get-UniquePdfPath([string]$SourcePath, [string]$Suffix) {
    $dir = [IO.Path]::GetDirectoryName($SourcePath)
    $base = [IO.Path]::GetFileNameWithoutExtension($SourcePath)
    $candidate = Join-Path $dir ($base + $Suffix + '.pdf')
    if (-not (Test-Path -LiteralPath $candidate)) { return $candidate }
    $i = 2
    while ($true) {
        $candidate = Join-Path $dir ("{0}{1}_{2}.pdf" -f $base, $Suffix, $i)
        if (-not (Test-Path -LiteralPath $candidate)) { return $candidate }
        $i++
    }
}

function Get-TempPdfPath([string]$SourcePath) {
    $base = [IO.Path]::GetFileNameWithoutExtension($SourcePath)
    return (Join-Path $TempDirectory ("{0}_{1}_native.pdf" -f $base, [guid]::NewGuid().ToString('N')))
}

function Get-TempOoxmlPath([string]$SourcePath, [string]$Extension) {
    $base = [IO.Path]::GetFileNameWithoutExtension($SourcePath)
    return (Join-Path $TempDirectory ("{0}_{1}{2}" -f $base, [guid]::NewGuid().ToString('N'), $Extension))
}

function Find-PythonRunner {
    $py = Get-Command 'py.exe' -ErrorAction SilentlyContinue
    if ($null -ne $py) {
        return [pscustomobject]@{ Exe = $py.Source; Prefix = @('-3'); Label = 'py -3' }
    }
    foreach ($name in @('python.exe', 'python3.exe')) {
        $cmd = Get-Command $name -ErrorAction SilentlyContinue
        if ($null -ne $cmd) {
            return [pscustomobject]@{ Exe = $cmd.Source; Prefix = @(); Label = $name }
        }
    }
    return $null
}

function Invoke-Python($Runner, [string[]]$Arguments, [switch]$Quiet) {
    $allArgs = @()
    $allArgs += $Runner.Prefix
    $allArgs += $Arguments
    $exitCode = 1
    try {
        if ($Quiet) {
            & $Runner.Exe @allArgs *> $null
            $exitCode = $LASTEXITCODE
        } else {
            $nativeOutput = @(& $Runner.Exe @allArgs 2>&1)
            $exitCode = $LASTEXITCODE
            foreach ($entry in $nativeOutput) {
                $text = ([string]$entry).TrimEnd()
                if ($text) { Write-Log 'DEBUG' ("Python: {0}" -f $text) }
            }
        }
    } catch {
        Write-Log 'ERROR' 'Python process invocation failed.' $_.Exception
        $exitCode = 1
    }
    if ($null -eq $exitCode) { $exitCode = 1 }
    return [int]$exitCode
}

function Get-PythonText($Runner, [string[]]$Arguments) {
    $allArgs = @()
    $allArgs += $Runner.Prefix
    $allArgs += $Arguments
    try { return ((& $Runner.Exe @allArgs 2>&1 | Out-String).Trim()) }
    catch { return $_.Exception.Message }
}

function Test-PythonPackage($Runner, [string]$PackageName) {
    return ((Invoke-Python $Runner @('-m', 'pip', 'show', $PackageName) -Quiet) -eq 0)
}

function Ensure-HqPythonDependencies($Runner) {
    Write-Log 'INFO' ("Python runtime: {0}" -f (Get-PythonText $Runner @('--version')))
    Write-Log 'INFO' ("Python package manager: {0}" -f (Get-PythonText $Runner @('-m', 'pip', '--version')))
    Write-Log 'INFO' 'Checking production packages using pip metadata: PyMuPDF + Pillow.'

    $hasMuPdf = Test-PythonPackage $Runner 'pymupdf'
    $hasPillow = Test-PythonPackage $Runner 'pillow'
    Write-Log 'INFO' ("Package status: pymupdf={0}; pillow={1}" -f $hasMuPdf, $hasPillow)
    if ($hasMuPdf -and $hasPillow) { return $true }

    $missing = @()
    if (-not $hasMuPdf) { $missing += 'pymupdf' }
    if (-not $hasPillow) { $missing += 'pillow' }

    Write-Log 'WARN' ("Missing package(s): {0}. Trying configured pip index." -f ($missing -join ', '))
    if ((Invoke-Python $Runner (@('-m', 'pip', 'install', '--user') + $missing)) -eq 0) {
        if ((Test-PythonPackage $Runner 'pymupdf') -and (Test-PythonPackage $Runner 'pillow')) { return $true }
    }

    Write-Log 'WARN' 'Configured index did not provide all dependencies. Trying official PyPI once.'
    if ((Invoke-Python $Runner (@('-m', 'pip', 'install', '--user', '--index-url', 'https://pypi.org/simple') + $missing)) -ne 0) {
        Write-Log 'ERROR' 'Automatic Python dependency installation failed.'
        return $false
    }
    return ((Test-PythonPackage $Runner 'pymupdf') -and (Test-PythonPackage $Runner 'pillow'))
}

function Export-WordNative([string]$SourcePath, [string]$PdfPath) {
    $word = $null; $doc = $null
    try {
        Write-Log 'INFO' 'Starting Word layout engine.'
        $word = New-Object -ComObject Word.Application
        $word.Visible = $false; $word.DisplayAlerts = 0
        Disable-OfficeMacros $word
        Write-Log 'INFO' ("Word Version={0}, Build={1}" -f $word.Version, $word.Build)
        $doc = $word.Documents.Open($SourcePath, $false, $true, $false)
        try { [void]$doc.Repaginate() } catch { }
        Write-Log 'INFO' 'Exporting structural PDF with classic Word ExportAsFixedFormat.'
        [void]$doc.ExportAsFixedFormat($PdfPath, 17, $false, 0, 0, 1, 1, 0, $true, $false, 1, $true, $true, $false)
    } finally {
        if ($null -ne $doc) { try { $doc.Close(0) } catch { } }
        Release-ComObject $doc
        if ($null -ne $word) { try { $word.Quit() } catch { } }
        Release-ComObject $word
        Force-ComCleanup
    }
}

function Export-ExcelNative([string]$SourcePath, [string]$PdfPath) {
    $excel = $null; $book = $null
    try {
        Write-Log 'INFO' 'Starting Excel layout engine.'
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false; $excel.DisplayAlerts = $false
        Disable-OfficeMacros $excel
        Write-Log 'INFO' ("Excel Version={0}, Build={1}" -f $excel.Version, $excel.Build)
        $book = $excel.Workbooks.Open($SourcePath, 0, $true)

        Write-Log 'INFO' 'Applying Excel PDF layout policy: A4 paper, preserve current orientation, fit every sheet to 1 page wide, unlimited pages tall.'
        foreach ($sheet in @($book.Worksheets)) {
            $pageSetup = $null
            try {
                $pageSetup = $sheet.PageSetup
                $orientation = $null
                try { $orientation = $pageSetup.Orientation } catch { }
                $pageSetup.PaperSize = 9
                $pageSetup.Zoom = $false
                $pageSetup.FitToPagesWide = 1
                $pageSetup.FitToPagesTall = $false
                Write-Log 'DEBUG' ("Excel sheet '{0}': PaperSize=A4; FitToPagesWide=1; FitToPagesTall=False; Orientation={1}" -f $sheet.Name, $orientation)
            } catch {
                Write-Log 'WARN' ("Could not apply A4/fit-to-width settings to Excel sheet '{0}'. Existing page setup will be used." -f $sheet.Name) $_.Exception
            } finally {
                Release-ComObject $pageSetup
                Release-ComObject $sheet
            }
        }

        Write-Log 'INFO' 'Exporting structural PDF with Excel xlQualityStandard after page-setup normalization.'
        [void]$book.ExportAsFixedFormat(0, $PdfPath, 0, $true, $true)
    } finally {
        if ($null -ne $book) { try { $book.Close($false) } catch { } }
        Release-ComObject $book
        if ($null -ne $excel) { try { $excel.Quit() } catch { } }
        Release-ComObject $excel
        Force-ComCleanup
    }
}

function Export-PowerPointNative([string]$SourcePath, [string]$PdfPath) {
    $ppt = $null; $pres = $null; $pageSetup = $null
    try {
        Write-Log 'INFO' 'Starting PowerPoint layout engine.'
        $ppt = New-Object -ComObject PowerPoint.Application
        Disable-OfficeMacros $ppt
        Write-Log 'INFO' ("PowerPoint Version={0}, Build={1}" -f $ppt.Version, $ppt.Build)
        $pres = $ppt.Presentations.Open($SourcePath, -1, 0, 0)

        try {
            $pageSetup = $pres.PageSetup
            $w = [double]$pageSetup.SlideWidth
            $h = [double]$pageSetup.SlideHeight
            $ratio = if ($h -ne 0) { [math]::Round($w / $h, 6) } else { $null }
            Write-Log 'INFO' ("PowerPoint slide canvas: {0:N2} x {1:N2} pt; aspect ratio={2}. PDF will preserve this slide ratio." -f $w, $h, $ratio)
        } catch {
            Write-Log 'WARN' 'Could not read PowerPoint slide dimensions; SaveAs PDF will still preserve the presentation canvas.' $_.Exception
        }

        Write-Log 'INFO' 'Exporting structural PDF with PowerPoint SaveAs PDF (stable path; slide canvas ratio preserved).'
        [void]$pres.SaveAs($PdfPath, 32)
    } finally {
        Release-ComObject $pageSetup
        if ($null -ne $pres) { try { $pres.Close() } catch { } }
        Release-ComObject $pres
        if ($null -ne $ppt) { try { $ppt.Quit() } catch { } }
        Release-ComObject $ppt
        Force-ComCleanup
    }
}

function Export-StructuralPdf([string]$SourcePath, [string]$PdfPath, [string]$Extension) {
    switch ($Extension) {
        '.docx' { Export-WordNative $SourcePath $PdfPath }
        '.doc'  { Export-WordNative $SourcePath $PdfPath }
        '.xlsx' { Export-ExcelNative $SourcePath $PdfPath }
        '.xls'  { Export-ExcelNative $SourcePath $PdfPath }
        '.pptx' { Export-PowerPointNative $SourcePath $PdfPath }
        '.ppt'  { Export-PowerPointNative $SourcePath $PdfPath }
        default { throw "Unsupported extension: $Extension" }
    }
}

function Convert-LegacyWordToOoxml([string]$SourcePath, [string]$OoxmlPath) {
    $word = $null; $doc = $null
    try {
        Write-Log 'INFO' 'Upgrading legacy DOC to temporary DOCX for HQ media recovery.'
        $word = New-Object -ComObject Word.Application
        $word.Visible = $false; $word.DisplayAlerts = 0
        Disable-OfficeMacros $word
        $doc = $word.Documents.Open($SourcePath, $false, $true, $false)
        [void]$doc.SaveAs2($OoxmlPath, 12)
    } finally {
        if ($null -ne $doc) { try { $doc.Close(0) } catch { } }
        Release-ComObject $doc
        if ($null -ne $word) { try { $word.Quit() } catch { } }
        Release-ComObject $word
        Force-ComCleanup
    }
}

function Convert-LegacyExcelToOoxml([string]$SourcePath, [string]$OoxmlPath) {
    $excel = $null; $book = $null
    try {
        Write-Log 'INFO' 'Upgrading legacy XLS to temporary XLSX for HQ media recovery.'
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false; $excel.DisplayAlerts = $false
        Disable-OfficeMacros $excel
        $book = $excel.Workbooks.Open($SourcePath, 0, $true)
        [void]$book.SaveAs($OoxmlPath, 51)
    } finally {
        if ($null -ne $book) { try { $book.Close($false) } catch { } }
        Release-ComObject $book
        if ($null -ne $excel) { try { $excel.Quit() } catch { } }
        Release-ComObject $excel
        Force-ComCleanup
    }
}

function Convert-LegacyPowerPointToOoxml([string]$SourcePath, [string]$OoxmlPath) {
    $ppt = $null; $pres = $null
    try {
        Write-Log 'INFO' 'Upgrading legacy PPT to temporary PPTX for HQ media recovery.'
        $ppt = New-Object -ComObject PowerPoint.Application
        Disable-OfficeMacros $ppt
        $pres = $ppt.Presentations.Open($SourcePath, -1, 0, 0)
        [void]$pres.SaveAs($OoxmlPath, 24)
    } finally {
        if ($null -ne $pres) { try { $pres.Close() } catch { } }
        Release-ComObject $pres
        if ($null -ne $ppt) { try { $ppt.Quit() } catch { } }
        Release-ComObject $ppt
        Force-ComCleanup
    }
}

function Get-LegacyUpgradeExtension([string]$Extension) {
    switch ($Extension) {
        '.doc' { return '.docx' }
        '.xls' { return '.xlsx' }
        '.ppt' { return '.pptx' }
        default { return $null }
    }
}

function Convert-LegacyToOoxml([string]$SourcePath, [string]$OoxmlPath, [string]$Extension) {
    switch ($Extension) {
        '.doc' { Convert-LegacyWordToOoxml $SourcePath $OoxmlPath }
        '.xls' { Convert-LegacyExcelToOoxml $SourcePath $OoxmlPath }
        '.ppt' { Convert-LegacyPowerPointToOoxml $SourcePath $OoxmlPath }
        default { throw "Not a supported legacy Office format: $Extension" }
    }
}

function Invoke-HqRebuild([string]$SourcePath, [string]$NativePdf, [string]$FinalPdf) {
    if (-not (Test-Path -LiteralPath $RebuilderScript)) {
        Write-Log 'ERROR' ("HQ rebuilder is missing: {0}" -f $RebuilderScript)
        return $false
    }
    $runner = Find-PythonRunner
    if ($null -eq $runner) {
        Write-Log 'ERROR' 'Python 3 was not found.'
        return $false
    }
    Write-Log 'INFO' ("Python runner: {0} ({1})" -f $runner.Exe, $runner.Label)
    if (-not (Ensure-HqPythonDependencies $runner)) { return $false }

    $safe = [IO.Path]::GetFileNameWithoutExtension($SourcePath) -replace '[^A-Za-z0-9._-]', '_'
    $report = Join-Path $DiagnosticsDirectory ("{0}_{1}_image_rebuild.json" -f (Get-Date -Format 'yyyyMMdd_HHmmss_fff'), $safe)
    Write-Log 'INFO' 'Starting lossless OOXML raster-image restoration.'

    $exitCode = Invoke-Python $runner @(
        $RebuilderScript,
        '--source-office', $SourcePath,
        '--input-pdf', $NativePdf,
        '--output-pdf', $FinalPdf,
        '--report', $report
    )

    Write-Log 'DEBUG' ("Python rebuilder exit code: {0}" -f $exitCode)
    if ($exitCode -ne 0) {
        Write-Log 'ERROR' ("HQ image restoration failed with exit code {0}." -f $exitCode)
        return $false
    }
    if (-not (Test-Path -LiteralPath $FinalPdf)) {
        Write-Log 'ERROR' 'HQ rebuilder returned success but no final PDF exists.'
        return $false
    }
    Write-Log 'OK' ("HQ image restoration completed. Diagnostic: {0}" -f [IO.Path]::GetFileName($report))
    return $true
}

function Save-SessionDiagnostic(
    [string]$SourcePath,
    [string]$OutputPath,
    [string]$Mode,
    [double]$Seconds,
    [string]$Message,
    [bool]$LegacyUpgradeAttempted = $false,
    [bool]$LegacyUpgradeSucceeded = $false
) {
    $source = Get-Item -LiteralPath $SourcePath
    $size = $null
    if (Test-Path -LiteralPath $OutputPath) { $size = (Get-Item -LiteralPath $OutputPath).Length }
    $record = [ordered]@{
        schema_version = 8
        timestamp_local = (Get-Date).ToString('o')
        source_file = $source.Name
        source_extension = $source.Extension.ToLowerInvariant()
        source_size_bytes = $source.Length
        output_file = [IO.Path]::GetFileName($OutputPath)
        output_size_bytes = $size
        result_mode = $Mode
        duration_seconds = [math]::Round($Seconds, 3)
        message = $Message
        production_pipeline = 'Office layout -> OOXML source raster restoration'
        excel_layout_policy = 'A4; FitToPagesWide=1; FitToPagesTall=False; preserve orientation'
        powerpoint_layout_policy = 'preserve source slide canvas aspect ratio'
        pptx_image_policy = 'crop-aware DrawingML srcRect semantic matching'
        legacy_upgrade_attempted = $LegacyUpgradeAttempted
        legacy_upgrade_succeeded = $LegacyUpgradeSucceeded
        legacy_upgrade_policy = 'DOC->DOCX(12); XLS->XLSX(51); PPT->PPTX(24); temp only'
        python_process_handling = 'stdout captured separately; integer exit code only'
        log_file = [IO.Path]::GetFileName($Script:LogPath)
    }
    $safe = [IO.Path]::GetFileNameWithoutExtension($SourcePath) -replace '[^A-Za-z0-9._-]', '_'
    $path = Join-Path $DiagnosticsDirectory ("{0}_{1}_conversion.json" -f (Get-Date -Format 'yyyyMMdd_HHmmss_fff'), $safe)
    $record | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $path -Encoding UTF8
}

function Convert-One([string]$InputPath) {
    $started = Get-Date
    $resolved = (Resolve-Path -LiteralPath $InputPath).Path
    $item = Get-Item -LiteralPath $resolved
    if ($item.PSIsContainer) { throw 'Folders are not supported.' }
    $ext = $item.Extension.ToLowerInvariant()
    if ($SupportedExtensions -notcontains $ext) { throw "Unsupported file type: $ext" }

    Write-Log 'INFO' ('=' * 72)
    Write-Log 'INFO' ("Source: {0} ({1:N3} MB)" -f $item.Name, ($item.Length / 1MB))

    $nativeTemp = Get-TempPdfPath $resolved
    $legacyTemp = $null
    $workingSource = $resolved
    $workingExt = $ext
    $legacyAttempted = $false
    $legacySucceeded = $false

    try {
        if ($LegacyExtensions -contains $ext) {
            $legacyAttempted = $true
            $upgradeExt = Get-LegacyUpgradeExtension $ext
            $legacyTemp = Get-TempOoxmlPath $resolved $upgradeExt
            try {
                Convert-LegacyToOoxml $resolved $legacyTemp $ext
                if (-not (Test-Path -LiteralPath $legacyTemp)) {
                    throw 'Legacy-to-OOXML conversion returned without creating the temporary file.'
                }
                $workingSource = $legacyTemp
                $workingExt = $upgradeExt
                $legacySucceeded = $true
                Write-Log 'OK' ("Legacy format upgraded in temp/: {0} -> {1} ({2:N3} MB)" -f $ext, $upgradeExt, ((Get-Item $legacyTemp).Length / 1MB))
            } catch {
                Write-Log 'WARN' 'Legacy-to-OOXML upgrade failed. Falling back to Office-native PDF export without HQ media recovery.' $_.Exception
                $workingSource = $resolved
                $workingExt = $ext
            }
        }

        Export-StructuralPdf $workingSource $nativeTemp $workingExt
        if (-not (Test-Path -LiteralPath $nativeTemp)) { throw 'Office structural PDF was not created.' }
        Write-Log 'INFO' ("Structural PDF created in temp/: {0:N3} MB" -f ((Get-Item $nativeTemp).Length / 1MB))

        $canRebuild = ($ModernExtensions -contains $workingExt)
        if ($canRebuild) {
            $final = Get-UniquePdfPath $resolved '_HQ'
            if (Invoke-HqRebuild $workingSource $nativeTemp $final) {
                $seconds = ((Get-Date) - $started).TotalSeconds
                $mode = if ($legacySucceeded) { 'HQ_REBUILT_FROM_LEGACY' } else { 'HQ_REBUILT' }
                $message = if ($legacySucceeded) {
                    'Legacy Office file upgraded to temporary OOXML; layout preserved and OOXML raster media restored.'
                } else {
                    'Office layout preserved; OOXML raster media restored.'
                }
                Write-Log 'OK' ("{0}: {1} ({2:N3} MB) in {3:N2} s" -f $mode, [IO.Path]::GetFileName($final), ((Get-Item $final).Length / 1MB), $seconds)
                Save-SessionDiagnostic $resolved $final $mode $seconds $message $legacyAttempted $legacySucceeded
                return $true
            }

            if (Test-Path -LiteralPath $final) { try { Remove-Item -LiteralPath $final -Force } catch { } }
            $fallback = Get-UniquePdfPath $resolved '_NATIVE_ONLY'
            Move-Item -LiteralPath $nativeTemp -Destination $fallback -Force
            $seconds = ((Get-Date) - $started).TotalSeconds
            Write-Log 'WARN' ("NATIVE_ONLY: HQ rebuild failed. Preserved {0}." -f [IO.Path]::GetFileName($fallback))
            Save-SessionDiagnostic $resolved $fallback 'NATIVE_ONLY' $seconds 'HQ restoration failed; Office raster quality retained.' $legacyAttempted $legacySucceeded
            return $false
        }

        $legacyOut = Get-UniquePdfPath $resolved '_NATIVE_ONLY'
        Move-Item -LiteralPath $nativeTemp -Destination $legacyOut -Force
        $seconds = ((Get-Date) - $started).TotalSeconds
        Write-Log 'WARN' ("NATIVE_ONLY: legacy binary format could not be upgraded to OOXML; Office-native PDF preserved as {0}." -f [IO.Path]::GetFileName($legacyOut))
        Save-SessionDiagnostic $resolved $legacyOut 'NATIVE_ONLY' $seconds 'Legacy-to-OOXML upgrade failed; no recoverable OOXML media package.' $legacyAttempted $legacySucceeded
        return $false
    } catch {
        $seconds = ((Get-Date) - $started).TotalSeconds
        Write-Log 'ERROR' ("FAILED: {0}" -f $item.Name) $_.Exception
        try { Save-SessionDiagnostic $resolved $nativeTemp 'FAILED' $seconds (Get-ExceptionText $_.Exception) $legacyAttempted $legacySucceeded } catch { }
        return $false
    } finally {
        if (Test-Path -LiteralPath $nativeTemp) {
            try { Remove-Item -LiteralPath $nativeTemp -Force } catch { }
        }
        if ($null -ne $legacyTemp -and (Test-Path -LiteralPath $legacyTemp)) {
            try { Remove-Item -LiteralPath $legacyTemp -Force } catch { }
        }
    }
}

Write-Log 'INFO' 'Office HQ PDF Converter production pipeline started.'
Write-Log 'INFO' 'Strategy: Office layout engine -> temporary structural PDF -> lossless OOXML raster restoration.'
Write-Log 'INFO' 'Excel policy: A4, 1 page wide, unlimited pages tall, preserve sheet orientation.'
Write-Log 'INFO' 'PowerPoint policy: preserve source slide canvas aspect ratio.'
Write-Log 'INFO' 'PPTX image policy: parse slide relationships + DrawingML srcRect crop before HQ matching.'
Write-Log 'INFO' 'Legacy policy: DOC/XLS/PPT are upgraded to temporary OOXML first; originals are never modified.'
Write-Log 'INFO' 'Python stdout is isolated from process exit codes for Windows PowerShell 5.1 compatibility.'

if ($null -eq $InputFiles -or $InputFiles.Count -eq 0) {
    Write-Log 'ERROR' 'No input files. Drag Office documents onto DragDrop_Office_to_PDF.cmd.'
    exit 1
}

$ok = 0; $failed = 0
foreach ($file in $InputFiles) {
    try {
        if (Convert-One $file) { $ok++ } else { $failed++ }
    } catch {
        Write-Log 'ERROR' ("Unhandled conversion error: {0}" -f $file) $_.Exception
        $failed++
    }
}

Write-Log 'INFO' ('=' * 72)
Write-Log 'INFO' ("Finished. Success={0}; FailedOrDegraded={1}" -f $ok, $failed)
Write-Log 'INFO' ("Log: {0}" -f [IO.Path]::GetFileName($Script:LogPath))
if ($failed -gt 0) { exit 2 }
exit 0
