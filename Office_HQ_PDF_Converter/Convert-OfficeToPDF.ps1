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

function Find-PythonRunner {
    foreach ($name in @('py.exe', 'python.exe', 'python3.exe')) {
        $cmd = Get-Command $name -ErrorAction SilentlyContinue
        if ($null -ne $cmd) {
            if ($name -eq 'py.exe') {
                return [pscustomobject]@{ Exe = $cmd.Source; Prefix = @('-3') }
            }
            return [pscustomobject]@{ Exe = $cmd.Source; Prefix = @() }
        }
    }
    return $null
}

function Invoke-Python($Runner, [string[]]$Arguments) {
    $allArgs = @()
    $allArgs += $Runner.Prefix
    $allArgs += $Arguments
    & $Runner.Exe @allArgs
    return $LASTEXITCODE
}

function Ensure-HqPythonDependencies($Runner) {
    Write-Log 'INFO' 'Checking Python modules: pymupdf + Pillow.'
    $check = Invoke-Python $Runner @('-c', 'import pymupdf, PIL; print("HQ Python dependencies OK")')
    if ($check -eq 0) { return $true }

    Write-Log 'WARN' 'Required Python modules are missing. Trying the currently configured pip index.'
    $install = Invoke-Python $Runner @('-m', 'pip', 'install', '--user', 'pymupdf', 'pillow')
    if ($install -eq 0) {
        $check2 = Invoke-Python $Runner @('-c', 'import pymupdf, PIL')
        if ($check2 -eq 0) { return $true }
    }

    Write-Log 'WARN' 'Configured pip index did not provide all dependencies. Trying official PyPI once.'
    $install2 = Invoke-Python $Runner @('-m', 'pip', 'install', '--user', '--index-url', 'https://pypi.org/simple', 'pymupdf', 'pillow')
    if ($install2 -ne 0) {
        Write-Log 'ERROR' 'Automatic Python dependency installation failed. Run Collect_Local_Environment.cmd and push the report for diagnosis.'
        return $false
    }
    $check3 = Invoke-Python $Runner @('-c', 'import pymupdf, PIL')
    return ($check3 -eq 0)
}

function Export-WordNative([string]$SourcePath, [string]$PdfPath) {
    $word = $null; $doc = $null
    try {
        Write-Log 'INFO' 'Starting Word layout engine.'
        $word = New-Object -ComObject Word.Application
        $word.Visible = $false
        $word.DisplayAlerts = 0
        Disable-OfficeMacros $word
        Write-Log 'INFO' ("Word Version={0}, Build={1}" -f $word.Version, $word.Build)
        $doc = $word.Documents.Open($SourcePath, $false, $true, $false)
        try { [void]$doc.Repaginate() } catch { }
        Write-Log 'INFO' 'Exporting structural PDF with classic Word ExportAsFixedFormat. Raster image quality will be restored afterwards.'
        [void]$doc.ExportAsFixedFormat(
            $PdfPath, 17, $false, 0, 0, 1, 1, 0,
            $true, $false, 1, $true, $true, $false
        )
    }
    finally {
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
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        Disable-OfficeMacros $excel
        Write-Log 'INFO' ("Excel Version={0}, Build={1}" -f $excel.Version, $excel.Build)
        $book = $excel.Workbooks.Open($SourcePath, 0, $true)
        Write-Log 'INFO' 'Exporting structural PDF with Excel xlQualityStandard; OOXML raster media will be restored afterwards.'
        [void]$book.ExportAsFixedFormat(0, $PdfPath, 0, $true, $true)
    }
    finally {
        if ($null -ne $book) { try { $book.Close($false) } catch { } }
        Release-ComObject $book
        if ($null -ne $excel) { try { $excel.Quit() } catch { } }
        Release-ComObject $excel
        Force-ComCleanup
    }
}

function Export-PowerPointNative([string]$SourcePath, [string]$PdfPath) {
    $ppt = $null; $pres = $null
    try {
        Write-Log 'INFO' 'Starting PowerPoint layout engine.'
        $ppt = New-Object -ComObject PowerPoint.Application
        Disable-OfficeMacros $ppt
        Write-Log 'INFO' ("PowerPoint Version={0}, Build={1}" -f $ppt.Version, $ppt.Build)
        $pres = $ppt.Presentations.Open($SourcePath, -1, 0, 0)
        Write-Log 'INFO' 'Exporting structural PDF with PowerPoint Print intent; OOXML raster media will be restored afterwards.'
        try {
            [void]$pres.ExportAsFixedFormat(
                $PdfPath, 2, 2, 0, 1, 1, 0,
                $null, 1, '', $true, $true, $true, $true, $false
            )
        } catch {
            Write-Log 'WARN' 'PowerPoint ExportAsFixedFormat failed; using SaveAs PDF as structural fallback.' $_.Exception
            [void]$pres.SaveAs($PdfPath, 32)
        }
    }
    finally {
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

function Invoke-HqRebuild([string]$SourcePath, [string]$NativePdf, [string]$FinalPdf) {
    if (-not (Test-Path -LiteralPath $RebuilderScript)) {
        Write-Log 'ERROR' ("HQ rebuilder is missing: {0}" -f $RebuilderScript)
        return $false
    }

    $runner = Find-PythonRunner
    if ($null -eq $runner) {
        Write-Log 'ERROR' 'Python 3 was not found. Run Collect_Local_Environment.cmd and push the report for diagnosis.'
        return $false
    }
    Write-Log 'INFO' ("Python runner: {0}" -f $runner.Exe)
    [void](Invoke-Python $runner @('-c', 'import sys; print("Python=" + sys.version.replace("\n"," ")); print("Executable=" + sys.executable)'))

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

    if ($exitCode -ne 0) {
        Write-Log 'ERROR' ("HQ image restoration failed with exit code {0}." -f $exitCode)
        return $false
    }
    if (-not (Test-Path -LiteralPath $FinalPdf)) {
        Write-Log 'ERROR' 'HQ rebuilder exited successfully but no final PDF was created.'
        return $false
    }

    Write-Log 'OK' ("HQ image restoration completed. Diagnostic: {0}" -f [IO.Path]::GetFileName($report))
    return $true
}

function Save-SessionDiagnostic([string]$SourcePath, [string]$OutputPath, [string]$Mode, [double]$Seconds, [string]$Message) {
    $source = Get-Item -LiteralPath $SourcePath
    $size = $null
    if (Test-Path -LiteralPath $OutputPath) { $size = (Get-Item -LiteralPath $OutputPath).Length }
    $record = [ordered]@{
        schema_version = 4
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
    try {
        Export-StructuralPdf $resolved $nativeTemp $ext
        if (-not (Test-Path -LiteralPath $nativeTemp)) { throw 'Office structural PDF was not created.' }
        $nativeSize = (Get-Item -LiteralPath $nativeTemp).Length
        Write-Log 'INFO' ("Structural PDF created in temp/: {0:N3} MB" -f ($nativeSize / 1MB))

        if ($ModernExtensions -contains $ext) {
            $final = Get-UniquePdfPath $resolved '_HQ'
            if (Invoke-HqRebuild $resolved $nativeTemp $final) {
                $seconds = ((Get-Date) - $started).TotalSeconds
                $finalSize = (Get-Item -LiteralPath $final).Length
                Write-Log 'OK' ("HQ_REBUILT: {0} ({1:N3} MB) in {2:N2} s" -f [IO.Path]::GetFileName($final), ($finalSize / 1MB), $seconds)
                Save-SessionDiagnostic $resolved $final 'HQ_REBUILT' $seconds 'Office layout preserved; source OOXML raster media restored losslessly.'
                return $true
            }

            if (Test-Path -LiteralPath $final) { try { Remove-Item -LiteralPath $final -Force } catch { } }
            $fallback = Get-UniquePdfPath $resolved '_NATIVE_ONLY'
            Move-Item -LiteralPath $nativeTemp -Destination $fallback -Force
            $seconds = ((Get-Date) - $started).TotalSeconds
            Write-Log 'WARN' ("NATIVE_ONLY: HQ rebuild failed. Low-resolution Office PDF preserved as {0}." -f [IO.Path]::GetFileName($fallback))
            Save-SessionDiagnostic $resolved $fallback 'NATIVE_ONLY' $seconds 'HQ rebuild unavailable or failed; raster images may be approximately 200 PPI.'
            return $false
        }

        $legacyOut = Get-UniquePdfPath $resolved '_NATIVE_ONLY'
        Move-Item -LiteralPath $nativeTemp -Destination $legacyOut -Force
        $seconds = ((Get-Date) - $started).TotalSeconds
        Write-Log 'WARN' ("NATIVE_ONLY: legacy binary format currently uses Office-native export: {0}." -f [IO.Path]::GetFileName($legacyOut))
        Save-SessionDiagnostic $resolved $legacyOut 'NATIVE_ONLY' $seconds 'Legacy DOC/XLS/PPT does not expose an OOXML media package for the current restoration stage.'
        return $true
    }
    catch {
        $seconds = ((Get-Date) - $started).TotalSeconds
        Write-Log 'ERROR' ("FAILED: {0}" -f $item.Name) $_.Exception
        try { Save-SessionDiagnostic $resolved $nativeTemp 'FAILED' $seconds (Get-ExceptionText $_.Exception) } catch { }
        return $false
    }
    finally {
        if (Test-Path -LiteralPath $nativeTemp) {
            try { Remove-Item -LiteralPath $nativeTemp -Force } catch { }
        }
    }
}

Write-Log 'INFO' 'Office HQ PDF Converter production pipeline started.'
Write-Log 'INFO' 'Strategy: Office layout engine -> temporary structural PDF -> lossless OOXML raster restoration.'
Write-Log 'INFO' 'Production path intentionally does NOT use Acrobat PDFMaker, ExportAsFixedFormat2, or ImageHash.'

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
