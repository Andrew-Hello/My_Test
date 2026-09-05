[CmdletBinding()]
param(
    [Parameter(Position = 0, ValueFromRemainingArguments = $true)]
    [string[]]$InputFiles
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$SupportedExtensions = @('.docx', '.doc', '.xlsx', '.xls', '.pptx', '.ppt')
$ScriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$LogsDirectory = Join-Path $ScriptRoot 'logs'
$DiagnosticsDirectory = Join-Path $ScriptRoot 'diagnostics'
New-Item -ItemType Directory -Path $LogsDirectory -Force | Out-Null
New-Item -ItemType Directory -Path $DiagnosticsDirectory -Force | Out-Null

$SessionStamp = Get-Date -Format 'yyyyMMdd_HHmmss_fff'
$Script:LogPath = Join-Path $LogsDirectory ("{0}_Office_HQ_PDF.log" -f $SessionStamp)

function Get-ExceptionText($Exception) {
    if ($null -eq $Exception) { return $null }
    $hresult = $null
    try { $hresult = ('0x{0:X8}' -f ($Exception.HResult -band 0xffffffff)) } catch { }
    $parts = @()
    $parts += ("Type={0}" -f $Exception.GetType().FullName)
    if ($hresult) { $parts += ("HRESULT={0}" -f $hresult) }
    $parts += ("Message={0}" -f $Exception.Message)
    if ($Exception.InnerException) {
        $parts += ("Inner={0}" -f $Exception.InnerException.Message)
    }
    return ($parts -join '; ')
}

function Write-Log([string]$Level, [string]$Message, $Exception = $null) {
    $timestamp = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss.fff zzz')
    $line = "[{0}] [{1}] {2}" -f $timestamp, $Level.ToUpperInvariant(), $Message
    if ($null -ne $Exception) {
        $line += " | " + (Get-ExceptionText $Exception)
    }
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
    $safe = $Name
    foreach ($c in [System.IO.Path]::GetInvalidFileNameChars()) {
        $safe = $safe.Replace([string]$c, '_')
    }
    return $safe
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
    try {
        $Application.AutomationSecurity = 3
        Write-Log 'DEBUG' 'Office AutomationSecurity set to ForceDisable.'
    } catch {
        Write-Log 'WARN' 'Could not set Office AutomationSecurity.' $_.Exception
    }
}

function Try-GetComProperty($Object, [string]$Name) {
    try { return $Object.$Name } catch { return $null }
}

function Try-SetComProperty($Object, [string]$Name, $Value) {
    try {
        $Object.$Name = $Value
        Write-Log 'DEBUG' ("PDFMaker setting {0}={1}" -f $Name, $Value)
        return $true
    } catch {
        Write-Log 'DEBUG' ("PDFMaker setting unavailable: {0}" -f $Name)
        return $false
    }
}

function Get-PdfMakerContext($Application) {
    $addin = $null
    $pdfMaker = $null
    $progId = $null
    $description = $null

    try {
        $addin = $Application.COMAddIns.Item('PDFMaker.OfficeAddin')
    } catch {
        Write-Log 'DEBUG' 'PDFMaker.OfficeAddin was not resolved directly; enumerating COM add-ins.'
    }

    if ($null -eq $addin) {
        try {
            $count = [int]$Application.COMAddIns.Count
            for ($i = 1; $i -le $count; $i++) {
                $candidate = $null
                try {
                    $candidate = $Application.COMAddIns.Item($i)
                    $candidateProgId = [string](Try-GetComProperty $candidate 'ProgId')
                    $candidateDesc = [string](Try-GetComProperty $candidate 'Description')
                    if (($candidateProgId -match 'PDFMaker') -or ($candidateDesc -match 'PDFMaker|Acrobat')) {
                        $addin = $candidate
                        $candidate = $null
                        break
                    }
                } catch { }
                finally {
                    if ($null -ne $candidate) { Release-ComObject $candidate }
                }
            }
        } catch {
            Write-Log 'WARN' 'Could not enumerate Office COM add-ins.' $_.Exception
        }
    }

    if ($null -eq $addin) {
        Write-Log 'INFO' 'Adobe Acrobat PDFMaker COM add-in not found.'
        return $null
    }

    try { $progId = [string]$addin.ProgId } catch { }
    try { $description = [string]$addin.Description } catch { }
    Write-Log 'INFO' ("Found PDFMaker add-in. ProgId='{0}', Description='{1}'" -f $progId, $description)

    try {
        if (-not [bool]$addin.Connect) {
            Write-Log 'INFO' 'Connecting PDFMaker COM add-in...'
            $addin.Connect = $true
            Start-Sleep -Milliseconds 1200
        }
    } catch {
        Write-Log 'WARN' 'Could not force PDFMaker add-in to Connect=True.' $_.Exception
    }

    try { $pdfMaker = $addin.Object } catch { }
    if ($null -eq $pdfMaker) {
        Write-Log 'WARN' 'PDFMaker add-in exists but its automation Object is null.'
        Release-ComObject $addin
        return $null
    }

    return [pscustomobject]@{
        Addin = $addin
        PdfMaker = $pdfMaker
        ProgId = $progId
        Description = $description
    }
}

function Wait-ForPdf([string]$OutputPath, [int]$TimeoutSeconds = 180) {
    $deadline = (Get-Date).AddSeconds($TimeoutSeconds)
    [int64]$lastSize = -1
    $stableTicks = 0
    while ((Get-Date) -lt $deadline) {
        if (Test-Path -LiteralPath $OutputPath) {
            try {
                $size = (Get-Item -LiteralPath $OutputPath).Length
                if ($size -gt 0 -and $size -eq $lastSize) {
                    $stableTicks++
                    if ($stableTicks -ge 3) { return $true }
                } else {
                    $stableTicks = 0
                    $lastSize = $size
                }
            } catch { }
        }
        Start-Sleep -Milliseconds 400
    }
    return $false
}

function Invoke-PdfMakerExport($Application, [string]$OutputPath, [string]$AppKind) {
    $ctx = $null
    $settings = $null
    try {
        $ctx = Get-PdfMakerContext $Application
        if ($null -eq $ctx) { return $null }

        try {
            $ctx.PdfMaker.GetCurrentConversionSettings([ref]$settings)
            Write-Log 'INFO' 'PDFMaker.GetCurrentConversionSettings succeeded.'
        } catch {
            Write-Log 'WARN' 'GetCurrentConversionSettings failed; trying default settings.' $_.Exception
            try {
                $ctx.PdfMaker.GetDefaultConversionSettings([ref]$settings)
                Write-Log 'INFO' 'PDFMaker.GetDefaultConversionSettings succeeded.'
            } catch {
                Write-Log 'ERROR' 'Unable to obtain PDFMaker conversion settings.' $_.Exception
                return $null
            }
        }

        if ($null -eq $settings) {
            Write-Log 'ERROR' 'PDFMaker returned a null settings object.'
            return $null
        }

        $jobOptions = Try-GetComProperty $settings 'JobOptions'
        Write-Log 'INFO' ("PDFMaker current JobOptions='{0}'" -f $jobOptions)

        [void](Try-SetComProperty $settings 'OutputPDFFileName' $OutputPath)
        [void](Try-SetComProperty $settings 'PromptForPDFFilename' $false)
        [void](Try-SetComProperty $settings 'ShouldShowProgressDialog' $false)
        [void](Try-SetComProperty $settings 'ViewPDFFile' $false)
        [void](Try-SetComProperty $settings 'IsAutomation' $true)
        [void](Try-SetComProperty $settings 'IsConversionSilent' $true)
        [void](Try-SetComProperty $settings 'ConvertAllPages' $true)
        [void](Try-SetComProperty $settings 'AddLinks' $true)
        [void](Try-SetComProperty $settings 'AddBookmarks' $true)
        [void](Try-SetComProperty $settings 'CreateMetadata' $true)

        if ($AppKind -eq 'Excel') {
            [void](Try-SetComProperty $settings 'FitToOnePage' $false)
            [void](Try-SetComProperty $settings 'PrintActiveSheetOnly' $false)
            [void](Try-SetComProperty $settings 'PromptForSheetSelection' $false)
        }
        if ($AppKind -eq 'PowerPoint') {
            [void](Try-SetComProperty $settings 'ConvertHiddenSlides' $true)
            [void](Try-SetComProperty $settings 'SaveAnimations' $false)
            [void](Try-SetComProperty $settings 'SaveSlideShowTransitions' $false)
        }

        Write-Log 'INFO' 'Calling PDFMaker.CreatePDFEx...'
        $callSucceeded = $false
        try {
            [void]$ctx.PdfMaker.CreatePDFEx($settings, 0)
            $callSucceeded = $true
            Write-Log 'INFO' 'PDFMaker.CreatePDFEx(settings, 0) returned without exception.'
        } catch {
            Write-Log 'WARN' 'PDFMaker.CreatePDFEx(settings, 0) failed; trying by-ref retval form.' $_.Exception
            try {
                $retval = 0
                [void]$ctx.PdfMaker.CreatePDFEx($settings, [ref]$retval)
                $callSucceeded = $true
                Write-Log 'INFO' ("PDFMaker.CreatePDFEx by-ref returned. retval={0}" -f $retval)
            } catch {
                Write-Log 'ERROR' 'PDFMaker.CreatePDFEx failed.' $_.Exception
            }
        }

        if (-not $callSucceeded) { return $null }

        if (-not (Wait-ForPdf $OutputPath 180)) {
            Write-Log 'ERROR' 'PDFMaker call completed, but a stable PDF file did not appear before timeout.'
            return $null
        }

        $pdf = Get-Item -LiteralPath $OutputPath
        Write-Log 'OK' ("Adobe PDFMaker created PDF: {0:N3} MB" -f ($pdf.Length / 1MB))
        return [ordered]@{
            engine = 'Adobe Acrobat PDFMaker'
            pdfmaker_progid = $ctx.ProgId
            pdfmaker_description = $ctx.Description
            pdfmaker_joboptions = [string]$jobOptions
            output_size_bytes = $pdf.Length
        }
    }
    finally {
        Release-ComObject $settings
        if ($null -ne $ctx) {
            Release-ComObject $ctx.PdfMaker
            Release-ComObject $ctx.Addin
        }
    }
}

function Invoke-WordNativeExport($Doc, [string]$OutputPath) {
    $errors = @()
    try {
        Write-Log 'INFO' 'Trying Word ExportAsFixedFormat3 with OptimizeForImageQuality=True.'
        [void]$Doc.ExportAsFixedFormat3(
            $OutputPath, 17, $false, 0, 0, 1, 1, 0,
            $true, $false, 1, $true, $true, $false,
            $true, $false, $null
        )
        return [ordered]@{ engine = 'Word ExportAsFixedFormat3'; optimize_for_image_quality = $true; errors = $errors }
    } catch {
        $errors += Get-ExceptionText $_.Exception
        Write-Log 'WARN' 'ExportAsFixedFormat3 failed.' $_.Exception
    }

    try {
        Write-Log 'INFO' 'Trying Word ExportAsFixedFormat2 with OptimizeForImageQuality=True.'
        [void]$Doc.ExportAsFixedFormat2(
            $OutputPath, 17, $false, 0, 0, 1, 1, 0,
            $true, $false, 1, $true, $true, $false,
            $true, $null
        )
        return [ordered]@{ engine = 'Word ExportAsFixedFormat2'; optimize_for_image_quality = $true; errors = $errors }
    } catch {
        $errors += Get-ExceptionText $_.Exception
        Write-Log 'WARN' 'ExportAsFixedFormat2 failed.' $_.Exception
    }

    Write-Log 'WARN' 'Falling back to classic Word ExportAsFixedFormat. This engine commonly downsamples images.'
    [void]$Doc.ExportAsFixedFormat(
        $OutputPath, 17, $false, 0, 0, 1, 1, 0,
        $true, $false, 1, $true, $true, $false
    )
    return [ordered]@{ engine = 'Word ExportAsFixedFormat classic'; optimize_for_image_quality = $false; errors = $errors }
}

function Convert-Word([string]$SourcePath, [string]$OutputPath) {
    $word = $null
    $doc = $null
    try {
        Write-Log 'INFO' 'Starting Microsoft Word.'
        $word = New-Object -ComObject Word.Application
        $word.Visible = $false
        $word.DisplayAlerts = 0
        Disable-OfficeMacros $word
        $version = $null; $build = $null
        try { $version = [string]$word.Version } catch { }
        try { $build = [string]$word.Build } catch { }
        Write-Log 'INFO' ("Word version={0}, build={1}" -f $version, $build)

        $doc = $word.Documents.Open($SourcePath, $false, $true, $false)
        try { [void]$doc.Repaginate() } catch { }
        try { $doc.Activate() } catch { }

        $adobe = Invoke-PdfMakerExport $word $OutputPath 'Word'
        if ($null -ne $adobe) {
            $adobe.office_version = $version
            $adobe.office_build = $build
            return $adobe
        }

        if (Test-Path -LiteralPath $OutputPath) {
            try { Remove-Item -LiteralPath $OutputPath -Force } catch { }
        }
        $native = Invoke-WordNativeExport $doc $OutputPath
        $native.office_version = $version
        $native.office_build = $build
        return $native
    }
    finally {
        if ($null -ne $doc) { try { $doc.Close(0) } catch { } }
        Release-ComObject $doc
        if ($null -ne $word) { try { $word.Quit() } catch { } }
        Release-ComObject $word
        Force-ComCleanup
    }
}

function Convert-Excel([string]$SourcePath, [string]$OutputPath) {
    $excel = $null
    $workbook = $null
    try {
        Write-Log 'INFO' 'Starting Microsoft Excel.'
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        try { $excel.AskToUpdateLinks = $false } catch { }
        Disable-OfficeMacros $excel
        $version = $null; $build = $null
        try { $version = [string]$excel.Version } catch { }
        try { $build = [string]$excel.Build } catch { }
        Write-Log 'INFO' ("Excel version={0}, build={1}" -f $version, $build)

        $workbook = $excel.Workbooks.Open($SourcePath, 0, $true)
        try { $workbook.Activate() } catch { }

        $adobe = Invoke-PdfMakerExport $excel $OutputPath 'Excel'
        if ($null -ne $adobe) {
            $adobe.office_version = $version
            $adobe.office_build = $build
            return $adobe
        }

        if (Test-Path -LiteralPath $OutputPath) {
            try { Remove-Item -LiteralPath $OutputPath -Force } catch { }
        }
        Write-Log 'WARN' 'Adobe PDFMaker unavailable/failed; using Excel ExportAsFixedFormat at xlQualityStandard.'
        [void]$workbook.ExportAsFixedFormat(0, $OutputPath, 0, $true, $true)
        return [ordered]@{
            engine = 'Excel Workbook.ExportAsFixedFormat'
            office_version = $version
            office_build = $build
            quality = 'xlQualityStandard'
            ignore_print_areas = $true
        }
    }
    finally {
        if ($null -ne $workbook) { try { $workbook.Close($false) } catch { } }
        Release-ComObject $workbook
        if ($null -ne $excel) { try { $excel.Quit() } catch { } }
        Release-ComObject $excel
        Force-ComCleanup
    }
}

function Convert-PowerPoint([string]$SourcePath, [string]$OutputPath) {
    $ppt = $null
    $presentation = $null
    try {
        Write-Log 'INFO' 'Starting Microsoft PowerPoint.'
        $ppt = New-Object -ComObject PowerPoint.Application
        Disable-OfficeMacros $ppt
        $version = $null; $build = $null
        try { $version = [string]$ppt.Version } catch { }
        try { $build = [string]$ppt.Build } catch { }
        Write-Log 'INFO' ("PowerPoint version={0}, build={1}" -f $version, $build)

        # For PDFMaker, a presentation window is more reliable than a fully hidden presentation.
        $presentation = $ppt.Presentations.Open($SourcePath, -1, 0, -1)
        try { $presentation.Windows.Item(1).Activate() } catch { }

        $adobe = Invoke-PdfMakerExport $ppt $OutputPath 'PowerPoint'
        if ($null -ne $adobe) {
            $adobe.office_version = $version
            $adobe.office_build = $build
            return $adobe
        }

        if (Test-Path -LiteralPath $OutputPath) {
            try { Remove-Item -LiteralPath $OutputPath -Force } catch { }
        }
        Write-Log 'WARN' 'Adobe PDFMaker unavailable/failed; using PowerPoint ExportAsFixedFormat with Print intent.'
        try {
            [void]$presentation.ExportAsFixedFormat(
                $OutputPath, 2, 2, 0, 1, 1, 0,
                $null, 1, '', $true, $true, $true, $true, $false
            )
            return [ordered]@{
                engine = 'PowerPoint ExportAsFixedFormat'
                office_version = $version
                office_build = $build
                intent = 'Print'
            }
        } catch {
            Write-Log 'WARN' 'PowerPoint ExportAsFixedFormat failed; using SaveAs PDF.' $_.Exception
            [void]$presentation.SaveAs($OutputPath, 32)
            return [ordered]@{
                engine = 'PowerPoint SaveAs PDF fallback'
                office_version = $version
                office_build = $build
            }
        }
    }
    finally {
        if ($null -ne $presentation) { try { $presentation.Close() } catch { } }
        Release-ComObject $presentation
        if ($null -ne $ppt) { try { $ppt.Quit() } catch { } }
        Release-ComObject $ppt
        Force-ComCleanup
    }
}

function Save-Diagnostics([string]$SourcePath, [string]$OutputPath, $EngineMeta, [double]$DurationSeconds, [string]$Status, [string]$ErrorText) {
    $source = Get-Item -LiteralPath $SourcePath
    $outputSize = $null
    if (Test-Path -LiteralPath $OutputPath) {
        try { $outputSize = (Get-Item -LiteralPath $OutputPath).Length } catch { }
    }
    $record = [ordered]@{
        schema_version = 2
        timestamp_local = (Get-Date).ToString('o')
        status = $Status
        source = [ordered]@{
            file_name = $source.Name
            extension = $source.Extension.ToLowerInvariant()
            size_bytes = $source.Length
            full_local_path_recorded = $false
        }
        output = [ordered]@{
            file_name = [System.IO.Path]::GetFileName($OutputPath)
            size_bytes = $outputSize
        }
        engine = $EngineMeta
        conversion = [ordered]@{
            duration_seconds = [math]::Round($DurationSeconds, 3)
            preferred_engine = 'Adobe Acrobat PDFMaker'
            success = ($Status -eq 'success')
            error = $ErrorText
            log_file = [System.IO.Path]::GetFileName($Script:LogPath)
        }
    }
    $safe = Get-SafeFileName ([System.IO.Path]::GetFileNameWithoutExtension($SourcePath))
    $diagPath = Join-Path $DiagnosticsDirectory ("{0}_{1}_v2.json" -f (Get-Date -Format 'yyyyMMdd_HHmmss_fff'), $safe)
    $record | ConvertTo-Json -Depth 12 | Set-Content -LiteralPath $diagPath -Encoding UTF8
    Write-Log 'INFO' ("Diagnostics written: {0}" -f [System.IO.Path]::GetFileName($diagPath))
}

function Convert-OneFile([string]$InputPath) {
    $start = Get-Date
    $resolved = (Resolve-Path -LiteralPath $InputPath).Path
    $item = Get-Item -LiteralPath $resolved
    if ($item.PSIsContainer) { throw 'Folders are not supported.' }
    $ext = $item.Extension.ToLowerInvariant()
    if ($SupportedExtensions -notcontains $ext) { throw ("Unsupported file type: {0}" -f $ext) }

    $outputPath = Get-OutputPdfPath $resolved
    Write-Log 'INFO' ('=' * 70)
    Write-Log 'INFO' ("Source file: {0} ({1:N3} MB)" -f $item.Name, ($item.Length / 1MB))
    Write-Log 'INFO' ("Output file: {0}" -f [System.IO.Path]::GetFileName($outputPath))
    Write-Log 'INFO' 'Preferred path: Adobe Acrobat PDFMaker -> Office native fallback.'

    $engineMeta = $null
    try {
        switch ($ext) {
            '.docx' { $engineMeta = Convert-Word $resolved $outputPath }
            '.doc'  { $engineMeta = Convert-Word $resolved $outputPath }
            '.xlsx' { $engineMeta = Convert-Excel $resolved $outputPath }
            '.xls'  { $engineMeta = Convert-Excel $resolved $outputPath }
            '.pptx' { $engineMeta = Convert-PowerPoint $resolved $outputPath }
            '.ppt'  { $engineMeta = Convert-PowerPoint $resolved $outputPath }
        }

        if (-not (Test-Path -LiteralPath $outputPath)) {
            throw 'Conversion returned without error, but no PDF file was created.'
        }
        $pdf = Get-Item -LiteralPath $outputPath
        if ($pdf.Length -le 0) { throw 'Generated PDF is empty.' }
        $duration = ((Get-Date) - $start).TotalSeconds
        Write-Log 'OK' ("Created {0}: {1:N3} MB in {2:N2} s using {3}" -f $pdf.Name, ($pdf.Length / 1MB), $duration, $engineMeta.engine)
        Save-Diagnostics $resolved $outputPath $engineMeta $duration 'success' $null
        return $true
    } catch {
        $duration = ((Get-Date) - $start).TotalSeconds
        $errorText = Get-ExceptionText $_.Exception
        Write-Log 'ERROR' ("Conversion failed for {0}" -f $item.Name) $_.Exception
        try { Save-Diagnostics $resolved $outputPath $engineMeta $duration 'failure' $errorText } catch { }
        return $false
    }
}

Write-Log 'INFO' 'Office HQ PDF Converter v2 started.'
Write-Log 'INFO' ("PowerShell={0}; OS={1}; Process64Bit={2}" -f $PSVersionTable.PSVersion, [Environment]::OSVersion.VersionString, [Environment]::Is64BitProcess)
Write-Log 'INFO' 'Quality strategy: Acrobat PDFMaker first; Word high-image-quality API second; classic Office export last.'

if ($null -eq $InputFiles -or $InputFiles.Count -eq 0) {
    Write-Log 'ERROR' 'No input files. Drag Office files onto DragDrop_Office_to_PDF.cmd.'
    exit 1
}

$success = 0
$failed = 0
foreach ($file in $InputFiles) {
    if (Convert-OneFile $file) { $success++ } else { $failed++ }
}

Write-Log 'INFO' ('=' * 70)
Write-Log 'INFO' ("Finished. Success={0}; Failed={1}" -f $success, $failed)
Write-Log 'INFO' ("Session log: {0}" -f [System.IO.Path]::GetFileName($Script:LogPath))

if ($failed -gt 0) { exit 2 }
exit 0
