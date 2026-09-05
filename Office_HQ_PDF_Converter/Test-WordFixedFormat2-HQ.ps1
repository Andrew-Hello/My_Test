[CmdletBinding()]
param(
    [Parameter(Mandatory = $true, Position = 0)]
    [string]$InputFile
)

$ErrorActionPreference = 'Stop'
$resolved = (Resolve-Path -LiteralPath $InputFile).Path
$dir = [IO.Path]::GetDirectoryName($resolved)
$name = [IO.Path]::GetFileNameWithoutExtension($resolved)
$out = Join-Path $dir ($name + '_FixedFormat2_HQ.pdf')

$word = $null
$doc = $null
try {
    Write-Host "Starting Word..."
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    try { $word.AutomationSecurity = 3 } catch { }
    Write-Host ("Word Version={0}, Build={1}" -f $word.Version, $word.Build)

    $doc = $word.Documents.Open($resolved, $false, $true, $false)
    try { [void]$doc.Repaginate() } catch { }

    if (Test-Path -LiteralPath $out) { Remove-Item -LiteralPath $out -Force }

    Write-Host "Calling Document.ExportAsFixedFormat2 with OptimizeForImageQuality=True..."
    Write-Host "Important: the optional FixedFormatExtClassPtr argument is intentionally omitted."

    # Microsoft-documented signature (15 arguments here; final optional pointer omitted):
    # OutputFileName, ExportFormat, OpenAfterExport, OptimizeFor, Range,
    # From, To, Item, IncludeDocProps, KeepIRM, CreateBookmarks,
    # DocStructureTags, BitmapMissingFonts, UseISO19005_1,
    # OptimizeForImageQuality
    [void]$doc.ExportAsFixedFormat2(
        $out,        # OutputFileName
        17,          # wdExportFormatPDF
        $false,      # OpenAfterExport
        0,           # wdExportOptimizeForPrint
        0,           # wdExportAllDocument
        1,           # From (ignored for whole document)
        1,           # To   (ignored for whole document)
        0,           # wdExportDocumentContent
        $true,       # IncludeDocProps
        $false,      # KeepIRM
        1,           # wdExportCreateHeadingBookmarks
        $true,       # DocStructureTags
        $true,       # BitmapMissingFonts
        $false,      # UseISO19005_1
        $true        # OptimizeForImageQuality <-- key
    )

    if (-not (Test-Path -LiteralPath $out)) {
        throw 'ExportAsFixedFormat2 returned but no PDF was created.'
    }
    $pdf = Get-Item -LiteralPath $out
    Write-Host ("SUCCESS: {0}" -f $pdf.FullName) -ForegroundColor Green
    Write-Host ("PDF size: {0:N3} MB ({1} bytes)" -f ($pdf.Length / 1MB), $pdf.Length) -ForegroundColor Green
}
catch {
    $hr = try { '0x{0:X8}' -f ($_.Exception.HResult -band 0xffffffff) } catch { 'n/a' }
    Write-Host ("FAILED: {0}" -f $_.Exception.Message) -ForegroundColor Red
    Write-Host ("HRESULT: {0}" -f $hr) -ForegroundColor Red
    exit 2
}
finally {
    if ($null -ne $doc) { try { $doc.Close(0) } catch { } }
    if ($null -ne $word) { try { $word.Quit() } catch { } }
    if ($null -ne $doc -and [Runtime.InteropServices.Marshal]::IsComObject($doc)) { try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($doc) } catch { } }
    if ($null -ne $word -and [Runtime.InteropServices.Marshal]::IsComObject($word)) { try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($word) } catch { } }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
