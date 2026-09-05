[CmdletBinding()]
param()

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Continue'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$ScriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$RepoRoot = Split-Path -Parent $ScriptRoot
$ReportsDir = Join-Path $ScriptRoot 'reports'
New-Item -ItemType Directory -Path $ReportsDir -Force | Out-Null

$Stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$JsonPath = Join-Path $ReportsDir ("{0}_environment.json" -f $Stamp)
$TextPath = Join-Path $ReportsDir ("{0}_environment.txt" -f $Stamp)
$UserProfile = [Environment]::GetFolderPath('UserProfile')
$UserName = [Environment]::UserName
$MachineName = [Environment]::MachineName

function Sanitize-Text([object]$Value) {
    if ($null -eq $Value) { return $null }
    $s = [string]$Value
    if ($UserProfile) { $s = $s.Replace($UserProfile, '%USERPROFILE%') }
    if ($UserName) { $s = $s.Replace($UserName, '%USERNAME%') }
    if ($MachineName) { $s = $s.Replace($MachineName, '%COMPUTERNAME%') }
    $s = [regex]::Replace($s, '(?i)(https?://)([^/@\s]+)@', '$1***@')
    $s = [regex]::Replace($s, '(?i)(token|password|passwd|secret|api[_-]?key)\s*[=:]\s*[^\s;]+', '$1=***')
    return $s
}

function Invoke-CapturedCommand([string]$Exe, [string[]]$Arguments) {
    $cmd = Get-Command $Exe -ErrorAction SilentlyContinue
    if ($null -eq $cmd) {
        return [ordered]@{ found = $false; executable = $null; exit_code = $null; output = $null }
    }
    try {
        $output = (& $cmd.Source @Arguments 2>&1 | Out-String).Trim()
        $exitCode = $LASTEXITCODE
    } catch {
        $output = $_.Exception.Message
        $exitCode = -1
    }
    return [ordered]@{
        found = $true
        executable = Sanitize-Text $cmd.Source
        exit_code = $exitCode
        output = Sanitize-Text $output
    }
}

function Invoke-PythonProbe([string]$PythonExe, [string[]]$Prefix, [string]$Code) {
    $cmd = Get-Command $PythonExe -ErrorAction SilentlyContinue
    if ($null -eq $cmd) {
        return [ordered]@{ found = $false; executable = $null; exit_code = $null; output = $null }
    }

    $probePath = Join-Path $env:TEMP ("chatgpt_env_probe_{0}.py" -f [guid]::NewGuid().ToString('N'))
    try {
        [IO.File]::WriteAllText($probePath, $Code, (New-Object System.Text.UTF8Encoding($false)))
        $args = @()
        $args += $Prefix
        $args += $probePath
        $output = (& $cmd.Source @args 2>&1 | Out-String).Trim()
        $exitCode = $LASTEXITCODE
        return [ordered]@{
            found = $true
            executable = Sanitize-Text $cmd.Source
            exit_code = $exitCode
            output = Sanitize-Text $output
        }
    } catch {
        return [ordered]@{
            found = $true
            executable = Sanitize-Text $cmd.Source
            exit_code = -1
            output = Sanitize-Text $_.Exception.Message
        }
    } finally {
        if (Test-Path -LiteralPath $probePath) { try { Remove-Item -LiteralPath $probePath -Force } catch { } }
    }
}

function Get-OfficeComInfo([string]$ProgId, [string]$Name) {
    $app = $null
    try {
        $app = New-Object -ComObject $ProgId
        $version = $null; $build = $null
        try { $version = [string]$app.Version } catch { }
        try { $build = [string]$app.Build } catch { }
        return [ordered]@{ name = $Name; available = $true; version = $version; build = $build; error = $null }
    } catch {
        return [ordered]@{ name = $Name; available = $false; version = $null; build = $null; error = Sanitize-Text $_.Exception.Message }
    } finally {
        if ($null -ne $app) {
            try { $app.Quit() } catch { }
            try {
                if ([Runtime.InteropServices.Marshal]::IsComObject($app)) {
                    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)
                }
            } catch { }
        }
        [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    }
}

function Test-RegistryPath([string]$Path) {
    try { return (Test-Path -LiteralPath $Path) } catch { return $false }
}

function Get-InstalledAdobeProducts {
    $items = @()
    $roots = @(
        'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*',
        'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*',
        'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*'
    )
    foreach ($root in $roots) {
        try {
            foreach ($item in (Get-ItemProperty $root -ErrorAction SilentlyContinue)) {
                if ($item.DisplayName -and $item.DisplayName -match '(?i)Adobe Acrobat|Adobe Reader') {
                    $items += [ordered]@{
                        name = Sanitize-Text $item.DisplayName
                        version = Sanitize-Text $item.DisplayVersion
                    }
                }
            }
        } catch { }
    }
    return @($items | Sort-Object { $_.name + $_.version } -Unique)
}

$os = $null
try { $os = Get-CimInstance Win32_OperatingSystem } catch { }
$cpu = $null
try { $cpu = Get-CimInstance Win32_Processor | Select-Object -First 1 } catch { }
$gpus = @()
try {
    $gpus = @(Get-CimInstance Win32_VideoController | ForEach-Object {
        [ordered]@{
            name = Sanitize-Text $_.Name
            driver_version = Sanitize-Text $_.DriverVersion
            current_resolution = if ($_.CurrentHorizontalResolution -and $_.CurrentVerticalResolution) { "{0}x{1}" -f $_.CurrentHorizontalResolution, $_.CurrentVerticalResolution } else { $null }
        }
    })
} catch { }

$displays = @()
try {
    Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop
    $displays = @([System.Windows.Forms.Screen]::AllScreens | ForEach-Object {
        [ordered]@{
            device_name = Sanitize-Text $_.DeviceName
            primary = $_.Primary
            bounds = "{0},{1} {2}x{3}" -f $_.Bounds.X, $_.Bounds.Y, $_.Bounds.Width, $_.Bounds.Height
            working_area = "{0},{1} {2}x{3}" -f $_.WorkingArea.X, $_.WorkingArea.Y, $_.WorkingArea.Width, $_.WorkingArea.Height
        }
    })
} catch { }

$dpi = [ordered]@{ log_pixels = $null; win8_dpi_scaling = $null }
try {
    $desktop = Get-ItemProperty 'HKCU:\Control Panel\Desktop' -ErrorAction SilentlyContinue
    if ($null -ne $desktop) {
        $dpi.log_pixels = $desktop.LogPixels
        $dpi.win8_dpi_scaling = $desktop.Win8DpiScaling
    }
} catch { }

$pythonLauncher = Invoke-CapturedCommand 'py.exe' @('-0p')
$pythonVersion = Invoke-CapturedCommand 'py.exe' @('-3', '--version')
$pipVersion = Invoke-CapturedCommand 'py.exe' @('-3', '-m', 'pip', '--version')
$pipConfig = Invoke-CapturedCommand 'py.exe' @('-3', '-m', 'pip', 'config', 'list')
$pipPyMuPDF = Invoke-CapturedCommand 'py.exe' @('-3', '-m', 'pip', 'show', 'pymupdf')
$pipPillow = Invoke-CapturedCommand 'py.exe' @('-3', '-m', 'pip', 'show', 'pillow')

$pythonProbeCode = @'
import importlib.util
import json
import platform
import sys
mods = ['pymupdf', 'fitz', 'PIL', 'numpy', 'scipy']
print(json.dumps({
    'version': sys.version,
    'executable': sys.executable,
    'architecture': platform.architecture()[0],
    'implementation': platform.python_implementation(),
    'modules': {m: (importlib.util.find_spec(m) is not None) for m in mods},
}, ensure_ascii=False))
'@
$pythonRuntimeProbe = Invoke-PythonProbe 'py.exe' @('-3') $pythonProbeCode

$gitVersion = Invoke-CapturedCommand 'git.exe' @('--version')
$gitBranch = Invoke-CapturedCommand 'git.exe' @('-C', $RepoRoot, 'branch', '--show-current')
$gitCommit = Invoke-CapturedCommand 'git.exe' @('-C', $RepoRoot, 'rev-parse', '--short', 'HEAD')
$gitOrigin = Invoke-CapturedCommand 'git.exe' @('-C', $RepoRoot, 'remote', 'get-url', 'origin')
$gitStatus = Invoke-CapturedCommand 'git.exe' @('-C', $RepoRoot, 'status', '--porcelain=v1', '-uno')

$nodeVersion = Invoke-CapturedCommand 'node.exe' @('--version')
$npmVersion = Invoke-CapturedCommand 'npm.cmd' @('--version')
$npmRegistry = Invoke-CapturedCommand 'npm.cmd' @('config', 'get', 'registry')
$dotnetInfo = Invoke-CapturedCommand 'dotnet.exe' @('--info')
$dotnetSdks = Invoke-CapturedCommand 'dotnet.exe' @('--list-sdks')
$dotnetRuntimes = Invoke-CapturedCommand 'dotnet.exe' @('--list-runtimes')
$javaVersion = Invoke-CapturedCommand 'java.exe' @('-version')

$office = [ordered]@{
    word = Get-OfficeComInfo 'Word.Application' 'Microsoft Word'
    excel = Get-OfficeComInfo 'Excel.Application' 'Microsoft Excel'
    powerpoint = Get-OfficeComInfo 'PowerPoint.Application' 'Microsoft PowerPoint'
}

$pdfMakerRegistered = (
    (Test-RegistryPath 'Registry::HKEY_CLASSES_ROOT\PDFMaker.OfficeAddin') -or
    (Test-RegistryPath 'Registry::HKEY_CLASSES_ROOT\PDFMaker.OfficeAddin\CLSID')
)

$memoryGb = $null
if ($null -ne $os -and $os.TotalVisibleMemorySize) {
    $memoryGb = [math]::Round(($os.TotalVisibleMemorySize * 1KB / 1GB), 2)
}

$report = [ordered]@{
    schema_version = 2
    generated_at = (Get-Date).ToString('o')
    privacy = [ordered]@{
        username_recorded = $false
        computer_name_recorded = $false
        user_profile_replaced_with_placeholder = $true
        arbitrary_environment_variables_recorded = $false
        url_credentials_masked = $true
    }
    windows = [ordered]@{
        caption = if ($null -ne $os) { Sanitize-Text $os.Caption } else { $null }
        version = if ($null -ne $os) { Sanitize-Text $os.Version } else { [Environment]::OSVersion.Version.ToString() }
        build = if ($null -ne $os) { Sanitize-Text $os.BuildNumber } else { $null }
        architecture = if ($null -ne $os) { Sanitize-Text $os.OSArchitecture } else { $null }
        culture = [Globalization.CultureInfo]::CurrentCulture.Name
        ui_culture = [Globalization.CultureInfo]::CurrentUICulture.Name
        timezone = [TimeZoneInfo]::Local.Id
        total_memory_gb = $memoryGb
    }
    hardware = [ordered]@{
        cpu = if ($null -ne $cpu) { [ordered]@{ name = Sanitize-Text $cpu.Name; cores = $cpu.NumberOfCores; logical_processors = $cpu.NumberOfLogicalProcessors } } else { $null }
        gpu = $gpus
        displays = $displays
        dpi = $dpi
    }
    powershell = [ordered]@{
        version = $PSVersionTable.PSVersion.ToString()
        edition = $PSVersionTable.PSEdition
        process_64bit = [Environment]::Is64BitProcess
        execution_policy = Sanitize-Text ((Get-ExecutionPolicy -List | Format-Table -AutoSize | Out-String).Trim())
    }
    python = [ordered]@{
        launcher_versions = $pythonLauncher
        py3_version = $pythonVersion
        runtime_probe = $pythonRuntimeProbe
        pip_version = $pipVersion
        pip_config = $pipConfig
        pymupdf = $pipPyMuPDF
        pillow = $pipPillow
        environment_index_url = Sanitize-Text $env:PIP_INDEX_URL
        environment_extra_index_url = Sanitize-Text $env:PIP_EXTRA_INDEX_URL
    }
    office = $office
    adobe = [ordered]@{
        pdfmaker_office_addin_registered = $pdfMakerRegistered
        installed_products = Get-InstalledAdobeProducts
    }
    developer_tools = [ordered]@{
        git = $gitVersion
        node = $nodeVersion
        npm = $npmVersion
        npm_registry = $npmRegistry
        dotnet_info = $dotnetInfo
        dotnet_sdks = $dotnetSdks
        dotnet_runtimes = $dotnetRuntimes
        java = $javaVersion
    }
    repository = [ordered]@{
        branch = $gitBranch
        commit = $gitCommit
        origin = $gitOrigin
        local_changes_without_untracked_files = $gitStatus
    }
}

$report | ConvertTo-Json -Depth 15 | Set-Content -LiteralPath $JsonPath -Encoding UTF8

$summary = @"
Local Environment Snapshot
==========================
Generated: $($report.generated_at)
Schema: $($report.schema_version)

Windows
-------
$($report.windows.caption) / version $($report.windows.version) / build $($report.windows.build) / $($report.windows.architecture)
CPU: $($report.hardware.cpu.name)
RAM: $($report.windows.total_memory_gb) GB
PowerShell: $($report.powershell.version) ($($report.powershell.edition))

Python
------
py -0p:
$($report.python.launcher_versions.output)

py -3 --version:
$($report.python.py3_version.output)

Runtime/import probe:
$($report.python.runtime_probe.output)

pip:
$($report.python.pip_version.output)

pip config:
$($report.python.pip_config.output)

PyMuPDF:
$($report.python.pymupdf.output)

Pillow:
$($report.python.pillow.output)

Office
------
Word: available=$($report.office.word.available), version=$($report.office.word.version), build=$($report.office.word.build)
Excel: available=$($report.office.excel.available), version=$($report.office.excel.version), build=$($report.office.excel.build)
PowerPoint: available=$($report.office.powerpoint.available), version=$($report.office.powerpoint.version), build=$($report.office.powerpoint.build)
PDFMaker registered: $($report.adobe.pdfmaker_office_addin_registered)

Developer tools
---------------
Git: $($report.developer_tools.git.output)
Node: $($report.developer_tools.node.output)
npm: $($report.developer_tools.npm.output)
npm registry: $($report.developer_tools.npm_registry.output)
.NET SDKs:
$($report.developer_tools.dotnet_sdks.output)
.NET runtimes:
$($report.developer_tools.dotnet_runtimes.output)
Java: $($report.developer_tools.java.output)

Repository
----------
Branch: $($report.repository.branch.output)
Commit: $($report.repository.commit.output)
Origin: $($report.repository.origin.output)

Privacy
-------
Username/computer name are not intentionally recorded. User-profile paths are replaced with %USERPROFILE% and URL credentials are masked.
"@

$summary | Set-Content -LiteralPath $TextPath -Encoding UTF8

Write-Host ''
Write-Host 'Environment snapshot created:' -ForegroundColor Green
Write-Host ("  {0}" -f $JsonPath)
Write-Host ("  {0}" -f $TextPath)
Write-Host ''
Write-Host 'Commit and Push the files in Local_Environment_Collector\reports, then ask ChatGPT to read the latest report.' -ForegroundColor Cyan
