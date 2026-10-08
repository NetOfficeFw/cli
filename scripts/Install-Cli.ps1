#requires -Version 7.0
<#
.SYNOPSIS
Builds and registers the 64-bit NetOffice PowerPoint add-in and puts `netoffice` on PATH.

.DESCRIPTION
Run from a checkout of this repository, with PowerPoint closed:
  1. verifies the prerequisites (64-bit Office, Node.js 20+, Visual Studio MSBuild, vcpkg),
  2. builds src/addin as Release x64,
  3. registers the add-in for the current user (HKCU, no elevation),
  4. installs the CLI dependencies and links `netoffice` globally (npm link),
  5. launches PowerPoint once to prove the add-in answers, then shuts it down.
Safe to repeat; each run rebuilds, re-registers, and re-links to this checkout.

.PARAMETER Port
Loopback port the add-in listens on, stored in HKCU. Default 50051.

.PARAMETER OfficeExtensibilityDir
Directory containing MSADDNDR.OLB. Default: the 64-bit Click-to-Run DESIGNER directory.

.PARAMETER SkipVerify
Skip step 5 (the PowerPoint launch/shutdown check).
#>
[CmdletBinding()]
param(
    [ValidateRange(1, 65535)]
    [int] $Port = 50051,

    [string] $OfficeExtensibilityDir = (Join-Path $env:ProgramFiles 'Microsoft Office\root\vfs\ProgramFilesCommonX64\DESIGNER'),

    [switch] $SkipVerify
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
if (-not $IsWindows) { throw 'NetOffice requires Windows.' }

$repo = Split-Path -Parent $PSScriptRoot
$projectFile = Join-Path $repo 'src\addin\addin.vcxproj'
$dll = Join-Path $repo 'src\addin\build\Release_x64\addin.dll'
$register = Join-Path $repo 'src\addin\Register-Addin.ps1'
$cliDir = Join-Path $repo 'src\cli'

function Step([string] $message) { Write-Host "==> $message" -ForegroundColor Cyan }
function Invoke-Native([string] $failure, [scriptblock] $command) {
    & $command
    if ($LASTEXITCODE -ne 0) { throw "$failure (exit code $LASTEXITCODE)." }
}

Step 'Checking prerequisites'
if (-not (Test-Path -LiteralPath $projectFile)) { throw "Not a NetOffice checkout: $projectFile is missing." }

$office = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration' -ErrorAction SilentlyContinue
if (-not $office -or $office.Platform -ne 'x64') {
    throw 'A 64-bit Click-to-Run Microsoft Office installation is required.'
}
if (-not (Test-Path -LiteralPath (Join-Path $OfficeExtensibilityDir 'MSADDNDR.OLB'))) {
    throw "MSADDNDR.OLB not found in '$OfficeExtensibilityDir'. Pass -OfficeExtensibilityDir."
}
if (Get-Process POWERPNT -ErrorAction SilentlyContinue) {
    throw 'PowerPoint is running. Save and close it, then run this script again.'
}

foreach ($tool in 'node', 'npm') {
    if (-not (Get-Command $tool -ErrorAction SilentlyContinue)) { throw "$tool was not found on PATH. Install Node.js 20 or newer." }
}
if ([version](node -p 'process.versions.node') -lt [version]'20.0.0') { throw 'Node.js 20 or newer is required.' }

$vswhere = Join-Path ${env:ProgramFiles(x86)} 'Microsoft Visual Studio\Installer\vswhere.exe'
$msbuild = if (Test-Path -LiteralPath $vswhere) {
    & $vswhere -latest -products * -requires Microsoft.Component.MSBuild -find 'MSBuild\**\Bin\amd64\MSBuild.exe' |
        Select-Object -First 1
}
if (-not $msbuild) { throw 'MSBuild not found. Install Visual Studio with the Desktop C++ workload (v145 toolset, Windows SDK, ATL).' }

# vcpkg is not pre-checked: MSBuild resolves it through Visual Studio's integration,
# which works without vcpkg on PATH. A missing integration fails in the build step.

Step 'Building the add-in (Release x64)'
Invoke-Native 'Add-in build failed' {
    & $msbuild $projectFile /p:Configuration=Release /p:Platform=x64 "/p:OfficeExtensibilityDir=$OfficeExtensibilityDir" /v:minimal /nologo
}
if (-not (Test-Path -LiteralPath $dll)) { throw "Build reported success but $dll is missing." }

Step "Registering the add-in for the current user (port $Port)"
& $register -Action Register -Architecture x64 -DllPath $dll -Port $Port

Step 'Installing CLI dependencies'
Invoke-Native 'npm ci failed' { npm --prefix $cliDir ci }

Step 'Linking netoffice onto PATH'
$globalRoot = (npm root -g).Trim()
$link = Join-Path $globalRoot 'netoffice'
$previous = (Get-Item -LiteralPath $link -ErrorAction SilentlyContinue)?.Target
if ($previous -and ([IO.Path]::GetFullPath($previous) -ne [IO.Path]::GetFullPath($cliDir))) {
    Write-Host "    Re-pointing the global netoffice link from $previous" -ForegroundColor Yellow
}
Push-Location $cliDir
try { Invoke-Native 'npm link failed' { npm link } } finally { Pop-Location }

$globalBin = (npm prefix -g).Trim()
$command = Get-Command netoffice -ErrorAction SilentlyContinue
if (-not $command) {
    throw "netoffice is linked in '$globalBin' but that directory is not on PATH. Add it to PATH and open a new terminal."
}
# A different netoffice earlier on PATH (another checkout, a stale shim) would run the wrong copy.
if ((Split-Path -Parent $command.Source).TrimEnd('\') -ne $globalBin.TrimEnd('\')) {
    throw "'netoffice' on PATH resolves to $($command.Source), not the link just created in '$globalBin'. Remove or reorder the stale entry on PATH."
}
Write-Host "    $($command.Source) -> $cliDir"

if (-not $SkipVerify) {
    Step 'Verifying the add-in answers (launches then closes PowerPoint)'
    Invoke-Native 'netoffice powerpoint launch failed' { netoffice powerpoint launch --port $Port --timeout 60000 }
    try {
        Invoke-Native 'netoffice presentation list failed' { netoffice presentation list --port $Port }
    } finally {
        netoffice powerpoint shutdown --port $Port --timeout 60000
    }
}

Write-Host "Done. Try: netoffice --help" -ForegroundColor Green
