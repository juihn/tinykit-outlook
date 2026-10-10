# Builds tinykit.Outlook.sln with the x64 MSBuild of the latest Visual Studio (the build also registers the add-in).
# x64 MSBuild even on Windows on ARM: the registration step loads the x64 vstoee.dll, which ARM64 MSBuild cannot.
# Visual Studio on ARM has no Office/SharePoint workload, so when its OfficeTools are missing the VSTO build files are
# taken from -Kit: OfficeTools\ (from an x64 VS's MSBuild\Microsoft\VisualStudio\v18.0\OfficeTools) and VSTO40\
# (Microsoft.Office.Tools.Common/Outlook.v4.0.Utilities.dll).
param(
    [ValidateSet('Debug', 'Release')] [string] $Configuration = 'Debug',
    [string] $Kit = (Join-Path $env:LOCALAPPDATA 'tinykit\vsto-build-kit')
)
$ErrorActionPreference = 'Stop'

if (Get-Process OUTLOOK -ErrorAction SilentlyContinue) { throw 'Close Outlook first (it locks the add-in DLL).' }

$vswhere = "${env:ProgramFiles(x86)}\Microsoft Visual Studio\Installer\vswhere.exe"
if (-not (Test-Path $vswhere)) { throw "Visual Studio not found: $vswhere" }
$vs = & $vswhere -latest -products * -requires Microsoft.Component.MSBuild -property installationPath
$msbuild = Join-Path "$vs" 'MSBuild\Current\Bin\amd64\MSBuild.exe'
if (-not $vs -or -not (Test-Path $msbuild)) { throw 'x64 MSBuild of Visual Studio not found.' }

if (-not (Test-Path (Join-Path $PSScriptRoot 'Signing.props'))) { & (Join-Path $PSScriptRoot 'New-SigningCert.ps1') }

$msbuildArgs = @((Join-Path $PSScriptRoot 'tinykit.Outlook.sln'), '-nologo', '-v:minimal', "-p:Configuration=$Configuration")
if (-not (Test-Path (Join-Path $vs 'MSBuild\Microsoft\VisualStudio\v*\OfficeTools'))) {
    if (-not (Test-Path (Join-Path $Kit 'OfficeTools\Microsoft.VisualStudio.Tools.Office.targets'))) {
        throw "Visual Studio has no Office/SharePoint workload and $Kit\OfficeTools is missing."
    }
    Write-Host "Using VSTO build files from $Kit"
    $msbuildArgs += "-p:VSToolsPath=$Kit", "-p:ReferencePath=$Kit\VSTO40"
}

& $msbuild @msbuildArgs
exit $LASTEXITCODE
