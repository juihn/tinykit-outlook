# Registers (or -Unregister) the built tinykit Outlook add-in for the current user.
# Visual Studio does this on F5/Build; a command-line MSBuild build does not.
# On first Outlook start, VSTO asks once to trust the self-signed publisher — choose Install.
param(
    [ValidateSet('Debug', 'Release')] [string] $Configuration = 'Debug',
    [switch] $Unregister,
    [switch] $ShareSettings
)

$key = 'HKCU:\Software\Microsoft\Office\Outlook\Addins\tinykit.Outlook'

if ($Unregister) {
    Remove-Item $key -Recurse -ErrorAction SilentlyContinue
    Write-Host "Unregistered $key"
    return
}

$vsto = Join-Path $PSScriptRoot "tinykit.Outlook\bin\$Configuration\tinykit.Outlook.vsto"
if (-not (Test-Path $vsto)) { throw "Build first: $vsto not found" }

New-Item $key -Force | Out-Null
Set-ItemProperty $key -Name FriendlyName -Value 'tinykit Outlook'
Set-ItemProperty $key -Name Description -Value 'Outlook tools: quick and saved filters, conditional formatting, custom mail fields, view font'
Set-ItemProperty $key -Name LoadBehavior -Value 3 -Type DWord
Set-ItemProperty $key -Name Manifest -Value (([Uri]$vsto).AbsoluteUri + '|vstolocal')
Write-Host "Registered $key -> $vsto"

# Shared settings (see README "Settings location"): OneDrive syncs the .config folder but not its Hidden attribute,
# so hide it on every PC. Pass -ShareSettings to create the folder (the add-in copies the current settings there).
$oneDrive = $env:OneDriveConsumer
if (-not $oneDrive) { $oneDrive = (Get-ItemProperty 'HKCU:\Software\Microsoft\OneDrive\Accounts\Personal' -ErrorAction SilentlyContinue).UserFolder }
if ($oneDrive) {
    $config = Join-Path $oneDrive '.config'
    $shared = Join-Path $config 'tinykit\Outlook'
    if ($ShareSettings -and -not (Test-Path $shared)) {
        New-Item -ItemType Directory -Force $shared | Out-Null
        Write-Host "Created $shared"
    }
    if (Test-Path $config) {
        (Get-Item -Force $config).Attributes = 'Directory, Hidden'
        Write-Host "Settings are shared through OneDrive: $shared ($config hidden)"
    }
}
