# Creates a self-signed code-signing certificate in CurrentUser\My and writes its thumbprint to Signing.props,
# which the project imports to sign the VSTO manifests. Run once per PC before the first build.
# Signing.props is per PC and is not committed. Use -Force to replace an existing Signing.props.
param(
    [string] $Subject = 'CN=TinyKit (self-signed)',
    [switch] $Force
)

$props = Join-Path $PSScriptRoot 'Signing.props'
if ((Test-Path $props) -and -not $Force) {
    Write-Host "$props already exists (use -Force to create a new certificate)."
    return
}

$cert = New-SelfSignedCertificate -Type CodeSigningCert -Subject $Subject `
    -CertStoreLocation Cert:\CurrentUser\My -KeyExportPolicy Exportable -NotAfter (Get-Date).AddYears(10)

@"
<Project>
  <!-- Per-PC manifest signing certificate (CurrentUser\My), written by New-SigningCert.ps1. Not committed. -->
  <PropertyGroup>
    <ManifestCertificateThumbprint>$($cert.Thumbprint)</ManifestCertificateThumbprint>
  </PropertyGroup>
</Project>
"@ | Set-Content -Path $props -Encoding UTF8

Write-Host "Created $($cert.Subject) ($($cert.Thumbprint)) and $props"
