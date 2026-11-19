# Signs one file with Techno Q's code-signing certificate from the Windows certificate store (no tools to buy or
# install: PowerShell's own Set-AuthenticodeSignature). build.ps1 calls it for TQConnector.exe, and Inno Setup
# calls it for the setup file and its uninstaller.
#
#   powershell -NoProfile -ExecutionPolicy Bypass -File installer\sign.ps1 <file>
#
# The certificate is the one whose subject is "CN=Techno Q Code Signing" (made once with New-SelfSignedCertificate,
# see docs\CLIENT-INSTALL.md), or the thumbprint in TQ_SIGN_THUMBPRINT. Its public part, installer\TechnoQ.cer, is
# what a client's IT trusts. The signature is timestamped when the timestamp server answers, so files stay valid
# after the certificate expires; offline, the file is signed without a timestamp.

param([Parameter(Mandatory = $true)][string]$File)
$ErrorActionPreference = 'Stop'

$certs = Get-ChildItem Cert:\CurrentUser\My -CodeSigningCert
if ($env:TQ_SIGN_THUMBPRINT) { $cert = $certs | Where-Object Thumbprint -eq $env:TQ_SIGN_THUMBPRINT }
else { $cert = $certs | Where-Object Subject -like 'CN=Techno Q Code Signing*' | Sort-Object NotAfter -Descending }
$cert = $cert | Select-Object -First 1
if (-not $cert) { Write-Error 'No Techno Q code-signing certificate in Cert:\CurrentUser\My.'; exit 1 }

$result = $null
try {
    $result = Set-AuthenticodeSignature -FilePath $File -Certificate $cert -HashAlgorithm SHA256 `
        -TimestampServer 'http://timestamp.digicert.com'
} catch { $result = $null }
if (-not $result -or -not $result.SignerCertificate) {
    $result = Set-AuthenticodeSignature -FilePath $File -Certificate $cert -HashAlgorithm SHA256
}
# On a PC that does not trust the certificate yet the status is UnknownError, but the signature is in the file.
if (-not $result.SignerCertificate -or $result.SignerCertificate.Thumbprint -ne $cert.Thumbprint) {
    Write-Error "Signing $File failed: $($result.StatusMessage)"; exit 1
}
Write-Host "   signed $(Split-Path -Leaf $File) ($($result.Status))"
exit 0
