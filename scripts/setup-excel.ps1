# One-time setup so desktop Excel can load the add-in from https://localhost:3000.
#   1. Creates the localhost HTTPS certificate if it does not exist yet.
#   2. Trusts that certificate for your Windows user (Windows asks you to confirm).
#   3. Registers addin\manifest.xml as a developer add-in for Excel.
#
# Run:   powershell -ExecutionPolicy Bypass -File scripts\setup-excel.ps1
# Undo:  powershell -ExecutionPolicy Bypass -File scripts\setup-excel.ps1 -Remove

param([switch]$Remove)

$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$certDir = Join-Path $root 'certs'
$crt = Join-Path $certDir 'localhost.crt'
$key = Join-Path $certDir 'localhost.key'
$manifest = Join-Path $root 'addin\manifest.xml'
$addinId = ([xml](Get-Content $manifest -Raw)).OfficeApp.Id
$devKey = 'HKCU:\Software\Microsoft\Office\16.0\WEF\Developer'

if ($Remove) {
    if (Test-Path $devKey) { Remove-ItemProperty -Path $devKey -Name $addinId -ErrorAction SilentlyContinue }
    if (Test-Path $crt) {
        $thumb = (New-Object System.Security.Cryptography.X509Certificates.X509Certificate2 $crt).Thumbprint
        Get-ChildItem Cert:\CurrentUser\Root | Where-Object Thumbprint -eq $thumb | Remove-Item
    }
    Write-Host 'Add-in unregistered and certificate trust removed.'
    return
}

if (-not ((Test-Path $crt) -and (Test-Path $key))) {
    $openssl = @(
        "$env:ProgramFiles\Git\mingw64\bin\openssl.exe",
        "$env:ProgramFiles\Git\usr\bin\openssl.exe",
        "$env:LOCALAPPDATA\Programs\Git\mingw64\bin\openssl.exe"
    ) | Where-Object { Test-Path $_ } | Select-Object -First 1
    if (-not $openssl) {
        $found = Get-Command openssl -ErrorAction SilentlyContinue
        if ($found) { $openssl = $found.Source }
    }
    if (-not $openssl) { throw 'openssl.exe not found. It ships with Git for Windows.' }
    New-Item -ItemType Directory -Force $certDir | Out-Null
    & $openssl req -x509 -newkey rsa:2048 -nodes -days 825 `
        -keyout $key -out $crt -subj '/CN=localhost' `
        -addext 'subjectAltName=DNS:localhost,IP:127.0.0.1' 2>$null
    if ($LASTEXITCODE -ne 0) { throw 'openssl could not create the certificate.' }
    Write-Host "Certificate created: $crt"
}

Import-Certificate -FilePath $crt -CertStoreLocation Cert:\CurrentUser\Root | Out-Null
Write-Host 'Certificate trusted for this Windows user.'

New-Item -Path $devKey -Force | Out-Null
Set-ItemProperty -Path $devKey -Name $addinId -Value $manifest
Write-Host "Add-in registered for Excel: $manifest"
Write-Host 'Close Excel completely, reopen it, then look under Home > Add-ins.'
