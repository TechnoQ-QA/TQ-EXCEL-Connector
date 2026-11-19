# TQ Connector for NetSuite - installs the connector for the current Windows user.
# Run by TQConnectorSetup.exe (or by hand:  powershell -ExecutionPolicy Bypass -File install.ps1).
#
#   1. Copies the connector to %LOCALAPPDATA%\TQ Connector (an update keeps data\: settings and licence).
#   2. Makes this PC's own certificate for https://localhost and trusts it for this user
#      (Windows asks you to confirm; the certificate can only vouch for localhost on this PC).
#   3. Registers the add-in with desktop Excel.
#   4. Starts the connector now and at every sign-in.
#   5. Opens the setup page for the NetSuite details and the licence.
# Nothing needs administrator rights.

param([string]$Payload = (Join-Path $PSScriptRoot 'payload.zip'))

$ErrorActionPreference = 'Stop'
$app = Join-Path $env:LOCALAPPDATA 'TQ Connector'
$exe = Join-Path $app 'TQConnector.exe'
$data = Join-Path $app 'data'
$crt = Join-Path $data 'certs\localhost.crt'
$manifest = Join-Path $app 'addin\manifest.xml'
$devKey = 'HKCU:\Software\Microsoft\Office\16.0\WEF\Developer'
$startup = Join-Path ([Environment]::GetFolderPath('Startup')) 'TQ Connector.lnk'
$uninstallKey = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\TQConnector'

function Step($text) { Write-Host "-> $text" }

try {
    Step 'Stopping a running connector (if any)'
    Get-Process TQConnector -ErrorAction SilentlyContinue | Stop-Process -Force
    Start-Sleep -Milliseconds 500

    Step "Copying the connector to $app"
    if (Test-Path $app) {
        Get-ChildItem $app -Force | Where-Object { $_.Name -ne 'data' } | Remove-Item -Recurse -Force
    }
    New-Item -ItemType Directory -Force $app | Out-Null
    Expand-Archive -Path $Payload -DestinationPath $app -Force
    New-Item -ItemType Directory -Force $data | Out-Null

    if (-not (Test-Path $crt)) {
        Step 'Making this PC''s certificate for https://localhost'
        $p = Start-Process -FilePath $exe -ArgumentList '--make-cert' -Wait -PassThru -WindowStyle Hidden
        if ($p.ExitCode -ne 0 -or -not (Test-Path $crt)) { throw 'The certificate could not be made. See data\server.log.' }
    }
    $cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2 $crt
    $trusted = Get-ChildItem Cert:\CurrentUser\Root | Where-Object Thumbprint -eq $cert.Thumbprint
    if (-not $trusted) {
        Step 'Trusting the certificate for this Windows user (please answer Yes in the Windows window)'
        Get-ChildItem Cert:\CurrentUser\Root | Where-Object { $_.Subject -like '*TQ Connector (this PC only)*' } | Remove-Item
        Import-Certificate -FilePath $crt -CertStoreLocation Cert:\CurrentUser\Root | Out-Null
    }

    Step 'Registering the add-in with Excel'
    $addinId = ([xml](Get-Content $manifest -Raw)).OfficeApp.Id
    New-Item -Path $devKey -Force | Out-Null
    Set-ItemProperty -Path $devKey -Name $addinId -Value $manifest

    Step 'Starting the connector at every sign-in'
    $shell = New-Object -ComObject WScript.Shell
    $link = $shell.CreateShortcut($startup)
    $link.TargetPath = $exe
    $link.WorkingDirectory = $app
    $link.Description = 'TQ Connector for NetSuite (local server for Excel)'
    $link.Save()

    $version = (Get-Content (Join-Path $app 'VERSION.txt') -ErrorAction SilentlyContinue | Select-Object -First 1)
    New-Item -Path $uninstallKey -Force | Out-Null
    Set-ItemProperty -Path $uninstallKey -Name DisplayName -Value 'TQ Connector for NetSuite'
    Set-ItemProperty -Path $uninstallKey -Name DisplayVersion -Value "$version"
    Set-ItemProperty -Path $uninstallKey -Name Publisher -Value 'Techno Q'
    Set-ItemProperty -Path $uninstallKey -Name InstallLocation -Value $app
    Set-ItemProperty -Path $uninstallKey -Name NoModify -Value 1 -Type DWord
    Set-ItemProperty -Path $uninstallKey -Name NoRepair -Value 1 -Type DWord
    Set-ItemProperty -Path $uninstallKey -Name UninstallString -Value ("powershell.exe -NoProfile -ExecutionPolicy Bypass -File `"" + (Join-Path $app 'uninstall.ps1') + "`"")

    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    Step 'Starting the connector'
    Start-Process -FilePath $exe -WorkingDirectory $app -WindowStyle Hidden
    $ok = $false
    for ($n = 0; $n -lt 20 -and -not $ok; $n++) {
        Start-Sleep -Milliseconds 500
        try { Invoke-WebRequest -Uri 'https://localhost:3000/api/status' -UseBasicParsing -TimeoutSec 2 | Out-Null; $ok = $true } catch { }
    }
    if (-not $ok) { Write-Warning 'The connector did not answer yet. See data\server.log in the install folder.' }

    Write-Host ''
    if (Test-Path (Join-Path $data 'config.json')) {
        # An update over a set-up PC: the settings, licence and seat are kept; no need to show the setup page.
        Write-Host "TQ Connector is updated to version $version." -ForegroundColor Green
        Write-Host 'Close and reopen the TQ Connector panel in Excel (or restart Excel) to use it.'
    } else {
        Write-Host 'TQ Connector is installed.' -ForegroundColor Green
        Write-Host '1. On the setup page that opens now, enter the NetSuite details and paste the licence.'
        Write-Host '2. Close Excel completely, open it again, then Home > Add-ins: TQ Connector is under Developer Add-ins.'
        Start-Process 'https://localhost:3000/setup.html'
    }
}
catch {
    Write-Host ''
    Write-Host ('Installation stopped: ' + $_.Exception.Message) -ForegroundColor Red
}
Write-Host ''
Read-Host 'Press Enter to close'
