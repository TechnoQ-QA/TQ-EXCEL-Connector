# Techno Q only: builds build\TQConnectorSetup.exe, the one file a customer runs on each PC.
#
#   powershell -ExecutionPolicy Bypass -File installer\build.ps1
#
# 1. Runs the tests, then makes TQConnector.exe from server\ with PyInstaller, so customers need no Python.
#    PyInstaller is installed into build\venv only; the PC's own Python is not changed.
# 2. Stages the exe and the add-in files in build\stage\app.
# 3. Makes TQConnectorSetup.exe with Inno Setup (installer\TQConnector.iss).
#
# 4. Writes build\TQConnectorSetup.exe.sha256.txt (the checksum) and copies installer\TechnoQ.cer next to it.
#
# Signing: TQConnector.exe, the setup file and its uninstaller are signed with Techno Q's own certificate
# (installer\sign.ps1) when it is in this Windows user's certificate store. For a bought certificate, set
# TQ_SIGNTOOL to its signing command instead, with $f where the file goes, e.g.
#   $env:TQ_SIGNTOOL = 'signtool sign /fd sha256 /tr http://timestamp.digicert.com /td sha256 /sha1 <thumbprint> $f'
# (or the cloud signing tool's command). TQ_SIGNTOOL=none builds unsigned.

$ErrorActionPreference = 'Stop'
$PSNativeCommandUseErrorActionPreference = $false
$root = Split-Path -Parent $PSScriptRoot
$build = Join-Path $root 'build'
$venv = Join-Path $build 'venv'
$py = Join-Path $venv 'Scripts\python.exe'
$stage = Join-Path $build 'stage'
$app = Join-Path $stage 'app'
$setup = Join-Path $build 'TQConnectorSetup.exe'
$pyinstaller = 'pyinstaller==6.11.1'

New-Item -ItemType Directory -Force $build | Out-Null
if (-not (Test-Path $py)) {
    Write-Host '-> Making build\venv'
    python -m venv $venv
    if ($LASTEXITCODE -ne 0) { throw 'python -m venv failed' }
}
if (-not (Test-Path (Join-Path $venv 'Lib\site-packages\PyInstaller'))) {
    Write-Host "-> Installing $pyinstaller into build\venv"
    $ErrorActionPreference = 'Continue'
    & $py -m pip install --disable-pip-version-check --quiet $pyinstaller 2>&1 | ForEach-Object { "   $_" }
    if ($LASTEXITCODE -ne 0) { throw 'pip install pyinstaller failed' }
    $ErrorActionPreference = 'Stop'
}

Write-Host '-> Running the tests'
$ErrorActionPreference = 'Continue'
$tests = python -m unittest discover -s (Join-Path $root 'tests') 2>&1 | ForEach-Object { "$_" }
$ErrorActionPreference = 'Stop'
$tests | Select-String -Pattern '^(Ran|OK|FAILED)' | ForEach-Object { Write-Host "   $_" }
if (-not ($tests -match '^OK')) { throw 'Tests failed: not building.' }

$version = ((Select-String -Path (Join-Path $root 'server\server.py') -Pattern '^VERSION = "(.+)"').Matches[0].Groups[1].Value)
$v4 = (($version.Split('.') + @('0', '0', '0', '0'))[0..3] | ForEach-Object { [int]$_ }) -join ', '
# Version information in the exe (Properties > Details): antivirus treats an exe without it as more suspicious.
$versionFile = Join-Path $build 'version_info.txt'
@"
VSVersionInfo(
  ffi=FixedFileInfo(filevers=($v4), prodvers=($v4), mask=0x3f, flags=0x0, OS=0x40004, fileType=0x1, subtype=0x0, date=(0, 0)),
  kids=[StringFileInfo([StringTable('040904B0', [
    StringStruct('CompanyName', 'Techno Q'),
    StringStruct('FileDescription', 'TQ Connector for NetSuite (local server for Excel)'),
    StringStruct('FileVersion', '$version'),
    StringStruct('InternalName', 'TQConnector'),
    StringStruct('LegalCopyright', 'Copyright (c) Techno Q'),
    StringStruct('OriginalFilename', 'TQConnector.exe'),
    StringStruct('ProductName', 'TQ Connector for NetSuite'),
    StringStruct('ProductVersion', '$version')])]),
    VarFileInfo([VarStruct('Translation', [1033, 1200])])]
)
"@ | Set-Content -Path $versionFile -Encoding UTF8

Write-Host '-> Building TQConnector.exe'
$ErrorActionPreference = 'Continue'
& $py -m PyInstaller --noconfirm --clean --log-level WARN --noconsole --onedir --noupx --name TQConnector `
    --version-file $versionFile `
    --icon (Join-Path $root 'addin\assets\tq.ico') `
    --paths (Join-Path $root 'server') --hidden-import licence_key `
    --distpath (Join-Path $build 'dist') --workpath (Join-Path $build 'work') --specpath $build `
    (Join-Path $root 'server\server.py')
if ($LASTEXITCODE -ne 0) { throw 'PyInstaller failed' }
$ErrorActionPreference = 'Stop'

Write-Host '-> Staging the files'
if (Test-Path $stage) { Remove-Item $stage -Recurse -Force }
New-Item -ItemType Directory -Force $app | Out-Null
Copy-Item (Join-Path $build 'dist\TQConnector\*') $app -Recurse
Copy-Item (Join-Path $root 'addin') (Join-Path $app 'addin') -Recurse
Set-Content -Path (Join-Path $app 'VERSION.txt') -Value $version -Encoding ASCII

$sign = $env:TQ_SIGNTOOL
if ($sign -eq 'none') { $sign = '' }
elseif (-not $sign -and (Get-ChildItem Cert:\CurrentUser\My -CodeSigningCert | Where-Object Subject -like 'CN=Techno Q Code Signing*')) {
    # the repository path must have no spaces here (Inno Setup passes the command on as it is)
    $sign = 'powershell -NoProfile -ExecutionPolicy Bypass -File ' + (Join-Path $PSScriptRoot 'sign.ps1') + ' $f'
}
if ($sign) {
    Write-Host '-> Signing TQConnector.exe'
    $cmd = $sign.Replace('$f', '"' + (Join-Path $app 'TQConnector.exe') + '"')
    cmd.exe /c $cmd
    if ($LASTEXITCODE -ne 0) { throw 'Signing TQConnector.exe failed' }
}

Write-Host '-> Making TQConnectorSetup.exe with Inno Setup'
$iscc = @("$env:LOCALAPPDATA\Programs\Inno Setup 6\ISCC.exe", "${env:ProgramFiles(x86)}\Inno Setup 6\ISCC.exe",
          "$env:ProgramFiles\Inno Setup 6\ISCC.exe") | Where-Object { Test-Path $_ } | Select-Object -First 1
if (-not $iscc) { throw 'Inno Setup 6 is not installed. Install it: winget install --id JRSoftware.InnoSetup -e' }
if (Test-Path $setup) { Remove-Item $setup -Force }
$isccArgs = @("/DAppVersion=$version", "/DStage=$app", '/Qp')
if ($sign) { $isccArgs += @('/DSignTool=tq', "/Stq=$sign") }
$isccArgs += (Join-Path $PSScriptRoot 'TQConnector.iss')
$ErrorActionPreference = 'Continue'
& $iscc @isccArgs 2>&1 | ForEach-Object { "   $_" } | Select-Object -Last 5
$ErrorActionPreference = 'Stop'
if (-not (Test-Path $setup)) { throw 'Inno Setup did not make the setup file (see the lines above).' }
$hash = (Get-FileHash $setup -Algorithm SHA256).Hash
Set-Content -Path "$setup.sha256.txt" -Value "$hash  TQConnectorSetup.exe (version $version)" -Encoding ASCII
Copy-Item (Join-Path $PSScriptRoot 'TechnoQ.cer') $build -Force
Write-Host "   SHA-256 $hash"
$size = [math]::Round((Get-Item $setup).Length / 1MB, 1)
Write-Host "Done: $setup ($size MB, version $version$(if ($sign) { ', signed' } else { ', not signed' }))" -ForegroundColor Green
