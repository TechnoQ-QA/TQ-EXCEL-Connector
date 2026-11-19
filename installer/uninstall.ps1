# TQ Connector for NetSuite - removes the connector for the current Windows user.
# The data folder (settings and licence) is kept, so a later install picks them up again;
# add -All to remove it too.

param([switch]$All)

$app = Join-Path $env:LOCALAPPDATA 'TQ Connector'
$manifest = Join-Path $app 'addin\manifest.xml'
$devKey = 'HKCU:\Software\Microsoft\Office\16.0\WEF\Developer'

Get-Process TQConnector -ErrorAction SilentlyContinue | Stop-Process -Force
if (Test-Path $manifest) {
    $addinId = ([xml](Get-Content $manifest -Raw)).OfficeApp.Id
    if (Test-Path $devKey) { Remove-ItemProperty -Path $devKey -Name $addinId -ErrorAction SilentlyContinue }
}
Get-ChildItem Cert:\CurrentUser\Root | Where-Object { $_.Subject -like '*TQ Connector (this PC only)*' } | Remove-Item
Remove-Item (Join-Path ([Environment]::GetFolderPath('Startup')) 'TQ Connector.lnk') -ErrorAction SilentlyContinue
Remove-Item 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\TQConnector' -Recurse -ErrorAction SilentlyContinue
Start-Sleep -Milliseconds 500
if (Test-Path $app) {
    Get-ChildItem $app -Force | Where-Object { $All -or $_.Name -ne 'data' } | Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
    if ($All) { Remove-Item $app -Recurse -Force -ErrorAction SilentlyContinue }
}
Write-Host ('TQ Connector removed.' + $(if ($All) { '' } else { " Settings and licence kept in $app\data." }))
