$ManifestDir = "$env:LOCALAPPDATA\ExcelCustomAddin"
$ManifestUrl = "https://liamlaidlaw.com/tl_bookkeeping_addin/manifest.xml"

Write-Host "Installing Custom Excel Add-in..."

# Create directory and download the XML
New-Item -ItemType Directory -Force -Path $ManifestDir | Out-Null
$LocalManifestPath = "$ManifestDir\manifest.xml"
Invoke-WebRequest -Uri $ManifestUrl -OutFile $LocalManifestPath

# Add to the Developer registry key (this natively accepts local file paths!)
$RegPath = "HKCU:\SOFTWARE\Microsoft\Office\16.0\WEF\Developer"
if (-not (Test-Path $RegPath)) { New-Item -Path $RegPath -Force | Out-Null }
Set-ItemProperty -Path $RegPath -Name "CustomSheetSearch" -Value $LocalManifestPath -Type String

Write-Host "Installation Complete! Please restart Excel."
