# 1. Define where to hide the manifest and where to download it from
$ManifestDir = "$env:LOCALAPPDATA\ExcelCustomAddin"
# CHANGE THIS URL to where you host your manifest.xml
$ManifestUrl = "https://liamlaidlaw.com/tl_bookkeeping_addin/manifest.xml"

Write-Host "Installing Custom Excel Add-in..."

# 2. Create the hidden local folder
New-Item -ItemType Directory -Force -Path $ManifestDir | Out-Null

# 3. Download the XML manifest
Invoke-WebRequest -Uri $ManifestUrl -OutFile "$ManifestDir\manifest.xml"

# 4. Add the folder to the Registry so Excel trusts it as a catalog
$RegPath = "HKCU:\SOFTWARE\Microsoft\Office\16.0\WEF\TrustedCatalogs\CustomSearchAddin"
New-Item -Path $RegPath -Force | Out-Null
Set-ItemProperty -Path $RegPath -Name "Url" -Value $ManifestDir
Set-ItemProperty -Path $RegPath -Name "Flags" -Value 1

Write-Host "Installation Complete! Please restart Excel."
Write-Host "To use it: Open Excel -> Insert tab -> My Add-ins -> Shared Folder."
Pause
