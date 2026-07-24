$ManifestDir = "$env:LOCALAPPDATA\ExcelCustomAddin"
$ManifestUrl = "https://liamlaidlaw.com/tl_bookkeeping_addin/manifest.xml"

Write-Host "Installing Custom Excel Add-in (Production Mode)..."

# 1. Create directory if it doesn't exist
if (-not (Test-Path $ManifestDir)) {
    New-Item -ItemType Directory -Force -Path $ManifestDir | Out-Null
}

# 2. Download the updated manifest
$LocalManifestPath = "$ManifestDir\manifest.xml"
Invoke-WebRequest -Uri $ManifestUrl -OutFile $LocalManifestPath

# 3. Clean up the old "Developer" mode installation to prevent conflicts
$DevRegPath = "HKCU:\SOFTWARE\Microsoft\Office\16.0\WEF\Developer"
if (Test-Path $DevRegPath) {
    Remove-ItemProperty -Path $DevRegPath -Name "CustomSheetSearch" -ErrorAction SilentlyContinue
}

# 4. Add to the Production "TrustedCatalogs" Registry Key
# (Injecting a local path here bypasses the UI's network share requirement)
$CatalogRegPath = "HKCU:\SOFTWARE\Microsoft\Office\16.0\WEF\TrustedCatalogs\CustomSheetSearch"
if (-not (Test-Path $CatalogRegPath)) {
    New-Item -Path $CatalogRegPath -Force | Out-Null
}
Set-ItemProperty -Path $CatalogRegPath -Name "Url" -Value $ManifestDir -Type String
Set-ItemProperty -Path $CatalogRegPath -Name "Flags" -Value 1 -Type DWord

# 5. Clear the Office Add-in Cache idempotently so the new ribbon button shows up instantly
$CachePath = "$env:LOCALAPPDATA\Microsoft\Office\16.0\Wef\Cache"
if (Test-Path $CachePath) {
    Remove-Item -Path $CachePath -Recurse -Force -ErrorAction SilentlyContinue
}

Write-Host "Installation Complete! Please restart Excel."
