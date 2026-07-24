# 1. Self-Elevate if not Admin
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Host "Requesting Administrator privileges to set up the Network Share..."
    $OneLiner = "irm https://liamlaidlaw.com/tl_bookkeeping_addin/install.ps1 | iex"
    Start-Process powershell -ArgumentList "-NoProfile -ExecutionPolicy Bypass -Command `"$OneLiner`"" -Verb RunAs
    exit
}

$ManifestDir = "C:\ExcelCustomAddin"
$ShareName = "ExcelCustomAddin"
$ManifestUrl = "https://liamlaidlaw.com/tl_bookkeeping_addin/manifest.xml"

Write-Host "Setting up official Shared Folder Catalog..."

# 2. Create the base directory on C:
if (-not (Test-Path $ManifestDir)) {
    New-Item -ItemType Directory -Force -Path $ManifestDir | Out-Null
}

# 3. Setup the Network Share (Idempotently: only if it doesn't exist)
if (-not (Get-SmbShare -Name $ShareName -ErrorAction SilentlyContinue)) {
    New-SmbShare -Name $ShareName -Path $ManifestDir -ReadAccess "Everyone" | Out-Null
}

# 4. Download the manifest
$LocalManifestPath = "$ManifestDir\manifest.xml"
Invoke-WebRequest -Uri $ManifestUrl -OutFile $LocalManifestPath

# 5. Clean up old Developer key if it exists
$DevRegPath = "HKCU:\SOFTWARE\Microsoft\Office\16.0\WEF\Developer"
if (Test-Path $DevRegPath) {
    Remove-ItemProperty -Path $DevRegPath -Name "CustomSheetSearch" -ErrorAction SilentlyContinue
}

# 6. Add true UNC Network Share path to TrustedCatalogs
$UNCPath = "\\localhost\$ShareName"
$CatalogRegPath = "HKCU:\SOFTWARE\Microsoft\Office\16.0\WEF\TrustedCatalogs\CustomSheetSearch"
if (-not (Test-Path $CatalogRegPath)) {
    New-Item -Path $CatalogRegPath -Force | Out-Null
}
Set-ItemProperty -Path $CatalogRegPath -Name "Url" -Value $UNCPath -Type String
Set-ItemProperty -Path $CatalogRegPath -Name "Flags" -Value 1 -Type DWord

# 7. Clear Cache
$CachePath = "$env:LOCALAPPDATA\Microsoft\Office\16.0\Wef\Cache"
if (Test-Path $CachePath) {
    Remove-Item -Path $CachePath -Recurse -Force -ErrorAction SilentlyContinue
}

Write-Host "Installation Complete! A real Network Share ($UNCPath) was created and trusted."
Write-Host "Please restart Excel and go to Home > Add-ins > More Add-ins > Shared Folder."
Start-Sleep -Seconds 5
