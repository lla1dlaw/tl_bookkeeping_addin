$ManifestDir = "C:\ExcelCustomAddin"
$ManifestUrl = "https://liamlaidlaw.com/tl_bookkeeping_addin/manifest.xml?t=$([guid]::NewGuid())"

Write-Host "Installing Custom Excel Add-in (Pure Local)..."

# 1. Create the directory
if (-not (Test-Path $ManifestDir)) {
    New-Item -ItemType Directory -Force -Path $ManifestDir | Out-Null
}

# 2. Download the manifest and Unblock
Invoke-WebRequest -Uri $ManifestUrl -OutFile "$ManifestDir\manifest.xml"
Unblock-File -Path "$ManifestDir\manifest.xml" -ErrorAction SilentlyContinue

# 3. Inject the Local Path directly into TrustedCatalogs
$TrustedRoot = "HKCU:\SOFTWARE\Microsoft\Office\16.0\Wef\TrustedCatalogs"
if (-not (Test-Path $TrustedRoot)) { New-Item -Path $TrustedRoot -Force | Out-Null }

# Clear old entries pointing to our dir or the old localhost share
Get-ChildItem -Path $TrustedRoot | ForEach-Object {
    $urlVal = Get-ItemProperty -Path $_.PSPath -Name "Url" -ErrorAction SilentlyContinue
    if ($urlVal -and ($urlVal.Url -eq $ManifestDir -or $urlVal.Url -match "localhost\\ExcelCustomAddin")) {
        Remove-Item -Path $_.PSPath -Force -Recurse
    }
}

$Guid = [guid]::NewGuid().ToString("B").ToUpper()
$CatalogRegPath = "$TrustedRoot\$Guid"
New-Item -Path $CatalogRegPath -Force | Out-Null
Set-ItemProperty -Path $CatalogRegPath -Name "Id" -Value $Guid -Type String
Set-ItemProperty -Path $CatalogRegPath -Name "Url" -Value $ManifestDir -Type String
Set-ItemProperty -Path $CatalogRegPath -Name "Flags" -Value 1 -Type DWord

# 4. Clean cache
Remove-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Office\16.0\Wef\Developer" -Name "CustomSheetSearch" -ErrorAction SilentlyContinue
Remove-Item -Path "$env:LOCALAPPDATA\Microsoft\Office\16.0\Wef\Cache" -Recurse -Force -ErrorAction SilentlyContinue

Write-Host "✅ Installation Complete! No network shares needed. Please restart Excel."
