$ManifestDir = "C:\ExcelCustomAddin"
$ShareName = "ExcelCustomAddin"
$ManifestUrl = "https://liamlaidlaw.com/tl_bookkeeping_addin/manifest.xml"
$UNCPath = "\\localhost\$ShareName"

Write-Host "Installing Custom Excel Add-in..."

# ====================================================================
# PHASE 1: REGISTRY SETUP (RUN IN CURRENT USER CONTEXT)
# ====================================================================

$TrustedRoot = "HKCU:\SOFTWARE\Microsoft\Office\16.0\Wef\TrustedCatalogs"
if (-not (Test-Path $TrustedRoot)) { New-Item -Path $TrustedRoot -Force | Out-Null }

# Clean up any existing catalogs pointing to our share
Get-ChildItem -Path $TrustedRoot | ForEach-Object {
    $urlVal = Get-ItemProperty -Path $_.PSPath -Name "Url" -ErrorAction SilentlyContinue
    if ($urlVal -and $urlVal.Url -eq $UNCPath) {
        Remove-Item -Path $_.PSPath -Force -Recurse
    }
}

# Create new catalog with EXACT GUID structure required by Excel
$Guid = [guid]::NewGuid().ToString("B").ToUpper() # e.g. {12345678-1234-1234-1234-123456789012}
$CatalogRegPath = "$TrustedRoot\$Guid"
New-Item -Path $CatalogRegPath -Force | Out-Null
Set-ItemProperty -Path $CatalogRegPath -Name "Id" -Value $Guid -Type String
Set-ItemProperty -Path $CatalogRegPath -Name "Url" -Value $UNCPath -Type String
Set-ItemProperty -Path $CatalogRegPath -Name "Flags" -Value 1 -Type DWord

$DevRegPath = "HKCU:\SOFTWARE\Microsoft\Office\16.0\Wef\Developer"
if (Test-Path $DevRegPath) {
    Remove-ItemProperty -Path $DevRegPath -Name "CustomSheetSearch" -ErrorAction SilentlyContinue
}

$CachePath = "$env:LOCALAPPDATA\Microsoft\Office\16.0\Wef\Cache"
if (Test-Path $CachePath) {
    Remove-Item -Path $CachePath -Recurse -Force -ErrorAction SilentlyContinue
}

Write-Host "✅ Registry configured with exact Microsoft GUID structure."

# ====================================================================
# PHASE 2: SMB SHARE SETUP (RUN IN ADMIN CONTEXT)
# ====================================================================

$AdminScript = {
    $ManifestDir = "C:\ExcelCustomAddin"
    $ShareName = "ExcelCustomAddin"
    $ManifestUrl = "https://liamlaidlaw.com/tl_bookkeeping_addin/manifest.xml?t=$([guid]::NewGuid())"
    
    Write-Host "Creating Network Share..."
    if (-not (Test-Path $ManifestDir)) {
        New-Item -ItemType Directory -Force -Path $ManifestDir | Out-Null
    }
    if (-not (Get-SmbShare -Name $ShareName -ErrorAction SilentlyContinue)) {
        New-SmbShare -Name $ShareName -Path $ManifestDir -ReadAccess "Everyone" | Out-Null
    }
    
    Write-Host "Downloading Manifest..."
    Invoke-WebRequest -Uri $ManifestUrl -OutFile "$ManifestDir\manifest.xml"
    
    Write-Host "✅ Network Share created successfully! You can close this window."
    Start-Sleep -Seconds 5
}

if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Host "Requesting Administrator privileges to set up the Network Share on C:..."
    $Encoded = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($AdminScript.ToString()))
    Start-Process powershell -ArgumentList "-NoProfile -ExecutionPolicy Bypass -EncodedCommand $Encoded" -Verb RunAs
} else {
    Invoke-Command -ScriptBlock $AdminScript
}

Write-Host "Installation Complete! Please restart Excel and go to Home > Add-ins > More Add-ins > Shared Folder."
