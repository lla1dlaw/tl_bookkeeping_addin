$ManifestDir = "C:\ExcelCustomAddin"
$ShareName = "ExcelCustomAddin"
$ManifestUrl = "https://liamlaidlaw.com/tl_bookkeeping_addin/manifest.xml"
$UNCPath = "\\localhost\$ShareName"

Write-Host "Installing Custom Excel Add-in: v1.2"

# ====================================================================
# PHASE 1: RUN IN CURRENT USER CONTEXT
# We MUST write the registry keys BEFORE elevating to Admin. 
# If we elevate first, HKCU writes to the Administrator's registry, 
# not the user's registry, so it never shows up in the user's Excel!
# ====================================================================

$CatalogRegPath = "HKCU:\SOFTWARE\Microsoft\Office\16.0\Wef\TrustedCatalogs\CustomSheetSearch"
if (-not (Test-Path $CatalogRegPath)) {
    New-Item -Path $CatalogRegPath -Force | Out-Null
}
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

Write-Host "Registry configured successfully for the current user."

# ====================================================================
# PHASE 2: RUN IN ADMIN CONTEXT
# Now we evaluate if we need Admin rights to create the SMB Share on C:
# ====================================================================

$AdminScript = {
    $ManifestDir = "C:\ExcelCustomAddin"
    $ShareName = "ExcelCustomAddin"
    $ManifestUrl = "https://liamlaidlaw.com/tl_bookkeeping_addin/manifest.xml"
    
    Write-Host "Creating Network Share..."
    if (-not (Test-Path $ManifestDir)) {
        New-Item -ItemType Directory -Force -Path $ManifestDir | Out-Null
    }
    if (-not (Get-SmbShare -Name $ShareName -ErrorAction SilentlyContinue)) {
        New-SmbShare -Name $ShareName -Path $ManifestDir -ReadAccess "Everyone" | Out-Null
    }
    
    Write-Host "Downloading Manifest..."
    Invoke-WebRequest -Uri $ManifestUrl -OutFile "$ManifestDir\manifest.xml"
    
    Write-Host "Network Share created successfully! You can close this window."
    Start-Sleep -Seconds 5
}

if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Host "Requesting Administrator privileges to set up the Network Share on C:..."
    # Package the admin script into a base64 encoded command to run in the elevated prompt securely
    $Encoded = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($AdminScript.ToString()))
    Start-Process powershell -ArgumentList "-NoProfile -ExecutionPolicy Bypass -EncodedCommand $Encoded" -Verb RunAs
} else {
    # Already running as Admin, execute directly
    Invoke-Command -ScriptBlock $AdminScript
}

Write-Host "Installation Complete! Please restart Excel and go to Home > Add-ins > More Add-ins > Shared Folder."
