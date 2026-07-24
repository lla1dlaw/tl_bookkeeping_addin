#!/usr/bin/env bash

echo "Installing Custom Excel Add-in for Mac..."

# 1. Define the hidden folder Mac Excel checks for side-loaded add-ins
WEF_DIR="$HOME/Library/Containers/com.microsoft.Excel/Data/Documents/wef"

# CHANGE THIS URL to where you host your manifest.xml
MANIFEST_URL="https://liamlaidlaw.com/tl_bookkeeping_addin/manifest.xml"

# 2. Create the directory if it doesn't exist
mkdir -p "$WEF_DIR"

# 3. Download the manifest
curl -s -o "$WEF_DIR/manifest.xml" "$MANIFEST_URL"

echo "Installation Complete!"
echo "Please restart Excel. Go to Insert > My Add-ins > Developer Add-ins to use it."
