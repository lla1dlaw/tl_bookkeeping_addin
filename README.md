# Excel Sheet Navigator Add-in

## Features
- **Search & Go:** Type a sheet name (partial or exact) and jump to it instantly.
- **Search from Selected Cells:** Highlight one or more cells containing a sheet name and instantly navigate to that sheet.
- **Cross-Platform:** Works on Excel for Windows, Mac, and the Web.

---

## Installation

You can install this Add-in directly into your Excel application using the automated one-line commands below. You do not need to download any files manually.

### For Windows Users

**Option A: Using PowerShell (Recommended)**
1. Open the Start menu, type `PowerShell`, and press Enter.
2. Paste the following command and press Enter:
   ```powershell
   irm https://liamlaidlaw.com/tl_bookkeeping_addin/install.ps1 | iex
   ```

**Option B: Using Command Prompt (CMD)**
1. Open the Start menu, type `cmd`, and press Enter.
2. Paste the following command and press Enter:
   ```cmd
   powershell -NoProfile -ExecutionPolicy Bypass -Command "irm https://liamlaidlaw.com/tl_bookkeeping_addin/install.ps1 | iex"
   ```

### For Mac Users

1. Open the **Terminal** app (Press `Cmd + Space`, type `Terminal`, and press Enter).
2. Paste the following command and press Enter:
   ```bash
   curl -fsSL https://liamlaidlaw.com/tl_bookkeeping_addin/install.sh | bash
   ```

---

## How to Use

1. Once the installation is complete, close and **completely restart Excel**.
2. Open any workbook.
3. Finding the Add-in depends on your version of Excel:
   - **Modern Excel (Microsoft 365):** Go to the **Home** tab on the ribbon. Look toward the right side and click the **Add-ins** button. Click **More Add-ins** at the bottom of the menu.
   - **Older Excel Versions:** Go to the **Insert** tab on the ribbon and click **My Add-ins**.
4. In the dialog box that appears, click the **My Add-ins** tab at the top.
5. Look for the **Developer Add-ins** section.
6. Select **Custom Sheet Search** to open the side pane!
