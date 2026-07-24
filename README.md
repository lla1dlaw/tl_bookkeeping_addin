# Excel Sheet Navigator Add-in

## Features
- **Search & Go:** Type a sheet name (partial or exact) and jump to it instantly.
- **Search from Selected Cells:** Highlight one or more cells containing a sheet name and instantly navigate to that sheet.
- **Cross-Platform:** Works on Excel for Windows, Mac, and the Web.

---

## Installation

You can install this Add-in directly into your Excel application using the automated one-line commands below. You do not need to download any files manually.

*(Note: Replace `your-username` and `your-repo` with your actual GitHub details. If you receive a security warning from Windows or your Mac, this is normal for internal scripts downloaded directly from the web.)*

### For Windows Users

**Option A: Using PowerShell (Recommended)**
1. Open the Start menu, type `PowerShell`, and press Enter.
2. Paste the following command and press Enter:
   ```powershell
   irm https://your-username.github.io/your-repo/install.ps1 | iex
   ```

**Option B: Using Command Prompt (CMD)**
1. Open the Start menu, type `cmd`, and press Enter.
2. Paste the following command and press Enter:
   ```cmd
   powershell -NoProfile -ExecutionPolicy Bypass -Command "irm https://your-username.github.io/your-repo/install.ps1 | iex"
   ```

### For Mac Users

1. Open the **Terminal** app (Press `Cmd + Space`, type `Terminal`, and press Enter).
2. Paste the following command and press Enter:
   ```bash
   curl -fsSL https://your-username.github.io/your-repo/install.sh | bash
   ```

---

## How to Use

1. Once the installation is complete, close and completely restart Excel.
2. Open any workbook.
3. Go to the **Insert** tab on the ribbon.
4. Click **My Add-ins** (or **Add-ins** depending on your Excel version).
5. On Windows, look under the **Shared Folder** tab. On Mac, look under **Developer Add-ins**.
6. Select **Custom Sheet Search** to open the side pane!

## Troubleshooting

- **I don't see the Add-in under "My Add-ins"**: Make sure you have completely closed all Excel windows and restarted the application after running the installation script.
- **Excel on the Web**: The installation scripts above are for Desktop versions of Excel (Windows/Mac). To use this Add-in on Excel for the Web, go to **Insert > Add-ins**, click **Upload My Add-in**, and manually upload the `manifest.xml` file.
