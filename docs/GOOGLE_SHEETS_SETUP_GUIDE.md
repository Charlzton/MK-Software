# MK Eye Center — Production Google Sheets Backend Integration Guide

This guide details the step-by-step setup to connect the MK Eye Center web application with your existing Google Sheets workbook using a production-grade, HMAC-SHA256 secured Google Apps Script backend.

---

### Workspaces Overview
The application supports two independent workspaces:
1. **Reception Workspace**:
   - **Spreadsheet ID**: `11rBhezcklOndIPa7YAHr3zIBh6WCXkp2235slmirnlI`
   - **Worksheet Tabs**: `PATIENT'S DATA`, `DVLA`, `MEDS`, `Sheet2` (plus user-created sheets via `+`)
   - **Save & Export**: Saves and exports Reception data only.
2. **Doctor Workspace**:
   - **Spreadsheet ID**: Configurable via `DOCTOR_SPREADSHEET_ID` environment variable.
   - **Architecture**: Completely isolated storage. Until the Doctor's workbook is provided and linked, an Excel-style workspace with an unconfigured notification banner is shown. Doctor changes are never mixed into the Reception workbook.

---

## 1. Step-by-Step Setup Instructions

### Step 1: Open Apps Script from Your Google Sheets Workbook
1. Open your browser and navigate to your existing MK Eye Center Google Sheets workbook:
   `https://docs.google.com/spreadsheets/d/11rBhezcklOndIPa7YAHr3zIBh6WCXkp2235slmirnlI/edit`
2. In the top Google Sheets menu bar, click **Extensions** > **Apps Script**.
3. A new browser tab will open showing the Apps Script code editor.

### Step 2: Create and Save the Script Files
1. In the Apps Script editor, locate the default file `Code.gs` in the left sidebar.
2. Select all existing boilerplate code in `Code.gs` and delete it.
3. Open `google_apps_script/Code.gs` from this project repository, copy the entire code, and paste it into the editor's `Code.gs` file.
4. Click the **Save** icon (floppy disk) or press `Ctrl+S` / `Cmd+S` to save.

### Step 3: Set the Required Script Properties (Security Secret)
The backend enforces server-to-server cryptographic HMAC-SHA256 signature verification so that public internet bots or unauthorized actors cannot access or modify clinic data:
1. In the Apps Script left navigation sidebar, click the **Project Settings** icon (gear icon at the bottom of the left sidebar).
2. Scroll down to the section titled **Script Properties**.
3. Click **Add script property**.
4. Set the fields:
   - **Property**: `API_SECRET`
   - **Value**: Enter a strong, high-entropy secret string (e.g. `mk_eye_center_production_secret_98234789234`).
5. Click **Save script properties**.

### Step 4: Deploy the Web App
1. In the top-right corner of the Apps Script editor, click the blue **Deploy** button > **New deployment**.
2. Click the gear icon next to "Select type" and choose **Web app**.
3. Fill in the deployment configuration:
   - **Description**: `MK Eye Center Production Spreadsheet API`
   - **Execute as**: `Me (your Google account)` *(ensures the script has permission to access the private spreadsheet without requiring staff Google accounts)*
   - **Who has access**: `Anyone` *(allows the secure AI Studio server-side backend to communicate with the endpoint using the HMAC signature)*
4. Click **Deploy**.
5. When prompted with "Authorization required", click **Authorize access**, choose your Google account, and grant the permissions.
   *(If Google displays "Google hasn't verified this app", click "Advanced" and then "Go to MK Eye Center (unsafe)" to approve).*

### Step 5: Obtain the Deployed `/exec` URL
1. Once deployed, Google will display a modal with the **Web app URL**.
2. The URL will look like:
   `https://script.google.com/macros/s/AKfycbxXXXXXXXXXXXXXXXXXXXXXXXXXXXXX/exec`
3. Click **Copy** to copy this URL to your clipboard.

### Step 6: Configure Matching Secrets in AI Studio Server-Side Environment
Configure the matching environment variables in your server environment (or AI Studio Secrets / `.env` file):
1. **`GOOGLE_APPS_SCRIPT_URL`**:
   `https://script.google.com/macros/s/YOUR_DEPLOYMENT_ID/exec`
2. **`GOOGLE_SHEETS_SHARED_SECRET`**:
   The exact same string you entered for `API_SECRET` in Step 3.
3. **`GOOGLE_SPREADSHEET_ID`**:
   `11rBhezcklOndIPa7YAHr3zIBh6WCXkp2235slmirnlI`
4. **`DOCTOR_SPREADSHEET_ID`**:
   *(Leave blank for now, or enter the Doctor's workbook ID when provided)*

---

## 2. Operations & Workspace Features

1. **Navigation**: Switch between **Reception** and **Doctor** via top navigation tabs. If you have pending edits, the application warns you before leaving the workspace.
2. **Worksheet Tabs & `+` Button**:
   - Both workspaces feature bottom tabs and an Excel-style `+` button.
   - Clicking `+` allows you to enter a worksheet name.
   - It validates against invalid characters and duplicates, adds the tab immediately, and syncs to Google Sheets.
3. **Save**:
   - The **Save** button (or `Ctrl+S`) saves the currently active workspace only.
4. **Export to Excel**:
   - Generates genuine `.xlsx` files:
     - `MK_Eye_Center_Reception_Export_<date>.xlsx` for Reception.
     - `MK_Eye_Center_Doctor_Export_<date>.xlsx` for Doctor.
