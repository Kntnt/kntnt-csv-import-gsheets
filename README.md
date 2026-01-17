# CSV Import for Google Sheets

A Google Apps Script that automatically syncs CSV files from a Google Drive folder into a Google Sheet. Features a progress dialog, duplicate detection, and optional cleanup of removed files.

## Features

- **Auto-import on open** – Imports new CSV files every time you open the spreadsheet
- **Progress dialog** – Shows which file is being processed in real-time
- **Duplicate detection** – Skips files that have already been imported
- **Sync deletions** – Optionally removes rows when source CSV files are deleted
- **Selective columns** – Import all columns or only the ones you need
- **Performance optimized** – Batch operations minimize API calls

## File Structure

```
├── Code.gs              # Main script logic
└── ProgressDialog.html  # Modal dialog UI
```

## Setup

### 1. Create the Apps Script Project

1. Open your Google Sheet
2. Go to **Extensions → Apps Script**
3. Delete any existing code in `Code.gs`
4. Copy the contents of `Code.gs` from this repo and paste it

### 2. Create the Dialog File

1. In the Apps Script editor, click **+** next to "Files"
2. Select **HTML**
3. Name it exactly: `ProgressDialog` (without .html extension)
4. Copy the contents of `ProgressDialog.html` from this repo and paste it

### 3. Configure

Edit the `CONFIG` object at the top of `Code.gs`:

```javascript
const CONFIG = {
  FOLDER_ID: 'your-folder-id-here',
  SHEET_NAME: 'Data',
  DELIMITER: ',',
  SKIP_HEADER: true,
  COLS_TO_INCLUDE: [0, 1, 3, 4, 9, 11],
  SYNC_DELETIONS: true
};
```

| Option | Description |
|--------|-------------|
| `FOLDER_ID` | Google Drive folder ID containing your CSV files. Find it in the folder's URL: `drive.google.com/drive/folders/[FOLDER_ID]` |
| `SHEET_NAME` | Name of the sheet tab where data will be imported. Case-sensitive. |
| `DELIMITER` | Character separating values in your CSV files: `','`, `';'`, or `'\t'` |
| `SKIP_HEADER` | Set `true` to skip the first row of each CSV file |
| `COLS_TO_INCLUDE` | Zero-indexed array of columns to import (0 = A, 1 = B, etc). Set to `null` or `[]` to import all columns |
| `SYNC_DELETIONS` | Set `true` to remove rows when their source CSV is deleted from the folder |

> **Note:** Column A in your sheet is reserved for the source filename. Your CSV data starts in column B.

### 4. Save the Project

Press **Ctrl+S** (Windows) or **Cmd+S** (Mac) and wait for "Project saved" confirmation.

### 5. Create the Trigger

Simple triggers can't show modal dialogs due to Google's security restrictions. You need an installable trigger:

1. In the Apps Script editor, click the **clock icon** (⏰) in the left sidebar
2. Click **+ Add Trigger** (bottom right)
3. Configure:
   - **Choose which function to run:** `onOpenTrigger`
   - **Choose which deployment:** `Head`
   - **Select event source:** `From spreadsheet`
   - **Select event type:** `On open`
4. Click **Save**
5. Authorize when prompted

### 6. Test

Reload your spreadsheet (**F5** or **Ctrl+R** / **Cmd+R**).

The import dialog should appear automatically. Note that it may take a few seconds after page load.

## How It Works

1. On spreadsheet open, the trigger displays a modal dialog
2. The dialog starts the import process and polls for status updates
3. The script reads all CSV filenames from the configured Drive folder
4. If `SYNC_DELETIONS` is enabled, rows from deleted files are removed
5. New CSV files are parsed and their data is batch-written to the sheet
6. The dialog shows a summary and a close button

## Limitations

- **Memory:** Very large sheets (approaching 10 million cells) may cause memory issues
- **Execution time:** Google Apps Script has a 6-minute timeout
- **Permissions:** The script needs access to Google Drive and Sheets
