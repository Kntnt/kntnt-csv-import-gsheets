/**
 * CSV Import for Google Sheets
 * Syncs CSV files from a Drive folder into a Sheet with progress dialog.
 * See README.md for setup instructions.
 */

const CONFIG = {
  FOLDER_ID: '1ZSMSVBw9NswwvIAhUvqa081RfQ2RNiM8',
  SHEET_NAME: 'Data',
  DELIMITER: ',',
  SKIP_HEADER: true,
  COLS_TO_INCLUDE: [0, 1, 3, 4, 9, 11],  // Set to null or [] to import all columns
  SYNC_DELETIONS: true
};

/**
 * Entry point for the installable onOpen trigger.
 * Displays the progress dialog which then initiates the import.
 */
function onOpenTrigger() {
  const html = HtmlService.createHtmlOutputFromFile('ProgressDialog')
    .setWidth(450)
    .setHeight(180);
  SpreadsheetApp.getUi().showModalDialog(html, '📥 CSV Import');
}

/**
 * Main sync logic. Called by the dialog via google.script.run.
 * Syncs the sheet with the Drive folder: removes orphaned rows, adds new files.
 * Progress is communicated via ScriptProperties for UI polling.
 */
function importNewCSVFiles() {
  const props = PropertiesService.getScriptProperties();
  const updateStatus = (message, done = false) => {
    props.setProperty('importStatus', JSON.stringify({ message, done }));
  };

  try {
    updateStatus('Initializing...');

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
    const folder = DriveApp.getFolderById(CONFIG.FOLDER_ID);

    // Build map of CSV files currently in the Drive folder
    const files = folder.getFilesByType(MimeType.CSV);
    const folderFiles = new Map();
    while (files.hasNext()) {
      const file = files.next();
      folderFiles.set(file.getName(), file);
    }

    const lastRow = sheet.getLastRow();
    let existingData = [];
    let existingSet = new Set();
    let deletedRowsCount = 0;
    let deletedFilesCount = 0;

    // Read existing data once and handle deletions in memory
    if (lastRow > 0) {
      existingData = sheet.getDataRange().getValues();

      if (CONFIG.SYNC_DELETIONS) {
        updateStatus('Checking for deleted files...');

        const deletedFiles = new Set();
        const filteredData = existingData.filter(row => {
          const fileName = row[0];
          if (fileName && !folderFiles.has(fileName)) {
            deletedFiles.add(fileName);
            return false;
          }
          return true;
        });

        deletedFilesCount = deletedFiles.size;
        deletedRowsCount = existingData.length - filteredData.length;

        // Rewrite sheet only if rows were removed (single API call vs N deleteRow calls)
        if (deletedRowsCount > 0) {
          sheet.clearContents();
          if (filteredData.length > 0) {
            sheet.getRange(1, 1, filteredData.length, filteredData[0].length)
              .setValues(filteredData);
          }
          existingData = filteredData;
        }
      }

      existingSet = new Set(existingData.map(row => row[0]));
    }

    // Import new files
    const allNewRows = [];
    let newFilesCount = 0;
    const importAllCols = !CONFIG.COLS_TO_INCLUDE || CONFIG.COLS_TO_INCLUDE.length === 0;

    for (const [fileName, file] of folderFiles) {
      if (existingSet.has(fileName)) continue;

      updateStatus(`Reading: ${fileName}`);

      const csvContent = file.getBlob().getDataAsString('UTF-8');
      const csvData = Utilities.parseCsv(csvContent, CONFIG.DELIMITER);

      if (csvData.length === 0) continue;

      const startIdx = CONFIG.SKIP_HEADER ? 1 : 0;

      for (let i = startIdx; i < csvData.length; i++) {
        const row = csvData[i];
        const dataRow = importAllCols
          ? row
          : CONFIG.COLS_TO_INCLUDE.map(idx => row[idx] ?? '');
        allNewRows.push([fileName, ...dataRow]);
      }

      newFilesCount++;
    }

    // Batch write all new rows
    if (allNewRows.length > 0) {
      sheet.getRange(sheet.getLastRow() + 1, 1, allNewRows.length, allNewRows[0].length)
        .setValues(allNewRows);
    }

    const summary = buildSummary(newFilesCount, allNewRows.length, deletedFilesCount, deletedRowsCount);
    updateStatus(summary, true);

  } catch (error) {
    updateStatus(`❌ Error: ${error.message}`, true);
  }
}

/**
 * Builds a human-readable summary of the sync operation.
 */
function buildSummary(newFiles, newRows, deletedFiles, deletedRows) {
  const parts = [];

  if (newFiles > 0) {
    parts.push(`imported ${newFiles} file(s) [${newRows} rows]`);
  }
  if (deletedFiles > 0) {
    parts.push(`removed ${deletedFiles} file(s) [${deletedRows} rows]`);
  }

  if (parts.length === 0) {
    return '✅ No changes.';
  }

  return '✅ ' + parts.map(p => p.charAt(0).toUpperCase() + p.slice(1)).join(', ') + '.';
}

/**
 * Returns current import status. Polled by the dialog UI.
 */
function getImportStatus() {
  const status = PropertiesService.getScriptProperties().getProperty('importStatus');
  return status ? JSON.parse(status) : { message: 'Ready', done: false };
}
