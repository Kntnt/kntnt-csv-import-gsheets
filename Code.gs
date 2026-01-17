/**
 * CSV Import for Google Sheets
 * Syncs CSV files from a Drive folder into a Sheet with progress dialog.
 * See README.md for setup instructions.
 */

const CONFIG = {
  FOLDER_ID: '1ZSMSVBw9NswwvIAhUvqa081RfQ2RNiM8',
  SHEET_NAME: 'Data',
  SHEET_START_ROW: 2
  DELIMITER: ',',
  SKIP_ROWS: 1,
  COLS_TO_INCLUDE: [0, 1, 3, 4, 9, 11],
  SYNC_DELETIONS: true,
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
 * 
 * Uses atomic batch processing: all data changes (deletions + additions) are
 * collected in memory and written in a single setValues() call. This ensures
 * that formula recalculation is triggered only once, after all data is in place.
 */
function importNewCSVFiles() {
  const props = PropertiesService.getScriptProperties();

  props.deleteProperty('importStatus');

  const updateStatus = (message, done = false) => {
    props.setProperty('importStatus', JSON.stringify({ message, done }));
  };

  try {
    updateStatus('Initializing...');

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
    const folder = DriveApp.getFolderById(CONFIG.FOLDER_ID);
    const startRow = CONFIG.SHEET_START_ROW;

    // Build map of CSV files currently in the Drive folder
    const files = folder.getFilesByType(MimeType.CSV);
    const folderFiles = new Map();
    while (files.hasNext()) {
      const file = files.next();
      folderFiles.set(file.getName(), file);
    }

    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn() || 1;
    let existingData = [];
    let deletedRowsCount = 0;
    let deletedFilesCount = 0;

    // Read existing data into memory
    if (lastRow >= startRow) {
      existingData = sheet.getRange(startRow, 1, lastRow - startRow + 1, lastCol).getValues();
    }

    // Filter out rows from deleted files (in memory)
    let filteredData = existingData;
    if (CONFIG.SYNC_DELETIONS && existingData.length > 0) {
      updateStatus('Checking for deleted files...');

      const deletedFiles = new Set();
      filteredData = existingData.filter(row => {
        const fileName = row[0];
        if (fileName && !folderFiles.has(fileName)) {
          deletedFiles.add(fileName);
          return false;
        }
        return true;
      });

      deletedFilesCount = deletedFiles.size;
      deletedRowsCount = existingData.length - filteredData.length;
    }

    // Build set of already imported filenames for duplicate detection
    const existingSet = new Set(filteredData.map(row => row[0]));

    // Import new files and collect rows in memory
    const newRows = [];
    let newFilesCount = 0;
    const importAllCols = !CONFIG.COLS_TO_INCLUDE || CONFIG.COLS_TO_INCLUDE.length === 0;

    for (const [fileName, file] of folderFiles) {
      if (existingSet.has(fileName)) continue;

      updateStatus(`Reading: ${fileName}`);

      const csvContent = file.getBlob().getDataAsString('UTF-8');
      const csvData = Utilities.parseCsv(csvContent, CONFIG.DELIMITER);

      if (csvData.length <= CONFIG.SKIP_ROWS) continue;

      for (let i = CONFIG.SKIP_ROWS; i < csvData.length; i++) {
        const row = csvData[i];
        const dataRow = importAllCols
          ? row
          : CONFIG.COLS_TO_INCLUDE.map(idx => row[idx] ?? '');
        newRows.push([fileName, ...dataRow]);
      }

      newFilesCount++;
    }

    // Combine filtered existing data with new rows
    const finalData = [...filteredData, ...newRows];

    // Determine if we need to write anything
    const hasChanges = deletedRowsCount > 0 || newRows.length > 0;

    if (hasChanges) {
      updateStatus('Writing data...');

      // Clear existing data area (clearContent doesn't trigger recalculation)
      if (lastRow >= startRow) {
        sheet.getRange(startRow, 1, lastRow - startRow + 1, lastCol).clearContent();
      }

      // Write all data in a single atomic operation (triggers ONE recalculation)
      if (finalData.length > 0) {
        // Normalize column count (pad shorter rows if needed)
        const maxCols = Math.max(...finalData.map(row => row.length));
        const normalizedData = finalData.map(row => {
          if (row.length < maxCols) {
            return [...row, ...Array(maxCols - row.length).fill('')];
          }
          return row;
        });

        sheet.getRange(startRow, 1, normalizedData.length, maxCols)
          .setValues(normalizedData);
      }
    }

    const summary = buildSummary(newFilesCount, newRows.length, deletedFilesCount, deletedRowsCount);
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
  return status ? JSON.parse(status) : null;
}
