/**
 * CSV Import for Google Sheets
 * Syncs CSV files from a Drive folder into a Sheet with progress dialog.
 * See README.md for setup instructions.
 */

const CONFIG = {
  FOLDER_ID: '1ZSMSVBw9NswwvIAhUvqa081RfQ2RNiM8',
  FILE_REGEX: '\\.csv$',
  DELIMITER: ',',
  SKIP_ROWS: 1,
  COLS_TO_INCLUDE: [0, 1, 3, 4, 9, 11],
  SHEET_NAME: 'Data',
  SHEET_START_ROW: 2,
  SYNC_DELETIONS: true,
  CSV_DECIMAL_SEPARATOR: ',',    // Decimal separator in CSV files ('.' or ',')
  SHEET_DECIMAL_SEPARATOR: ',',  // Decimal separator expected by the sheet ('.' or ',')
};

/**
 * Converts decimal separators in a value if needed.
 * Only converts values that look like numbers (integers or decimals).
 */
function convertDecimalSeparator(value, fromSeparator, toSeparator) {
  if (fromSeparator === toSeparator) return value;
  if (typeof value !== 'string') value = String(value);

  const pattern = fromSeparator === '.'
    ? /^-?[0-9]+(\.[0-9]+)?$/
    : /^-?[0-9]+(,[0-9]+)?$/;

  if (pattern.test(value.trim())) {
    return value.replace(fromSeparator, toSeparator);
  }
  return value;
}

/**
 * Converts all decimal separators in a row of data.
 */
function convertRowDecimals(row, fromSeparator, toSeparator) {
  if (fromSeparator === toSeparator) return row;
  return row.map(cell => convertDecimalSeparator(cell, fromSeparator, toSeparator));
}

/**
 * Recursively collects CSV files from a folder that match FILE_REGEX.
 * Uses searchFiles for efficient CSV filtering in each folder.
 * Returns a Map of relative path -> file object.
 */
function getMatchingFiles(rootFolder, regexPattern) {
  const regex = new RegExp(regexPattern, 'i');
  const result = new Map();

  function traverse(folder, pathPrefix) {
    // Use searchFiles for efficient CSV filtering
    const files = folder.searchFiles("mimeType = 'text/csv'");
    while (files.hasNext()) {
      const file = files.next();
      const fileName = file.getName();
      const relativePath = pathPrefix ? pathPrefix + '/' + fileName : fileName;

      if (regex.test(relativePath)) {
        result.set(relativePath, file);
      }
    }

    // Recurse into subfolders
    const subfolders = folder.getFolders();
    while (subfolders.hasNext()) {
      const subfolder = subfolders.next();
      const subfolderName = subfolder.getName();
      const newPrefix = pathPrefix ? pathPrefix + '/' + subfolderName : subfolderName;
      traverse(subfolder, newPrefix);
    }
  }

  traverse(rootFolder, '');
  return result;
}

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
 * Clears import status. Called by dialog before starting import
 * to prevent stale status from previous runs being displayed.
 */
function clearImportStatus() {
  PropertiesService.getScriptProperties().deleteProperty('importStatus');
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

  const updateStatus = (message, done = false) => {
    props.setProperty('importStatus', JSON.stringify({ message, done }));
  };

  try {
    updateStatus('Initializing...');

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
    const folder = DriveApp.getFolderById(CONFIG.FOLDER_ID);
    const startRow = CONFIG.SHEET_START_ROW;

    // Build map of CSV files matching the regex (supports recursive search)
    updateStatus('Scanning for files...');
    const folderFiles = getMatchingFiles(folder, CONFIG.FILE_REGEX);

    // Determine the data region (from startRow to actual last row of content)
    const sheetLastRow = sheet.getLastRow();
    const sheetLastCol = sheet.getLastColumn() || 1;
    
    let existingData = [];
    let deletedRowsCount = 0;
    let deletedFilesCount = 0;

    // Read existing data from the data region (startRow onwards)
    if (sheetLastRow >= startRow) {
      const numRows = sheetLastRow - startRow + 1;
      existingData = sheet.getRange(startRow, 1, numRows, sheetLastCol).getValues();
      
      // Filter out completely empty rows (rows where all cells are empty)
      existingData = existingData.filter(row => row.some(cell => cell !== ''));
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
    const existingSet = new Set(filteredData.map(row => row[0]).filter(name => name));

    // Determine decimal separator conversion needs
    const csvDecimal = CONFIG.CSV_DECIMAL_SEPARATOR || ',';
    const sheetDecimal = CONFIG.SHEET_DECIMAL_SEPARATOR || ',';
    const needsDecimalConversion = csvDecimal !== sheetDecimal;

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
        let dataRow = importAllCols
          ? row
          : CONFIG.COLS_TO_INCLUDE.map(idx => row[idx] ?? '');

        // Convert decimal separators if needed
        if (needsDecimalConversion) {
          dataRow = convertRowDecimals(dataRow, csvDecimal, sheetDecimal);
        }

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

      // Clear the entire data region from startRow downwards
      if (sheetLastRow >= startRow) {
        sheet.getRange(startRow, 1, sheetLastRow - startRow + 1, sheetLastCol).clearContent();
      }

      // Write all data starting at startRow (atomic operation = one recalculation)
      if (finalData.length > 0) {
        // Normalize column count (pad shorter rows if needed)
        const maxCols = Math.max(...finalData.map(row => row.length));
        const normalizedData = finalData.map(row => {
          if (row.length < maxCols) {
            return [...row, ...Array(maxCols - row.length).fill('')];
          }
          return row;
        });

        // Always write from startRow, not from getLastRow()
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
