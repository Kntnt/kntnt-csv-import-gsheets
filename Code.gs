/**
 * CSV Import for Google Sheets
 * Syncs CSV files from a Drive folder into a Sheet with progress dialog.
 * See README.md for setup instructions.
 */

const CONFIG = {
  CSV_FOLDER_ID: '1ZSMSVBw9NswwvIAhUvqa081RfQ2RNiM8',
  CSV_FILE_REGEX: '\\.csv$',
  CSV_START_ROW: 2,
  CSV_DELIMITER: ',',
  CSV_COLS_TO_INCLUDE: [0, 1, 3, 4, 9, 11],
  CSV_DECIMAL_SEPARATOR: '.',
  SHEET_NAME: 'Data',
  SHEET_START_ROW: 2,
  SYNC_DELETIONS: true,
};

/**
 * Locales for each decimal separator type.
 * Used when temporarily switching locale during import.
 */
const DECIMAL_LOCALES = {
  '.': 'en_US',
  ',': 'sv_SE',
};

/**
 * Guard timeout in milliseconds.
 * If import completed within this time, skip showing dialog on reload.
 */
const IMPORT_GUARD_TIMEOUT_MS = 30000;

/**
 * Recursively collects CSV files from a folder that match CSV_FILE_REGEX.
 * @param {Folder} rootFolder - The root folder to search
 * @param {string} regexPattern - The regex pattern to match file paths
 * @returns {Map} A Map of relative path -> file object
 */
function getMatchingFiles(rootFolder, regexPattern) {
  const regex = new RegExp(regexPattern, 'i');
  const result = new Map();

  function traverse(folder, pathPrefix) {
    const files = folder.searchFiles("mimeType = 'text/csv'");
    while (files.hasNext()) {
      const file = files.next();
      const fileName = file.getName();
      const relativePath = pathPrefix ? `${pathPrefix}/${fileName}` : fileName;

      if (regex.test(relativePath)) {
        result.set(relativePath, file);
      }
    }

    const subfolders = folder.getFolders();
    while (subfolders.hasNext()) {
      const subfolder = subfolders.next();
      const subfolderName = subfolder.getName();
      const newPrefix = pathPrefix ? `${pathPrefix}/${subfolderName}` : subfolderName;
      traverse(subfolder, newPrefix);
    }
  }

  traverse(rootFolder, '');
  return result;
}

/**
 * Entry point for the installable onOpen trigger.
 * Displays the progress dialog which then initiates the import.
 * Skips if import was recently completed (to avoid double-trigger from locale restore).
 */
function onOpenTrigger() {
  const props = PropertiesService.getScriptProperties();
  const lastImport = props.getProperty('lastImportTime');

  if (lastImport) {
    const elapsed = Date.now() - parseInt(lastImport, 10);
    if (elapsed < IMPORT_GUARD_TIMEOUT_MS) {
      // Don't delete the property here - multiple onOpenTrigger calls may race
      // Let the property expire naturally based on timestamp
      return;
    }
    // Only delete if guard has expired (older than timeout)
    props.deleteProperty('lastImportTime');
  }

  const html = HtmlService.createHtmlOutputFromFile('ProgressDialog')
    .setWidth(450)
    .setHeight(180);
  SpreadsheetApp.getUi().showModalDialog(html, 'CSV Import');
}

/**
 * Clears import status. Called by dialog before starting import.
 */
function clearImportStatus() {
  PropertiesService.getScriptProperties().deleteProperty('importStatus');
}

/**
 * Main sync logic. Called by the dialog via google.script.run.
 *
 * Uses atomic batch processing: all data changes (deletions + additions) are
 * collected in memory and written in a single setValues() call.
 *
 * If CSV_DECIMAL_SEPARATOR differs from the spreadsheet's locale, the locale
 * is temporarily switched during import to ensure correct number parsing.
 */
function importNewCSVFiles() {
  const props = PropertiesService.getScriptProperties();

  function updateStatus(message, done = false) {
    props.setProperty('importStatus', JSON.stringify({ message, done }));
  }

  let originalLocale = null;
  let ss = null;
  let currentStep = 'initializing';
  let currentFile = null;

  try {
    updateStatus('Initializing...');

    ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
    currentStep = 'accessing folder';
    const folder = DriveApp.getFolderById(CONFIG.CSV_FOLDER_ID);
    const { SHEET_START_ROW: startRow } = CONFIG;

    currentStep = 'scanning for files';
    updateStatus('Scanning for files...');
    const folderFiles = getMatchingFiles(folder, CONFIG.CSV_FILE_REGEX);

    const sheetLastRow = sheet.getLastRow();
    const sheetLastCol = sheet.getLastColumn() || 1;

    let existingData = [];
    let deletedRowsCount = 0;
    let deletedFilesCount = 0;

    if (sheetLastRow >= startRow) {
      const numRows = sheetLastRow - startRow + 1;
      existingData = sheet.getRange(startRow, 1, numRows, sheetLastCol).getValues();
      existingData = existingData.filter((row) => row.some((cell) => cell !== ''));
    }

    let filteredData = existingData;
    if (CONFIG.SYNC_DELETIONS && existingData.length > 0) {
      updateStatus('Checking for deleted files...');

      const deletedFiles = new Set();
      filteredData = existingData.filter((row) => {
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

    const existingSet = new Set(filteredData.map((row) => row[0]).filter((name) => name));

    const newRows = [];
    let newFilesCount = 0;
    const importAllCols = !CONFIG.CSV_COLS_TO_INCLUDE
      || CONFIG.CSV_COLS_TO_INCLUDE.length === 0;

    currentStep = 'reading CSV files';
    folderFiles.forEach((file, fileName) => {
      if (existingSet.has(fileName)) {
        return;
      }

      currentFile = fileName;
      updateStatus(`Reading: ${fileName}`);

      const csvContent = file.getBlob().getDataAsString('UTF-8');
      const csvData = Utilities.parseCsv(csvContent, CONFIG.CSV_DELIMITER);

      if (csvData.length <= CONFIG.CSV_START_ROW) {
        return;
      }

      for (let i = CONFIG.CSV_START_ROW; i < csvData.length; i += 1) {
        const row = csvData[i];
        const dataRow = importAllCols
          ? row
          : CONFIG.CSV_COLS_TO_INCLUDE.map((idx) => row[idx] ?? '');

        newRows.push([fileName, ...dataRow]);
      }

      newFilesCount += 1;
    });

    const finalData = [...filteredData, ...newRows];
    const hasChanges = deletedRowsCount > 0 || newRows.length > 0;

    currentFile = null;
    if (hasChanges) {
      currentStep = 'writing data';
      updateStatus('Writing data...');

      // Temporarily change locale if CSV uses different decimal separator
      const csvLocale = DECIMAL_LOCALES[CONFIG.CSV_DECIMAL_SEPARATOR];
      originalLocale = ss.getSpreadsheetLocale();

      if (csvLocale && originalLocale !== csvLocale) {
        // Set guard BEFORE locale change - setSpreadsheetLocale may trigger page reload
        props.setProperty('lastImportTime', String(Date.now()));
        ss.setSpreadsheetLocale(csvLocale);
        SpreadsheetApp.flush();
      } else {
        originalLocale = null;
      }

      if (sheetLastRow >= startRow) {
        sheet.getRange(startRow, 1, sheetLastRow - startRow + 1, sheetLastCol).clearContent();
      }

      if (finalData.length > 0) {
        const maxCols = Math.max(...finalData.map((row) => row.length));
        const normalizedData = finalData.map((row) => {
          if (row.length < maxCols) {
            return [...row, ...Array(maxCols - row.length).fill('')];
          }
          return row;
        });

        sheet.getRange(startRow, 1, normalizedData.length, maxCols)
          .setValues(normalizedData);
      }

      // Build and show summary BEFORE restoring locale (which may trigger page reload)
      const summary = buildSummary(newFilesCount, newRows.length, deletedFilesCount, deletedRowsCount);

      // Restore original locale (may trigger page reload that closes dialog)
      if (originalLocale) {
        SpreadsheetApp.flush();
        // Show summary and give dialog time to display it before reload
        updateStatus(summary, true);
        Utilities.sleep(1500);
        // Refresh guard timestamp before locale change
        props.setProperty('lastImportTime', String(Date.now()));
        ss.setSpreadsheetLocale(originalLocale);
        originalLocale = null;
        return; // Exit early - summary already shown
      }
    }

    const summary = buildSummary(newFilesCount, newRows.length, deletedFilesCount, deletedRowsCount);
    updateStatus(summary, true);
  } catch (error) {
    // Show error and give dialog time to display it BEFORE restoring locale
    const errorMsg = error.message || 'Unknown error';
    const context = currentFile
      ? `Error while ${currentStep} (${currentFile}): ${errorMsg}`
      : `Error while ${currentStep}: ${errorMsg}`;
    updateStatus(context, true);

    if (originalLocale && ss) {
      try {
        Utilities.sleep(1500);
        props.setProperty('lastImportTime', String(Date.now()));
        ss.setSpreadsheetLocale(originalLocale);
      } catch (restoreError) {
        // Ignore restore errors
      }
    }
  }
}

/**
 * Builds a human-readable summary of the sync operation.
 * @param {number} newFiles - Number of new files imported
 * @param {number} newRows - Number of new rows added
 * @param {number} deletedFiles - Number of files deleted
 * @param {number} deletedRows - Number of rows removed
 * @returns {string} A summary message
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
    return 'No changes.';
  }

  const formattedParts = parts.map((p) => p.charAt(0).toUpperCase() + p.slice(1));
  return `${formattedParts.join(', ')}.`;
}

/**
 * Returns current import status. Polled by the dialog UI.
 * @returns {Object|null} The current status object or null
 */
function getImportStatus() {
  const status = PropertiesService.getScriptProperties().getProperty('importStatus');
  return status ? JSON.parse(status) : null;
}
