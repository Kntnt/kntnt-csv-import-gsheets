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
  CSV_LOCALE: 'en_US',
  SHEET_NAME: 'Data',
  SHEET_START_ROW: 2,
  SYNC_DELETIONS: true,
};

/**
 * Property key for storing original locale during import.
 */
const ORIGINAL_LOCALE_KEY = 'originalLocale';

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
 *
 * If a previous import's locale was not restored (user closed dialog via X),
 * restores it first which triggers a page reload.
 */
function onOpenTrigger() {
  const props = PropertiesService.getScriptProperties();

  // Check if locale needs to be restored from a previous import
  // (happens if user closed the dialog via X icon instead of Close button)
  const savedLocale = props.getProperty(ORIGINAL_LOCALE_KEY);
  if (savedLocale) {
    props.deleteProperty(ORIGINAL_LOCALE_KEY);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (ss.getSpreadsheetLocale() !== savedLocale) {
      // Restore locale - this will trigger a page reload and new onOpenTrigger
      ss.setSpreadsheetLocale(savedLocale);
      return;
    }
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
 * Always switches to CSV_LOCALE during import to ensure correct parsing of
 * numbers, dates, and formulas. The original locale is restored when the
 * user clicks the Close button (via restoreLocale).
 */
function importNewCSVFiles() {
  const props = PropertiesService.getScriptProperties();

  function updateStatus(message, done = false) {
    props.setProperty('importStatus', JSON.stringify({ message, done }));
  }

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

    // Always switch to CSV locale for correct parsing
    // Save original locale so it can be restored when dialog closes
    currentStep = 'switching locale';
    const originalLocale = ss.getSpreadsheetLocale();
    if (originalLocale !== CONFIG.CSV_LOCALE) {
      props.setProperty(ORIGINAL_LOCALE_KEY, originalLocale);
      ss.setSpreadsheetLocale(CONFIG.CSV_LOCALE);
      SpreadsheetApp.flush();
    }

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

      SpreadsheetApp.flush();
    }

    const summary = buildSummary(newFilesCount, newRows.length, deletedFilesCount, deletedRowsCount);
    updateStatus(summary, true);
  } catch (error) {
    const errorMsg = error.message || 'Unknown error';
    const context = currentFile
      ? `Error while ${currentStep} (${currentFile}): ${errorMsg}`
      : `Error while ${currentStep}: ${errorMsg}`;
    updateStatus(context, true);
  }
}

/**
 * Restores the original locale after import completes.
 * Called by the dialog when the user clicks the Close button.
 * This triggers a page reload which closes the dialog.
 */
function restoreLocale() {
  const props = PropertiesService.getScriptProperties();
  const savedLocale = props.getProperty(ORIGINAL_LOCALE_KEY);

  if (savedLocale) {
    props.deleteProperty(ORIGINAL_LOCALE_KEY);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (ss.getSpreadsheetLocale() !== savedLocale) {
      ss.setSpreadsheetLocale(savedLocale);
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
