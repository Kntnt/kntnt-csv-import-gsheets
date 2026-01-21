/**
 * CSV Import for Google Sheets
 *
 * Automatically syncs CSV files from a Google Drive folder into a Google Sheet.
 * Features locale switching for proper decimal separator handling, progress dialog,
 * duplicate detection, and optional cleanup of removed files.
 *
 * @see README.md for setup instructions and configuration options.
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

/** Script property key for storing the original locale before switching. */
const ORIGINAL_LOCALE_KEY = 'originalLocale';

/** Script property key to prevent trigger restart after locale restore. */
const IMPORT_DONE_KEY = 'importDone';

/**
 * Recursively collects CSV files from a folder that match the configured regex pattern.
 *
 * @param {Folder} rootFolder - The Google Drive folder to search.
 * @param {string} regexPattern - Regular expression pattern to filter files.
 * @returns {Map<string, File>} Map of relative file paths to File objects.
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
 *
 * Handles locale switching and displays the import dialog. The flow is:
 * 1. If locale differs from CSV_LOCALE: save original, switch locale (triggers page reload)
 * 2. After reload (or if no switch needed): display the import dialog
 * 3. If returning from a completed import: clear flag and exit without action
 */
function onOpenTrigger() {
  const props = PropertiesService.getScriptProperties();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const currentLocale = ss.getSpreadsheetLocale();
  const savedLocale = props.getProperty(ORIGINAL_LOCALE_KEY);
  const importDone = props.getProperty(IMPORT_DONE_KEY);

  // After import completion and locale restore, just clear the flag and exit
  if (importDone) {
    props.deleteProperty(IMPORT_DONE_KEY);
    return;
  }

  // Switch to CSV locale if needed (this triggers a page reload)
  if (currentLocale !== CONFIG.CSV_LOCALE && !savedLocale) {
    props.setProperty(ORIGINAL_LOCALE_KEY, currentLocale);
    ss.setSpreadsheetLocale(CONFIG.CSV_LOCALE);
    return;
  }

  // Display the import dialog
  const template = HtmlService.createTemplateFromFile('ProgressDialog');
  template.originalLocale = savedLocale || '';
  template.currentLocale = CONFIG.CSV_LOCALE;

  const dialogHeight = savedLocale ? 250 : 180;
  const html = template.evaluate()
    .setWidth(450)
    .setHeight(dialogHeight);
  SpreadsheetApp.getUi().showModalDialog(html, 'CSV Import');
}

/**
 * Restores the original spreadsheet locale after import completes.
 * Called by the dialog when the user clicks the Close button.
 *
 * Sets IMPORT_DONE_KEY before changing locale to prevent onOpenTrigger
 * from restarting the import process after the page reloads.
 */
function restoreLocale() {
  const props = PropertiesService.getScriptProperties();
  const savedLocale = props.getProperty(ORIGINAL_LOCALE_KEY);

  if (savedLocale) {
    props.setProperty(IMPORT_DONE_KEY, 'true');
    props.deleteProperty(ORIGINAL_LOCALE_KEY);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (ss.getSpreadsheetLocale() !== savedLocale) {
      ss.setSpreadsheetLocale(savedLocale);
    }
  }
}

/**
 * Main import logic. Syncs CSV files from the configured Drive folder to the sheet.
 * Called by the dialog via google.script.run.
 *
 * Updates import status throughout the process so the dialog can display progress.
 */
function importNewCSVFiles() {
  const props = PropertiesService.getScriptProperties();

  /**
   * Updates the import status for the dialog to poll.
   * @param {string} message - Status message to display.
   * @param {boolean} [done=false] - Whether the import is complete.
   */
  function updateStatus(message, done = false) {
    props.setProperty('importStatus', JSON.stringify({ message, done }));
  }

  let currentStep = 'initializing';
  let currentFile = null;

  try {
    updateStatus('Initializing...');

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
    const folder = DriveApp.getFolderById(CONFIG.CSV_FOLDER_ID);
    const { SHEET_START_ROW: startRow } = CONFIG;

    currentStep = 'scanning for files';
    updateStatus('Scanning for files...');
    const folderFiles = getMatchingFiles(folder, CONFIG.CSV_FILE_REGEX);

    const sheetLastRow = sheet.getLastRow();
    const sheetLastCol = sheet.getLastColumn() || 1;

    // Load existing data from sheet
    let existingData = [];
    if (sheetLastRow >= startRow) {
      const numRows = sheetLastRow - startRow + 1;
      existingData = sheet.getRange(startRow, 1, numRows, sheetLastCol).getValues();
      existingData = existingData.filter((row) => row.some((cell) => cell !== ''));
    }

    // Remove rows from deleted files if sync deletions is enabled
    let filteredData = existingData;
    let deletedRowsCount = 0;
    let deletedFilesCount = 0;

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

    // Build set of already-imported file names
    const existingSet = new Set(
      filteredData.map((row) => row[0]).filter((name) => name),
    );

    // Import new CSV files
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

    // Write data to sheet if there are changes
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
 * Builds a human-readable summary of the import operation.
 *
 * @param {number} newFiles - Number of new files imported.
 * @param {number} newRows - Total number of new rows added.
 * @param {number} deletedFiles - Number of files whose rows were removed.
 * @param {number} deletedRows - Total number of rows removed.
 * @returns {string} Summary message for display in the dialog.
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
 * Returns the current import status. Polled by the dialog to update the UI.
 *
 * @returns {Object|null} Status object with message and done properties, or null.
 */
function getImportStatus() {
  const status = PropertiesService.getScriptProperties().getProperty('importStatus');
  return status ? JSON.parse(status) : null;
}

/**
 * Clears the import status. Called by the dialog before starting a new import.
 */
function clearImportStatus() {
  PropertiesService.getScriptProperties().deleteProperty('importStatus');
}
