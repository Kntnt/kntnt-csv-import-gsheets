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

/** Property key for storing original locale during import. */
const ORIGINAL_LOCALE_KEY = 'originalLocale';

/** Property key to prevent restart after locale restore. */
const IMPORT_DONE_KEY = 'importDone';

/**
 * Recursively collects CSV files from a folder that match CSV_FILE_REGEX.
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
 * Flow:
 * 1. If locale needs to change: save original, change locale, return (page reloads)
 * 2. After reload (or if no change needed): show dialog
 */
function onOpenTrigger() {
  const props = PropertiesService.getScriptProperties();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const currentLocale = ss.getSpreadsheetLocale();
  const savedLocale = props.getProperty(ORIGINAL_LOCALE_KEY);
  const importDone = props.getProperty(IMPORT_DONE_KEY);

  console.log('onOpenTrigger: currentLocale=' + currentLocale + ', CSV_LOCALE=' + CONFIG.CSV_LOCALE + ', savedLocale=' + savedLocale + ', importDone=' + importDone);

  // If import just completed and locale was restored, don't restart
  if (importDone) {
    console.log('onOpenTrigger: Import was just completed, clearing flag and exiting');
    props.deleteProperty(IMPORT_DONE_KEY);
    return;
  }

  // If we haven't switched to CSV locale yet, do it now
  // This triggers a page reload, after which we'll show the dialog
  if (currentLocale !== CONFIG.CSV_LOCALE && !savedLocale) {
    console.log('onOpenTrigger: Switching locale from ' + currentLocale + ' to ' + CONFIG.CSV_LOCALE);
    props.setProperty(ORIGINAL_LOCALE_KEY, currentLocale);
    ss.setSpreadsheetLocale(CONFIG.CSV_LOCALE);
    return; // Page will reload
  }

  // Show dialog - either after locale switch or if no switch was needed
  const html = HtmlService.createHtmlOutputFromFile('ProgressDialog')
    .setWidth(450)
    .setHeight(250);
  SpreadsheetApp.getUi().showModalDialog(html, 'CSV Import');
}

/**
 * Restores the original locale after import completes.
 * Called by the dialog's closeHandler when user closes the dialog.
 */
function restoreLocale() {
  const props = PropertiesService.getScriptProperties();
  const savedLocale = props.getProperty(ORIGINAL_LOCALE_KEY);

  console.log('restoreLocale: savedLocale=' + savedLocale);

  if (savedLocale) {
    // Set flag BEFORE changing locale to prevent onOpenTrigger from restarting
    props.setProperty(IMPORT_DONE_KEY, 'true');
    props.deleteProperty(ORIGINAL_LOCALE_KEY);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const currentLocale = ss.getSpreadsheetLocale();
    console.log('restoreLocale: currentLocale=' + currentLocale + ', restoring to ' + savedLocale);
    if (currentLocale !== savedLocale) {
      ss.setSpreadsheetLocale(savedLocale);
      // Page will reload, onOpenTrigger will see IMPORT_DONE_KEY and exit
    }
  } else {
    console.log('restoreLocale: No saved locale found');
  }
}

/**
 * Main sync logic. Called by the dialog via google.script.run.
 */
function importNewCSVFiles() {
  const props = PropertiesService.getScriptProperties();

  function updateStatus(message, done = false) {
    props.setProperty('importStatus', JSON.stringify({ message, done }));
  }

  let currentStep = 'initializing';
  let currentFile = null;

  try {
    updateStatus('Initializing...');

    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // Log locale at import time to verify it changed
    const importLocale = ss.getSpreadsheetLocale();
    console.log('importNewCSVFiles: Current locale at import time = ' + importLocale + ', expected = ' + CONFIG.CSV_LOCALE);

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
    return 'No changes.';
  }

  const formattedParts = parts.map((p) => p.charAt(0).toUpperCase() + p.slice(1));
  return `${formattedParts.join(', ')}.`;
}

/**
 * Returns current import status. Polled by the dialog UI.
 */
function getImportStatus() {
  const status = PropertiesService.getScriptProperties().getProperty('importStatus');
  return status ? JSON.parse(status) : null;
}

/**
 * Clears import status. Called by dialog before starting import.
 */
function clearImportStatus() {
  PropertiesService.getScriptProperties().deleteProperty('importStatus');
}

/**
 * Returns locale change info for the dialog to display a warning.
 * Returns null if no locale change was made.
 */
function getLocaleChangeInfo() {
  const savedLocale = PropertiesService.getScriptProperties().getProperty(ORIGINAL_LOCALE_KEY);
  if (savedLocale) {
    return {
      originalLocale: savedLocale,
      currentLocale: CONFIG.CSV_LOCALE
    };
  }
  return null;
}
