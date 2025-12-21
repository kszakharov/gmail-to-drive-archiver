// Configuration Constants
const CONFIG = {
  DUPLICATE_MODE: '__DUPLICATE_MODE__',
  FOLDER_ID: '__FOLDER_ID__',
  INITIAL_LAST_RUN: '__INITIAL_LAST_RUN__',
  SEARCH_QUERY: '__SEARCH_QUERY__',
  DUPLICATE_MODES: {
    IGNORE: 'ignore',
    OVERWRITE: 'overwrite',
  },
};

// Property Keys
const PROPS = {
  LAST_RUN: 'lastRun',
};

/**
 * Main entry point: Saves new emails from Gmail to Google Drive
 * Tracks last run time to avoid processing duplicate emails
 */
function saveNewEmailsToDrive() {
  try {
    const startTime = new Date();
    const stats = { savedCount: 0, skippedCount: 0, errorCount: 0 };

    const folder = DriveApp.getFolderById(CONFIG.FOLDER_ID);
    const lastRun = getLastRunTimestamp();
    let newestTimestamp = 0;

    const threads = GmailApp.search(`after:${lastRun}`);

    threads.forEach(thread => {
      thread.getMessages().forEach(msg => {
        try {
          const result = processSingleEmail(msg, folder);
          stats[result.status]++;

          if (result.timestamp && result.timestamp > newestTimestamp) {
            newestTimestamp = result.timestamp;
          }
        } catch (error) {
          stats.errorCount++;
          Logger.log(`ERROR processing email: ${error.message}`);
        }
      });
    });

    updateLastRunTimestamp(newestTimestamp);
    logExecutionSummary(startTime, stats);
  } catch (error) {
    Logger.log(`CRITICAL ERROR: ${error.message}`);
    throw error;
  }
}

/**
 * Processes a single email: validates, checks for duplicates, and saves to Drive
 * @param {GmailMessage} msg - The email message to process
 * @param {Folder} folder - The root Google Drive folder
 * @returns {Object} Result object with status ('savedCount', 'skippedCount', 'errorCount') and timestamp
 */
function processSingleEmail(msg, folder) {
  const date = msg.getDate();
  const filename = generateEmailFilename(msg, date);
  const yearFolder = getOrCreateYearFolder(folder, date);

  const existingFile = findExistingFile(yearFolder, filename);
  if (existingFile) {
    return handleDuplicateEmail(existingFile, filename, CONFIG.DUPLICATE_MODE);
  }

  const savedFile = saveEmailToFolder(yearFolder, filename, msg, date);
  Logger.log(`Saved email: ${filename}`);

  return {
    status: 'savedCount',
    timestamp: Math.floor(date.getTime() / 1000),
  };
}

/**
 * Generates a sanitized filename for the email
 * @param {GmailMessage} msg - The email message
 * @param {Date} date - The email date
 * @returns {string} Sanitized filename with .eml extension
 */
function generateEmailFilename(msg, date) {
  const baseFilename = `${date.toISOString()} - ${msg.getSubject()}`;
  const sanitized = baseFilename.replace(/[\/\\?%*:|"<>]/g, '_');
  return `${sanitized}.eml`;
}

/**
 * Gets or creates a year-based folder in Google Drive
 * @param {Folder} parentFolder - The parent folder
 * @param {Date} date - The date used to determine the year
 * @returns {Folder} The year folder
 */
function getOrCreateYearFolder(parentFolder, date) {
  const year = String(date.getFullYear());
  const existingFolders = parentFolder.getFoldersByName(year);

  return existingFolders.hasNext() ? existingFolders.next() : parentFolder.createFolder(year);
}

/**
 * Finds an existing file with the given name in a folder
 * @param {Folder} folder - The folder to search
 * @param {string} filename - The filename to find
 * @returns {File|null} The file if found, null otherwise
 */
function findExistingFile(folder, filename) {
  const files = folder.getFilesByName(filename);
  return files.hasNext() ? files.next() : null;
}

/**
 * Handles duplicate email based on configured duplicate mode
 * @param {File} existingFile - The existing file
 * @param {string} filename - The filename for logging
 * @param {string} duplicateMode - The duplicate handling mode ('ignore' or 'overwrite')
 * @returns {Object} Result object indicating the action taken
 */
function handleDuplicateEmail(existingFile, filename, duplicateMode) {
  if (duplicateMode === CONFIG.DUPLICATE_MODES.IGNORE) {
    Logger.log(`Skipped (duplicate): ${filename}`);
    return { status: 'skippedCount' };
  }

  if (duplicateMode === CONFIG.DUPLICATE_MODES.OVERWRITE) {
    existingFile.setTrashed(true);
    Logger.log(`Overwritten (duplicate): ${filename}`);
    return { status: 'skippedCount' };
  }

  throw new Error(`Unknown duplicate mode: ${duplicateMode}`);
}

/**
 * Saves the email as an .eml file to the specified folder
 * @param {Folder} folder - The folder to save to
 * @param {string} filename - The filename
 * @param {GmailMessage} msg - The email message
 * @param {Date} date - The email date (used to set file modification time)
 * @returns {File} The created file
 */
function saveEmailToFolder(folder, filename, msg, date) {
  const file = folder.createFile(filename, msg.getRawContent(), MimeType.PLAIN_TEXT);

  // Set the file's modification time to match the email date
  Drive.Files.update(
    { modifiedTime: date.toISOString() },
    file.getId()
  );

  return file;
}

/**
 * Gets the last run timestamp from script properties
 * @returns {string} The last run timestamp or INITIAL_LAST_RUN
 */
function getLastRunTimestamp() {
  const props = PropertiesService.getScriptProperties();
  return props.getProperty(PROPS.LAST_RUN) || CONFIG.INITIAL_LAST_RUN;
}

/**
 * Updates the last run timestamp in script properties
 * @param {number} newestTimestamp - The newest email timestamp (seconds since epoch)
 */
function updateLastRunTimestamp(newestTimestamp) {
  if (newestTimestamp <= 0) {
    Logger.log('No new emails saved, lastRun timestamp not updated');
    return;
  }

  const props = PropertiesService.getScriptProperties();
  props.setProperty(PROPS.LAST_RUN, newestTimestamp);
  Logger.log(`Updated lastRun to: ${newestTimestamp}`);
}

/**
 * Logs execution summary with timing and statistics
 * @param {Date} startTime - The execution start time
 * @param {Object} stats - Statistics object with savedCount, skippedCount, errorCount
 */
function logExecutionSummary(startTime, stats) {
  const endTime = new Date();
  const duration = ((endTime - startTime) / 1000).toFixed(2);

  Logger.log(
    `=== Execution Complete ===
Saved: ${stats.savedCount} | Skipped: ${stats.skippedCount} | Errors: ${stats.errorCount}
Duration: ${duration}s`
  );
}
