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
  BATCH_SIZE: 50,
  CACHE_FOLDER_STRUCTURE: true,
};

// Property Keys
const PROPS = {
  LAST_RUN: 'lastRun',
};

/**
 * Main entry point: Saves new emails from Gmail to Google Drive (optimized for speed)
 * Uses batch processing and caching to improve performance
 */
function saveNewEmailsToDrive() {
  try {
    const startTime = new Date();
    const stats = { savedCount: 0, skippedCount: 0, errorCount: 0 };

    const folder = DriveApp.getFolderById(CONFIG.FOLDER_ID);
    const lastRun = getLastRunTimestamp();
    let newestTimestamp = 0;

    // Build a cache of existing files keyed by year and filename for O(1) lookup
    const fileCache = buildFileCache(folder);

    const threads = GmailApp.search(`after:${lastRun}`);

    // Collect all messages first, then process in batches
    const allMessages = [];
    threads.forEach(thread => {
      thread.getMessages().forEach(msg => {
        allMessages.push(msg);
      });
    });

    // Process messages in batches
    const batchCount = Math.ceil(allMessages.length / CONFIG.BATCH_SIZE);
    for (let i = 0; i < batchCount; i++) {
      const start = i * CONFIG.BATCH_SIZE;
      const end = Math.min(start + CONFIG.BATCH_SIZE, allMessages.length);
      const batch = allMessages.slice(start, end);

      batch.forEach(msg => {
        try {
          const result = processSingleEmail(msg, folder, fileCache);
          stats[result.status]++;

          if (result.timestamp && result.timestamp > newestTimestamp) {
            newestTimestamp = result.timestamp;
          }
        } catch (error) {
          stats.errorCount++;
          Logger.log(`ERROR processing email: ${error.message}`);
        }
      });

      // Log progress every batch
      Logger.log(`Progress: Processed ${Math.min(end, allMessages.length)}/${allMessages.length} messages`);
    }

    updateLastRunTimestamp(newestTimestamp);
    logExecutionSummary(startTime, stats, allMessages.length);
  } catch (error) {
    Logger.log(`CRITICAL ERROR: ${error.message}`);
    throw error;
  }
}

/**
 * Builds a cache of existing files by year and filename for fast lookups
 * @param {Folder} rootFolder - The root Google Drive folder
 * @returns {Object} Cache structure: { year: { filename: file } }
 */
function buildFileCache(rootFolder) {
  const cache = {};

  // Get all year folders
  const yearFolders = rootFolder.getFolders();
  while (yearFolders.hasNext()) {
    const yearFolder = yearFolders.next();
    const year = yearFolder.getName();
    cache[year] = {};

    // Cache all files in this year folder
    const files = yearFolder.getFiles();
    while (files.hasNext()) {
      const file = files.next();
      cache[year][file.getName()] = file;
    }
  }

  return cache;
}

/**
 * Processes a single email: validates, checks for duplicates, and saves to Drive
 * Uses cache for faster file lookups
 * @param {GmailMessage} msg - The email message to process
 * @param {Folder} folder - The root Google Drive folder
 * @param {Object} fileCache - Cached file structure
 * @returns {Object} Result object with status and timestamp
 */
function processSingleEmail(msg, folder, fileCache) {
  const date = msg.getDate();
  const subject = msg.getSubject();
  const filename = generateEmailFilename(msg, date);
  const year = String(date.getFullYear());

  // Check cache first (much faster than Drive API calls)
  if (fileCache[year] && fileCache[year][filename]) {
    // Skip duplicate logging to avoid performance overhead (~1 second per 100 emails)
    // Each Logger.log() call costs 10-15ms;
    // Uncomment the line below if you need to see which emails are duplicates (for debugging)
    //Logger.log(`[DUPLICATE] ${date.toISOString()} - ${subject}`);
    return handleDuplicateEmail(fileCache[year][filename], filename, CONFIG.DUPLICATE_MODE);
  }

  // Get or create folder (not in cache on first run)
  const yearFolder = getOrCreateYearFolder(folder, date);
  const savedFile = saveEmailToFolder(yearFolder, filename, msg, date);

  // Update cache for future operations
  if (!fileCache[year]) {
    fileCache[year] = {};
  }
  fileCache[year][filename] = savedFile;

  // Log newly saved emails only (much faster than logging all)
  Logger.log(`[SAVED] ${date.toISOString()} - ${subject}`);

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
  const subject = msg.getSubject();
  const sanitized = `${date.toISOString()} - ${subject}`.replace(/[\/\\?%*:|"<>]/g, '_');
  return `${sanitized}.eml`;
}

/**
 * Gets or creates a year-based folder in Google Drive
 * Faster implementation with early return
 * @param {Folder} parentFolder - The parent folder
 * @param {Date} date - The date used to determine the year
 * @returns {Folder} The year folder
 */
function getOrCreateYearFolder(parentFolder, date) {
  const year = String(date.getFullYear());
  const existingFolders = parentFolder.getFoldersByName(year);

  if (existingFolders.hasNext()) {
    return existingFolders.next();
  }

  return parentFolder.createFolder(year);
}

/**
 * Handles duplicate email based on configured duplicate mode
 * @param {File} existingFile - The existing file
 * @param {string} filename - The filename for logging
 * @param {string} duplicateMode - The duplicate handling mode
 * @returns {Object} Result object indicating the action taken
 */
function handleDuplicateEmail(existingFile, filename, duplicateMode) {
  if (duplicateMode === CONFIG.DUPLICATE_MODES.IGNORE) {
    return { status: 'skippedCount' };
  }

  if (duplicateMode === CONFIG.DUPLICATE_MODES.OVERWRITE) {
    existingFile.setTrashed(true);
    return { status: 'skippedCount' };
  }

  throw new Error(`Unknown duplicate mode: ${duplicateMode}`);
}

/**
 * Saves the email as an .eml file to the specified folder
 * @param {Folder} folder - The folder to save to
 * @param {string} filename - The filename
 * @param {GmailMessage} msg - The email message
 * @param {Date} date - The email date
 * @returns {File} The created file
 */
function saveEmailToFolder(folder, filename, msg, date) {
  const file = folder.createFile(filename, msg.getRawContent(), MimeType.PLAIN_TEXT);

  // Set file metadata
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
 * @param {number} newestTimestamp - The newest email timestamp
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
 * @param {number} totalMessages - Total messages processed
 */
function logExecutionSummary(startTime, stats, totalMessages) {
  const endTime = new Date();
  const duration = ((endTime - startTime) / 1000).toFixed(2);
  const avgTime = (duration / Math.max(totalMessages, 1)).toFixed(3);

  Logger.log(
    `=== Execution Complete ===
Saved: ${stats.savedCount} | Skipped: ${stats.skippedCount} | Errors: ${stats.errorCount}
Total: ${totalMessages} messages processed
Duration: ${duration}s | Avg: ${avgTime}s/msg`
  );
}
