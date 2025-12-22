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

  // Folder granularity: 'yearly', 'monthly', or 'daily'
  // yearly:  folder/YYYY/email.eml
  // monthly: folder/YYYY/MM/email.eml
  // daily:   folder/YYYY/YYYYMMDD/email.eml
  GRANULARITY: '__GRANULARITY__',
};

// Property Keys
const PROPS = {
  LAST_RUN: 'lastRun',
};

// Script-level Constants (cached at startup)
const SCRIPT_TIMEZONE = Session.getScriptTimeZone();

/**
 * Main entry point: Saves new emails from Gmail to Google Drive (optimized for speed)
 * Uses batch processing and caching to improve performance
 * Supports multiple folder organization granularities
 */
function saveNewEmailsToDrive() {
  try {
    const startTime = new Date();
    const stats = { savedCount: 0, skippedCount: 0, errorCount: 0 };

    const folder = DriveApp.getFolderById(CONFIG.FOLDER_ID);
    const lastRun = getLastRunTimestamp();
    const lastRunTimestamp = parseInt(lastRun);
    let newestTimestamp = 0;

    const threads = GmailApp.search(`${CONFIG.SEARCH_QUERY} after:${lastRun}`);
    const allMessages = GmailApp.getMessagesForThreads(threads);

    // Collect new messages
    const newMessages = [];

    allMessages.forEach(msg => {
      if (Math.floor(msg.getDate().getTime() / 1000) > lastRunTimestamp) {
        newMessages.push(msg);
      }
    });

    // Determine which folder paths we actually need to cache based on granularity
    const pathsNeeded = new Set();
    newMessages.forEach(msg => {
      const path = getFolderPath(msg.getDate());
      pathsNeeded.add(path);
    });

    // Build cache only for the folder paths we need
    const fileCache = buildFileCacheForPaths(folder, pathsNeeded);

    // Process all messages
    let messageCounter = 0;
    newMessages.forEach(msg => {
      try {
        messageCounter++;
        const result = processSingleEmail(msg, folder, fileCache, messageCounter, newMessages.length);
        stats[result.status]++;

        if (result.timestamp && result.timestamp > newestTimestamp) {
          newestTimestamp = result.timestamp;
        }
      } catch (error) {
        stats.errorCount++;
        Logger.log(`ERROR processing email: ${error.message}`);
      }
    });

    updateLastRunTimestamp(newestTimestamp);
    logExecutionSummary(startTime, stats, newMessages.length);
  } catch (error) {
    Logger.log(`CRITICAL ERROR: ${error.message}`);
    throw error;
  }
}

/**
 * Generates the folder path based on date and configured granularity
 * @param {Date} date - The email date
 * @returns {string} Folder path (e.g., '2025', '2025/01', '2025/20250315')
 */
function getFolderPath(date) {
  const year = String(date.getFullYear());

  switch (CONFIG.GRANULARITY) {
    case 'yearly':
      return year;
    case 'monthly':
      const month = String(date.getMonth() + 1).padStart(2, '0');
      return `${year}/${month}`;
    case 'daily':
      const month_d = String(date.getMonth() + 1).padStart(2, '0');
      const day = String(date.getDate()).padStart(2, '0');
      return `${year}/${year}${month_d}${day}`;
    default:
      throw new Error(`Unknown granularity: ${CONFIG.GRANULARITY}`);
  }
}

/**
 * Builds a cache of existing files ONLY for specified folder paths
 * Much faster than caching all folders when you only need a few
 * @param {Folder} rootFolder - The root Google Drive folder
 * @param {Set} pathsNeeded - Set of folder paths to cache (e.g., {'2025/01', '2025/20250215'})
 * @returns {Object} Cache structure: { 'path': { filename: file } }
 */
function buildFileCacheForPaths(rootFolder, pathsNeeded) {
  const cache = {};

  // For each path needed, navigate to it and cache files
  pathsNeeded.forEach(path => {
    try {
      const targetFolder = navigateToFolderPath(rootFolder, path);
      if (targetFolder) {
        cache[path] = {};

        const files = targetFolder.getFiles();
        while (files.hasNext()) {
          const file = files.next();
          cache[path][file.getName()] = file;
        }
      }
    } catch (error) {
      // Folder doesn't exist yet, will be created when saving
      cache[path] = {};
    }
  });

  return cache;
}

/**
 * Navigates to a folder path, returns null if path doesn't exist
 * @param {Folder} rootFolder - Starting folder
 * @param {string} path - Path like '2025/01' or '2025/20250315'
 * @returns {Folder|null} The target folder or null if not found
 */
function navigateToFolderPath(rootFolder, path) {
  const parts = path.split('/');
  let currentFolder = rootFolder;

  for (const part of parts) {
    const folders = currentFolder.getFoldersByName(part);
    if (folders.hasNext()) {
      currentFolder = folders.next();
    } else {
      return null; // Path doesn't exist yet
    }
  }

  return currentFolder;
}

/**
 * Processes a single email: validates, checks for duplicates, and saves to Drive
 * Uses cache for faster file lookups
 * @param {GmailMessage} msg - The email message to process
 * @param {Folder} folder - The root Google Drive folder
 * @param {Object} fileCache - Cached file structure
 * @param {number} messageCounter - Current message number
 * @param {number} totalMessages - Total messages to process
 * @returns {Object} Result object with status and timestamp
 */
function processSingleEmail(msg, folder, fileCache, messageCounter, totalMessages) {
  const date = msg.getDate();
  const subject = msg.getSubject();
  const filename = generateEmailFilename(date, subject);
  const folderPath = getFolderPath(date);

  // Check cache first (much faster than Drive API calls)
  if (fileCache[folderPath] && fileCache[folderPath][filename]) {
    // Skip duplicate logging to avoid performance overhead (~1 second per 100 emails)
    // Each Logger.log() call costs 10-15ms;
    // Uncomment the line below if you need to see which emails are duplicates (for debugging)
    //Logger.log(`[DUPLICATE] ${filename}`);
    return handleDuplicateEmail(fileCache[folderPath][filename], filename, CONFIG.DUPLICATE_MODE);
  }

  // Get or create folder (not in cache on first run)
  const targetFolder = getOrCreateFolderPath(folder, folderPath);
  const savedFile = saveEmailToFolder(targetFolder, filename, msg, date);

  // Update cache for future operations
  if (!fileCache[folderPath]) {
    fileCache[folderPath] = {};
  }
  fileCache[folderPath][filename] = savedFile;

  // Log newly saved emails with progress counter
  Logger.log(`[SAVED] ${messageCounter}/${totalMessages} ${filename}`);

  return {
    status: 'savedCount',
    timestamp: Math.floor(date.getTime() / 1000),
  };
}

/**
 * Gets or creates a folder path, creating intermediate folders as needed
 * @param {Folder} rootFolder - The root folder
 * @param {string} path - Path like '2025/01' or '2025/20250315'
 * @returns {Folder} The target folder
 */
function getOrCreateFolderPath(rootFolder, path) {
  const parts = path.split('/');
  let currentFolder = rootFolder;

  for (const part of parts) {
    const folders = currentFolder.getFoldersByName(part);
    if (folders.hasNext()) {
      currentFolder = folders.next();
    } else {
      currentFolder = currentFolder.createFolder(part);
    }
  }

  return currentFolder;
}

/**
 * Generates a sanitized filename for the email using local timezone
 * Format: YYYY-MM-DDTHH_MM_SS Subject.eml
 * @param {Date} date - The email date (in local timezone)
 * @param {string} subject - The email subject
 * @returns {string} Sanitized filename with .eml extension
 */
function generateEmailFilename(date, subject) {
  // Format date using Utilities.formatDate() for local timezone
  // Format: YYYY-MM-DDTHH_MM_SS
  const timestamp = Utilities.formatDate(date, SCRIPT_TIMEZONE, "yyyy-MM-dd'T'HH_mm_ss");

  // Sanitize subject line and remove problematic characters
  const sanitizedSubject = subject.replace(/[\/\\?%*:|"<>]/g, '_');

  // Combine timestamp and subject
  return `${timestamp} ${sanitizedSubject}.eml`;
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

  // Set file metadata to match the email's received date
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
Duration: ${duration}s | Avg: ${avgTime}s/msg
Granularity: ${CONFIG.GRANULARITY}`
  );
}
