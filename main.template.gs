// Configuration Constants
const CONFIG = {
  DUPLICATE_MODE: '__DUPLICATE_MODE__',
  FOLDER_ID: '__FOLDER_ID__',
  INITIAL_LAST_RUN: '__INITIAL_LAST_RUN__',
  LOOKBACK_DAYS: '__LOOKBACK_DAYS__',
  SEARCH_QUERY: '__SEARCH_QUERY__',

  get LOOKBACK_SECONDS() {
    return this.LOOKBACK_DAYS * 86400;
  },

  DUPLICATE_MODES: {
    IGNORE: 'ignore',
    OVERWRITE: 'overwrite',
  },

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
const SCRIPT_PROPS = PropertiesService.getScriptProperties();
const SCRIPT_LOCK = LockService.getScriptLock();

const MIMETYPE_EMAIL = 'message/rfc822';
const MAX_EXECUTION_TIME_SECONDS = 360;  // 6 minutes

/**
 * Main entry point: Saves new emails from Gmail to Google Drive (optimized for speed)
 * Uses lazy caching and batch fetching to improve performance
 * Supports multiple folder organization granularities
 */
function saveNewEmailsToDrive() {
  try {
    const startTime = new Date();
    const startTs = startTime.getTime() / 1000;
    const stats = { savedCount: 0, skippedCount: 0, errorCount: 0, unprocessedCount: 0 };

    const folder = DriveApp.getFolderById(CONFIG.FOLDER_ID);
    const lastRunTs = getLastRunTimestamp();

    var afterTs = lastRunTs - 2;
    var beforeTs = Math.min(afterTs + CONFIG.LOOKBACK_SECONDS, startTs);

    const batchSize = 500;  // Gmail limit: max 500 threads per GmailApp.search() and GmailApp.getMessagesForThreads() call

    // try to acquire lock immediately
    if (!SCRIPT_LOCK.tryLock(0)) {
      Logger.log('Another instance is already running. Exiting.');
      return;
    }
    Logger.log(`Searching for emails between ${formatDate(afterTs)} and ${formatDate(beforeTs)}`);

    const threads = [];
    while (true) {
      const threadsBatch = GmailApp.search(`${CONFIG.SEARCH_QUERY} after:${afterTs} before:${beforeTs}`, threads.length, batchSize);
      if (threadsBatch.length === 0) {
        if (threads.length === 0) {
          Logger.log('No emails found in the specified date range.');

          if (beforeTs >= startTs) {
            Logger.log('Reached current time, nothing more to search');
            if (threads.length === 0) {
              // No emails found at all, update lastRun to current time
              SCRIPT_PROPS.setProperty(PROPS.LAST_RUN, startTs);
              Logger.log(`Updated lastRun to: ${startTs} (${formatDate(startTs)})`);
            }
            break;
          }

          // Move the window forward
          SCRIPT_PROPS.setProperty(PROPS.LAST_RUN, beforeTs);
          Logger.log(`Updated lastRun to: ${beforeTs} (${formatDate(beforeTs)})`);

          var afterTs = beforeTs - 1;
          var beforeTs = afterTs + CONFIG.LOOKBACK_SECONDS;

          Logger.log(`Searching for emails between ${formatDate(afterTs)} and ${formatDate(beforeTs)}`);

          continue;
        } else {
          // No more emails found in the specified date range
          break;
        }
      }
      threads.push(...threadsBatch);
    }
    Logger.log(`Found ${threads.length} threads`);

    const messages = [];
    for (let start = 0; start < threads.length; start += batchSize) {
      const messagesBatch = GmailApp.getMessagesForThreads(threads.slice(start, start + batchSize)).flat();

      Logger.log(`Found ${messagesBatch.length} messages in this batch`);

      messagesBatch.forEach(message => {
        const messageTs = Math.floor(message.getDate().getTime() / 1000);
        if (afterTs < messageTs && messageTs < beforeTs) {
          messages.push(message);
        }
      });

      Logger.log(`Total new messages collected so far: ${messages.length}`);
    }

    // Sort messages by date (oldest first) for deterministic processing (~5s per 1k messages)
    messages.sort((a, b) => a.getDate() - b.getDate());

    // Lazy-load cache - build only when needed
    const fileCache = {};

    // Process all messages
    let latestMessage = null;
    let messageCounter = 0;
    for (const message of messages) {
      if (getExecutionSecondsLeft(startTime) < 5) {
        Logger.log('Approaching execution time limit, stopping processing.');
        stats.unprocessedCount = messages.length - messageCounter;
        break;
      }
      try {
        messageCounter++;
        const result = processSingleEmail(message, folder, fileCache, messageCounter, messages.length);
        stats[result.status]++;
        latestMessage = message;

      } catch (error) {
        stats.errorCount++;
        Logger.log(`ERROR processing email: ${error.message}`);
        break;
      }
    }

    if (latestMessage) {
      const latestRunTs =
        latestMessage === messages[messages.length - 1]
          ? Math.min(beforeTs, startTs)
          : Math.floor(latestMessage.getDate().getTime() / 1000);

      SCRIPT_PROPS.setProperty(PROPS.LAST_RUN, latestRunTs);
      Logger.log(`Updated lastRun to: ${latestRunTs} (${formatDate(latestRunTs)})`);
    }

    SCRIPT_LOCK.releaseLock();

    logExecutionSummary(startTime, stats, messages.length);
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
 * Formats a Date or Unix timestamp (seconds) to `yyyy-MM-dd HH:mm` in script timezone.
 * @param {Date|number} date - Date object or Unix timestamp in seconds.
 * @returns {string} Formatted date string.
 */
function formatDate(date) {
  if (typeof date === 'number') {
    date = new Date(date * 1000);
  }
  return Utilities.formatDate(date, SCRIPT_TIMEZONE, 'yyyy-MM-dd HH:mm:ss');
}

/**
 * Returns the remaining execution time for the current script run.
 * @param {Date} startTime - The execution start time
 * @returns {number} Seconds left in the execution window
 */
function getExecutionSecondsLeft(startTime) {
  const now = new Date();
  const diff = (now.getTime() - startTime.getTime()) / 1000;
  return MAX_EXECUTION_TIME_SECONDS - diff;
}

/**
 * Builds cache for a single folder path (lazy-loaded on demand)
 * Only called when processing emails from a new folder
 * @param {Folder} rootFolder - The root Google Drive folder
 * @param {string} path - Path like '2025/01' or '2025/20250315'
 * @param {Object} fileCache - Cache object to populate
 */
function buildCacheForPath(rootFolder, path, fileCache) {
  const targetFolder = navigateToFolderPath(rootFolder, path);
  fileCache[path] = {};

  if (targetFolder) {
    const files = targetFolder.getFiles();
    while (files.hasNext()) {
      const file = files.next();
      fileCache[path][file.getName()] = file;
    }
  }
  // If targetFolder doesn't exist, fileCache[path] remains empty {}
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
 * Uses lazy cache for faster file lookups
 * @param {GmailMessage} message - The email message to process
 * @param {Folder} folder - The root Google Drive folder
 * @param {Object} fileCache - Cached file structure (lazy-loaded)
 * @param {number} messageCounter - Current message number
 * @param {number} totalMessages - Total messages to process
 * @returns {Object} Result object with status and timestamp
 */
function processSingleEmail(message, folder, fileCache, messageCounter, totalMessages) {
  const date = message.getDate();
  const subject = message.getSubject();
  const filename = generateEmailFilename(date, subject);
  const folderPath = getFolderPath(date);

  // Lazy-load cache for this path if not already cached
  if (!fileCache[folderPath]) {
    buildCacheForPath(folder, folderPath, fileCache);
  }

  // Check cache first (much faster than Drive API calls)
  if (fileCache[folderPath][filename]) {
    Logger.log(`[DUPLICATE] ${messageCounter}/${totalMessages} ${filename}`);
    return handleDuplicateEmail(fileCache[folderPath][filename], filename, CONFIG.DUPLICATE_MODE);
  }

  // Get or create folder (not in cache on first run)
  const targetFolder = getOrCreateFolderPath(folder, folderPath);
  const savedFile = saveEmailToFolder(targetFolder, filename, message, date);

  // Update cache for future operations
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
 * @param {GmailMessage} message - The email message
 * @param {Date} date - The email date
 * @returns {File} The created file
 */
function saveEmailToFolder(folder, filename, message, date) {
  const file = folder.createFile(filename, message.getRawContent(), MIMETYPE_EMAIL);

  // Set file metadata to match the email's received date
  Drive.Files.update(
    { modifiedTime: date.toISOString() },
    file.getId()
  );

  return file;
}

/**
 * Gets the last run timestamp from script properties
 * Converts date format YYYY/MM/DD to Unix timestamp
 * Supports both formats: 'YYYY/MM/DD' or Unix timestamp string
 * @returns {number} Unix timestamp in seconds
 */
function getLastRunTimestamp() {
  let lastRun = SCRIPT_PROPS.getProperty(PROPS.LAST_RUN) || CONFIG.INITIAL_LAST_RUN;

  // Convert date format YYYY/MM/DD to Unix timestamp
  // Example: '2004/04/01' -> 1080795600
  if (typeof lastRun === 'string' && lastRun.includes('/')) {
    const [year, month, day] = lastRun.split('/');
    const date = new Date(parseInt(year), parseInt(month) - 1, parseInt(day));
    return Math.floor(date.getTime() / 1000);
  }

  // If already a timestamp (number or string number), return as-is
  return Number(lastRun);
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
  const avgTime = (duration / Math.max(totalMessages - stats.unprocessedCount, 1)).toFixed(3);

  Logger.log(
    `=== Execution Complete ===
Saved: ${stats.savedCount} | Skipped: ${stats.skippedCount} | Errors: ${stats.errorCount} | Unprocessed: ${stats.unprocessedCount}
Total: ${totalMessages} messages processed
Duration: ${duration}s | Avg: ${avgTime}s/msg
Granularity: ${CONFIG.GRANULARITY}`
  );
}
