const DUPLICATE_MODE = '__DUPLICATE_MODE__';
const FOLDER_ID = '__FOLDER_ID__';
const INITIAL_LAST_RUN = '__INITIAL_LAST_RUN__';
const SEARCH_QUERY = '__SEARCH_QUERY__';


function saveNewEmailsToDrive() {
  const startTime = new Date();
  const folder = DriveApp.getFolderById(FOLDER_ID);

  const props = PropertiesService.getScriptProperties();
  const lastRun = props.getProperty('lastRun') || INITIAL_LAST_RUN;

  let newestTimestamp = 0;
  let savedCount = 0;

  const threads = GmailApp.search(`after:${lastRun}`);
  threads.forEach(thread => {
    thread.getMessages().forEach(msg => {
      const date = msg.getDate();
      const year = String(date.getFullYear());
      const filename = `${date.toISOString()} - ${msg.getSubject()}.eml`
        .replace(/[\/\\?%*:|"<>]/g, '_');

      const folders = folder.getFoldersByName(year);
      const yearFolder = folders.hasNext()
        ? folders.next()
        : folder.createFolder(year);

      const existing = yearFolder.getFilesByName(filename);

      if (existing.hasNext()) {
        if (DUPLICATE_MODE === 'ignore') {
          Logger.log(`Skipped (duplicate): ${filename}`);
          return;
        }
        if (DUPLICATE_MODE === 'overwrite') {
          while (existing.hasNext()) {
            existing.next().setTrashed(true);
          }
        }
      }

      // Save email as .eml
      file = yearFolder.createFile(filename, msg.getRawContent(), MimeType.PLAIN_TEXT);
      Drive.Files.update(
        { modifiedTime: date.toISOString() },
        file.getId()
      );

      savedCount++;
      const ts = Math.floor(date.getTime() / 1000);
      if (ts > newestTimestamp) {
        newestTimestamp = ts;
      }

      Logger.log(`Saved email: ${filename}}`);

    });
  });

  if (newestTimestamp > 0) {
    props.setProperty('lastRun', newestTimestamp);
    Logger.log(`Updated lastRun to: ${newestTimestamp}`);
  } else {
    Logger.log('No new emails saved');
  }
  const endTime = new Date();
  const duration = (endTime - startTime) / 1000;
  Logger.log(`=== Finished. Emails saved: ${savedCount}. Execution time: ${duration} sec ===`);
}
