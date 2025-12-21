# gmail-to-drive-archiver

Archive Gmail messages to Google Drive using Google Apps Script.

## Build

```bash
cp .env.example .env
make build
```

## Deploy

1. Open https://script.google.com
2. Create a new project
3. Copy `main.gs` into the editor
4. Save and authorize
5. _(Optional)_ Add a time-based trigger

## Usage

Run `saveNewEmailsToDrive` manually, or set up a time-based trigger to execute automatically.
