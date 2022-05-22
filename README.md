# Online-Survey-To-Google-Sheets

## Getting started
### First time
1. Install dependencies using `npm install`
2. [Create a Google Cloud project](https://developers.google.com/workspace/guides/create-project)
3. [Enable Google Sheets API](https://developers.google.com/workspace/guides/enable-apis)
4. [Create OAuth client ID access credentials for desktop app](https://developers.google.com/workspace/guides/create-credentials#oauth-client-id)
5. Download OAuth client json file from APIs and services > Credentials > OAuth 2.0 Client IDs.
6. Rename downloaded file to `credentials.json` and place it in the project folder.
7. Create spreadsheet [Google Docs](https://docs.google.com/spreadsheets)
8. Copy spreadsheet id from the url, for example 1t5speBA-0VIzuMcRumRsPrUVESHQINBXxnb7Z3CHCr4
9. Set the value of the `SPREADSHEET_ID` environment variable in the `.env` file to the copied spreadsheet id.

### First time and occasionally
1. On the first run, a message will pop up in the console that asks for authorization: `Authorize this app by visiting this url: ...`. Follow the link and log into the Google account where the Google Sheets API was enabled.
2. Copy the code that was generated on log in and enter it in the console when prompted with: `Enter the code from that page here:`.
3. A file called `token.json` will be generated and stored in the current working directory. This token will expire occasionally, when it does delete the file and run the script to redo this process.

### Run script
1. Run the script `node index.js`

### Filtering of countries
Unwanted countries can be filtered out from the script by changing the value of the `FILTER_IN` environment variable in the `.env`file.

