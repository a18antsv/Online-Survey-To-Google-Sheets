# Online-Survey-To-Google-Sheets

## Getting started
### First time
1. [Create a Google Cloud project](https://developers.google.com/workspace/guides/create-project)
2. [Enable Google Sheets API](https://developers.google.com/workspace/guides/enable-apis)
3. [Create OAuth client ID access credentials for desktop app](https://developers.google.com/workspace/guides/create-credentials#oauth-client-id)
4. Create service account (Google Sheets API -> + Create credentials -> Service account).
5. Create a new key for the service account (click on the created service account -> Keys -> Add key -> Create new key -> JSON -> Create)
6. Copy value of `client_email` from the downloaded JSON file and set it to the `SERVICE_ACCOUNT_EMAIL` environment variable in `.env`file.
7. Copy value of `private_key` and set it to the `SERVICE_ACCOUNT_PRIVATE_KEY` environment variable in `.env` file. The whole string should be copied (including -----BEGIN PRIVATE KEY-----). Since this string contains new lines, it is important to keep the new lines and without spaces before the new lines.
8. Create spreadsheet [Google Docs](https://docs.google.com/spreadsheets)
9. Copy spreadsheet id from the url, for example `1t5speBA-0VIzuMcRumRsPrUVESHQINBXxnb7Z3CHCr4` and set it to  the value of `SPREADSHEET_ID` environment variable in the `.env`file. 
10. Share the spreadsheet you are working with to the email address of the service account (for example: `service-account@onlinesurvey-project.iam.gserviceaccount.com`)
11. Install dependencies using `npm install`
12. Build the typescript code `npm run build`
13. Start the script `npm run start`

### Run script
1. Start the script `npm run start`

### Filtering of countries
Unwanted countries can be filtered out from the script by changing the value of the `FILTER_IN` environment variable in the `.env`file.

