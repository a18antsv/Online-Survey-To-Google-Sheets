# Online-Survey-To-Google-Sheets

## Getting started
### First time
1. Install dependencies using `npm install`
2. [Create a Google Cloud project](https://developers.google.com/workspace/guides/create-project)
3. [Enable Google Sheets API](https://developers.google.com/workspace/guides/enable-apis)
4. [Create OAuth client ID access credentials for desktop app](https://developers.google.com/workspace/guides/create-credentials#oauth-client-id)
5. Create service account (Google Sheets API -> + Create credentials -> Service account).
6. Create a new key for the service account (click on the created service account -> Keys -> Add key -> Create new key -> JSON -> Create)
7. Copy value of `client_email` from the downloaded JSON file and set it to the `SERVICE_ACCOUNT_EMAIL` environment variable in `.env`file.
8. Copy value of `private_key` and set it to the `SERVICE_ACCOUNT_PRIVATE_KEY` environment variable in `.env` file. The whole string should be copied (including -----BEGIN PRIVATE KEY-----). Since this string contains new lines, it is important to keep the new lines and without spaces before the new lines.
7. Create spreadsheet [Google Docs](https://docs.google.com/spreadsheets)
8. Copy spreadsheet id from the url, for example `1t5speBA-0VIzuMcRumRsPrUVESHQINBXxnb7Z3CHCr4` and set it to  the value of `SPREADSHEET_ID` environment variable in the `.env`file. 
9. Share the spreadsheet you are working with to the email address of the service account (for example: `service-account@onlinesurvey-project.iam.gserviceaccount.com`)
10. Build the typescript code using command `npm run build`
11. Start the script using command `npm run start`

### Run script
1. Run the script `npm run start`

### Filtering of countries
Unwanted countries can be filtered out from the script by changing the value of the `FILTER_IN` environment variable in the `.env`file.

