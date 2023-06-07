import {google, sheets_v4} from "googleapis";
import {SheetData} from "./index.js";
import {handler} from "./util.js";
import * as console from "console";
import Schema$Request = sheets_v4.Schema$Request;
import Schema$Spreadsheet = sheets_v4.Schema$Spreadsheet;
import Schema$ValueRange = sheets_v4.Schema$ValueRange;


const {
    SPREADSHEET_ID,
    SERVICE_ACCOUNT_EMAIL,
    SERVICE_ACCOUNT_PRIVATE_KEY,
} = process.env;

const api = google.sheets({
    version: 'v4',
    auth: new google.auth.JWT({
        email: SERVICE_ACCOUNT_EMAIL,
        key: SERVICE_ACCOUNT_PRIVATE_KEY,
        scopes: ["https://www.googleapis.com/auth/spreadsheets"]
    }),
});

const getSpreadsheet = async (): Promise<Schema$Spreadsheet> => {
    const {response, error} = await handler(api.spreadsheets.get({
        spreadsheetId: SPREADSHEET_ID
    }));
    if (error) throw new Error(`Could not get spreadsheet by id '${SPREADSHEET_ID}'`, error);

    return response!.data;
}

const addSheets = async (sheetNames: string[], spreadsheet: Schema$Spreadsheet): Promise<Schema$Spreadsheet> => {
    const availableSheetNames = spreadsheet.sheets?.map(sheet => sheet.properties?.title) ?? [];
    const addSheetRequests = sheetNames
        .filter(sheetName => !availableSheetNames.includes(sheetName))
        .map<Schema$Request>(sheetName => ({
            addSheet: {
                properties: {
                    title: sheetName,
                }
            },
        } satisfies Schema$Request));

    if (addSheetRequests.length === 0) return spreadsheet;

    console.log('Adding sheets...');

    const {response, error} = await handler(api.spreadsheets.batchUpdate({
        spreadsheetId: SPREADSHEET_ID,
        requestBody: {
            requests: addSheetRequests,
            includeSpreadsheetInResponse: true
        }
    }));
    if (error) throw new Error('Could not add sheets.', {cause: error});

    return response!.data.updatedSpreadsheet!;
}

const getSheetIdByName = (sheetNames: string[], spreadsheet: Schema$Spreadsheet): Map<string, number | null | undefined> => {
    return new Map(sheetNames.map(sheetName => {
        const sheetId = spreadsheet.sheets?.find(sheet => sheet.properties?.title === sheetName)?.properties?.sheetId;
        return [sheetName, sheetId];
    }));
}

const getValueRequests = (sheets: SheetData[]): Schema$ValueRange[] => {
    return sheets.flatMap(sheet => {
        return sheet.tables.map(table => ({
            range: `'${sheet.sheetName}'!${table.area.toString()}`,
            values: table.rows,
        } satisfies Schema$ValueRange))
    });
}

const getMergeRequests = (sheets: SheetData[], sheetIdByName: Map<string, number | null | undefined>): Schema$Request[] => {
    return sheets.flatMap(sheet => {
        const sheetId = sheetIdByName.get(sheet.sheetName);
        return sheet.tables
            .flatMap(table => table.mergeRequests ?? [])
            .map(mergeRequest => {
                mergeRequest.range = {
                    ...mergeRequest.range,
                    sheetId
                };

                return {mergeCells: mergeRequest} satisfies Schema$Request;
            });
    });
}

const getUpdateBorderRequests = (sheets: SheetData[], sheetIdByName: Map<string, number | null | undefined>): Schema$Request[] => {
    return sheets.flatMap(sheet => {
        const sheetId = sheetIdByName.get(sheet.sheetName);
        return sheet.tables
            .flatMap(table => table.updateBorderRequests ?? [])
            .map(request => {
                request.range = {
                    ...request.range,
                    sheetId,
                };

                return {updateBorders: request} satisfies Schema$Request;
            });
    })
}

export const updateSheets = async (sheets: SheetData[]): Promise<void> => {
    console.log('Updating sheets...');

    let spreadsheet = await getSpreadsheet();
    const sheetNames = sheets.map(({sheetName}) => sheetName);
    spreadsheet = await addSheets(sheetNames, spreadsheet);

    const sheetIdByName = getSheetIdByName(sheetNames, spreadsheet);
    const valueRequests = getValueRequests(sheets);
    const mergeRequests = getMergeRequests(sheets, sheetIdByName);
    const borderRequests = getUpdateBorderRequests(sheets, sheetIdByName);

    const {error: valueRequestError} = await handler(api.spreadsheets.values.batchUpdate({
        spreadsheetId: SPREADSHEET_ID,
        requestBody: {
            valueInputOption: 'USER_ENTERED',
            data: valueRequests,
        }
    }));
    if (valueRequestError) throw new Error('Value request error', {cause: valueRequestError});

    const {error: formatRequestError} = await handler(api.spreadsheets.batchUpdate({
        spreadsheetId: SPREADSHEET_ID,
        requestBody: {
            requests: [...mergeRequests, ...borderRequests],
        }
    }));
    if (formatRequestError) throw new Error('Format request error', {cause: formatRequestError});

    console.log('Updated tabs successfully!');
}