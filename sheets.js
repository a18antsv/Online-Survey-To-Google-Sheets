import {promiseHandler as handler} from "./util.js";
import { styles } from "./style.js";
import {google} from "googleapis";

const {
    SPREADSHEET_ID,
} = process.env;

const CAPITAL_ASCII_START = 65;

const getAddTabsRequests = (spreadsheet, countryDataByKey) => {
    return Object.values(countryDataByKey).reduce((current, {tab}) => {
        const tabExists = spreadsheet.data.sheets.some(sheet => sheet.properties.title === tab);
        if (tabExists) return current;

        return [
            ...current,
            {
                addSheet: {
                    properties: {
                        title: tab
                    }
                }
            }
        ];
    }, []);
}


const addTabs = async (api, addTabRequests) => {
    if (addTabRequests.length === 0) return 0;

    console.log("Adding tabs...");
    const [sheetUpdateError] = await handler(api.spreadsheets.batchUpdate({
        spreadsheetId: SPREADSHEET_ID,
        resource: {
            requests: addTabRequests
        }
    }));
    if (sheetUpdateError) return console.error(sheetUpdateError);
    console.log("Finished adding tabs!");

    return addTabRequests.length;
}

const indexToChar = index => String.fromCharCode(CAPITAL_ASCII_START - 1 + index);

const getRange = (rowStartIndex, columnStartIndex, rows) => {
    const rowEndIndex = rows.length === 0 ? rowStartIndex : rowStartIndex - 1 + rows.length;
    const columnEndIndex = columnStartIndex - 1 + (rows?.[0]?.length ?? 1);
    const startColumn = indexToChar(columnStartIndex);
    const endColumn = indexToChar(columnEndIndex);

    return {
        row: {
            start: rowStartIndex,
            end: rowEndIndex,
        },
        column: {
            start: columnStartIndex,
            end: columnEndIndex,
        },
        toString: () => `${startColumn}${rowStartIndex}:${endColumn}${rowEndIndex}`,
    }
}

const getStyleRequest = (sheetId, range) => {
    return [
        {
            repeatCell: {
                range: {
                    sheetId,
                    startRowIndex: range.row.start - 1,
                    endRowIndex: range.row.start,
                    startColumnIndex: range.column.start - 1,
                    endColumnIndex: range.column.end
                },
                ...styles.header
            },
        },
        {
            repeatCell: {
                range: {
                    sheetId,
                    startRowIndex: range.row.start,
                    endRowIndex: range.row.end - 1,
                    startColumnIndex: range.column.start - 1,
                    endColumnIndex: range.column.start
                },
                ...styles.firstColumn
            },
        },
        {
            repeatCell: {
                range: {
                    sheetId,
                    startRowIndex: range.row.end - 1,
                    endRowIndex: range.row.end,
                    startColumnIndex: range.column.start - 1,
                    endColumnIndex: range.column.end
                },
                ...styles.footer
            },
        },
        {
            updateBorders: {
                range: {
                    sheetId,
                    startRowIndex: range.row.start - 1,
                    endRowIndex: range.row.end,
                    startColumnIndex: range.column.start - 1,
                    endColumnIndex: range.column.end
                },
                ...styles.borders
            }
        }
    ];
}

const getValueRequest = (tab, range, table) => {
    return {
        range: `${tab}!${range.toString()}`,
        values: table
    }
}

export const updateTabs = async (auth, countryDataByKey) => {
    console.log("Updating tabs...");

    const api = google.sheets({version: 'v4', auth});
    let [getSpreadsheetError, spreadsheet] = await handler(api.spreadsheets.get({
        spreadsheetId: SPREADSHEET_ID
    }));
    if (getSpreadsheetError) return console.error(`Could not get spreadsheet by id ${SPREADSHEET_ID}`);

    const addTabRequests = getAddTabsRequests(spreadsheet, countryDataByKey);
    const numberOfAddedTabs = await addTabs(api, addTabRequests);
    if (numberOfAddedTabs > 0) {
        [getSpreadsheetError, spreadsheet] = await handler(api.spreadsheets.get({
            spreadsheetId: SPREADSHEET_ID
        }));
        if (getSpreadsheetError) return console.error(`Could not get spreadsheet by id ${SPREADSHEET_ID}`);
    }

    const requests = Object.values(countryDataByKey).reduce((current, {tab, tables}) => {
        const sheetId = spreadsheet.data.sheets.find(object => tab === object.properties.title).properties.sheetId;

        const genderRange = getRange(1, 1, tables["gender"]);
        const ageRange = getRange(genderRange.row.end + 2, 1, tables["age"]);
        const regionRange = getRange(ageRange.row.end + 2, 1, tables["region"])
        const segmentRange = getRange(regionRange.row.end + 3, 1, tables["mainSegment"]);
        const brandRange = getRange(1, genderRange.column.end + 2, tables["brand"]);
        const brandByOwnerRange = getRange(segmentRange.row.end + 2, 1, tables["ownerBrand"]);
        const segmentBoosterRange = getRange(brandByOwnerRange.row.end + 2, 1, tables["segmentBooster"]);
        const evBoosterRange = getRange(segmentBoosterRange.row.end + 2, 1, tables["evBooster"]);

        const valueRequests = [
            getValueRequest(tab, genderRange, tables["gender"]),
            getValueRequest(tab, ageRange, tables["age"]),
            getValueRequest(tab, brandRange, tables["brand"]),
            getValueRequest(tab, regionRange, tables["region"]),
            getValueRequest(tab, segmentRange, tables["mainSegment"]),
            getValueRequest(tab, brandByOwnerRange, tables["ownerBrand"]),
            getValueRequest(tab, segmentBoosterRange, tables["segmentBooster"]),
            getValueRequest(tab, evBoosterRange, tables["evBooster"]),
            {
                range: `${tab}!F${segmentRange.row.start - 1}`,
                values: [["Owner"]]
            },
            {
                range: `${tab}!J${segmentRange.row.start - 1}`,
                values: [["Intender"]]
            },
        ];

        const mergeRequests = [
            {
                mergeCells: {
                    range: {
                        sheetId,
                        startRowIndex: segmentRange.row.start - 2,
                        endRowIndex: segmentRange.row.start - 1,
                        startColumnIndex: 5,
                        endColumnIndex: 9
                    },
                    mergeType: "MERGE_ROWS"
                }
            },
            {
                mergeCells: {
                    range: {
                        sheetId,
                        startRowIndex: segmentRange.row.start - 2,
                        endRowIndex: segmentRange.row.start - 1,
                        startColumnIndex: 9,
                        endColumnIndex: 13
                    },
                    mergeType: "MERGE_ROWS"
                }
            },
/*             {
                mergeCells: {
                    range: {
                        sheetId,
                        startRowIndex: segmentRange.row.start - 2,
                        endRowIndex: segmentRange.row.start,
                        startColumnIndex: 0,
                        endColumnIndex: 5
                    },
                    mergeType: "MERGE_COLUMNS"
                }
            } */
        ];

        const styleRequests = [
/*             {
                repeatCell: {
                    range: {
                        sheetId: sheetId,
                        startRowIndex: segmentRange.row.start - 2,
                        endRowIndex: segmentRange.row.start,
                        startColumnIndex: 0,
                        endColumnIndex: 5
                    },
                    cell: {
                        userEnteredFormat: {
                            verticalAlignment: "MIDDLE"
                        }
                    },
                    fields: "userEnteredFormat.verticalAlignment"
                }
            }, */
            ...getStyleRequest(sheetId, genderRange),
            ...getStyleRequest(sheetId, ageRange),
            ...getStyleRequest(sheetId, regionRange),
            ...getStyleRequest(sheetId, segmentRange),
            ...getStyleRequest(sheetId, brandRange),
            ...getStyleRequest(sheetId, brandByOwnerRange),
            ...getStyleRequest(sheetId, segmentBoosterRange),
            ...getStyleRequest(sheetId, evBoosterRange),
        ];

        return {
            valueRequests: [current.valueRequests, ...valueRequests],
            mergeRequests: [current.mergeRequests, ...mergeRequests],
            styleRequests: [current.styleRequests, ...styleRequests],
        }
    }, {});

    const [valueUpdateError] = await handler(api.spreadsheets.values.batchUpdate({
        spreadsheetId: SPREADSHEET_ID,
        requestBody: {valueInputOption: "USER_ENTERED", data: requests.valueRequests}
    }));
    if (valueUpdateError) console.error(`Value update error: ${valueUpdateError}`);

    const [mergeUpdateError] = await handler(api.spreadsheets.batchUpdate({
        spreadsheetId: SPREADSHEET_ID,
        resource: {requests: requests.mergeRequests}
    }));
    if (mergeUpdateError) console.error(`Merge update error: ${mergeUpdateError}`);

    const [styleUpdateError] = await handler(api.spreadsheets.batchUpdate({
        spreadsheetId: SPREADSHEET_ID,
        resource: {requests: requests.styleRequests}
    }));
    if (styleUpdateError) console.error(`Style update error: ${styleUpdateError}`);

    console.log("Finished updating tabs!");
}