import {promiseHandler as handler} from "./util.js";
import { styles } from "./style.js";
import {google} from "googleapis";

const {
    SPREADSHEET_ID,
} = process.env;

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

const getStyleRequest = (sheetId, range, options = {}) => {
    options = {
        top: {
            size: options?.top?.size ?? 1,
            hasStyle: options?.top?.hasStyle ?? true,
        },
        right: {
            size: options?.right?.size ?? 1,
            hasStyle: options?.right?.hasStyle ?? true,
        },
        bottom: {
            size: options?.bottom?.size ?? 1,
            hasStyle: options?.bottom?.hasStyle ?? true,
        },
        left: {
            size: options?.left?.size ?? 1,
            hasStyle: options?.left?.hasStyle ?? true,
        },
        border: {
            hasStyle: options?.border?.hasStyle ?? true,
        }
    };

    const topEnd = range.row.start + options.top.size - 1;
    const bottomStart = range.row.end - options.bottom.size;

    const top = {
        repeatCell: {
            range: {
                sheetId,
                startRowIndex: range.row.start - 1,
                endRowIndex: topEnd,
                startColumnIndex: range.column.start - 1,
                endColumnIndex: range.column.end
            },
            ...styles.header
        },
    };

    const bottom = {
        repeatCell: {
            range: {
                sheetId,
                startRowIndex: bottomStart,
                endRowIndex: range.row.end,
                startColumnIndex: range.column.start - 1,
                endColumnIndex: range.column.end
            },
            ...styles.footer
        },
    };

    const left = {
        repeatCell: {
            range: {
                sheetId,
                startRowIndex: topEnd,
                endRowIndex: bottomStart,
                startColumnIndex: range.column.start - 1,
                endColumnIndex: range.column.start
            },
            ...styles.firstColumn
        },
    };

    const border = {
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
    };

    const styleRequests = [];
    if (options.top.hasStyle) styleRequests.push(top);
    if (options.bottom.hasStyle) styleRequests.push(bottom);
    if (options.left.hasStyle) styleRequests.push(left);
    if (options.border.hasStyle) styleRequests.push(border);

    return styleRequests;
}

const getValueRequest = (tab, range, table) => {
    return {
        range: `${tab}!${range.toString()}`,
        values: table
    }
}

const getMergeRequest = (sheetId, mergeType, {startRowIndex, endRowIndex, startColumnIndex, endColumnIndex}) => {
    return {
        mergeCells: {
            mergeType,
            range: {
                sheetId,
                startRowIndex,
                endRowIndex,
                startColumnIndex,
                endColumnIndex
            },
        }
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

    const valueRequests = [];
    const styleRequests = [];
    const mergeRequests = [];

    Object.values(countryDataByKey).forEach(({tab, tables}) => {
        const sheetId = spreadsheet.data.sheets.find(object => tab === object.properties.title).properties.sheetId;

        Object.values(tables).forEach(({range, rows, styleOptions, merges}) => {
            if (rows.length === 0) return;
            valueRequests.push(getValueRequest(tab, range, rows));
            styleRequests.push(getStyleRequest(sheetId, range, styleOptions));
            mergeRequests.push(...(merges ?? []).map(merge => getMergeRequest(sheetId, merge.mergeType, merge.range)));
        });
    });

    if (valueRequests.length > 0) {
        const [valueUpdateError] = await handler(api.spreadsheets.values.batchUpdate({
            spreadsheetId: SPREADSHEET_ID,
            requestBody: {valueInputOption: "USER_ENTERED", data: valueRequests}
        }));
        if (valueUpdateError) console.error(`Value update error: ${valueUpdateError}`);
    }

    if (mergeRequests.length > 0) {
        const [mergeUpdateError] = await handler(api.spreadsheets.batchUpdate({
            spreadsheetId: SPREADSHEET_ID,
            resource: {requests: mergeRequests}
        }));
        if (mergeUpdateError) console.error(`Merge update error: ${mergeUpdateError}`);
    }

    if (styleRequests.length > 0) {
        const [styleUpdateError] = await handler(api.spreadsheets.batchUpdate({
            spreadsheetId: SPREADSHEET_ID,
            resource: {requests: styleRequests}
        }));
        if (styleUpdateError) console.error(`Style update error: ${styleUpdateError}`);
    }

    console.log("Finished updating tabs!");
}