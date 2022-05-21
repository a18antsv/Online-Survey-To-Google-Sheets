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

    const allCountriesValues = [];
    const allCountriesMergeRequests = [];
    const allCountriesStyleRequests = [];

    for (const {tab, tables} of Object.values(countryDataByKey)) {
        const tabId = spreadsheet.data.sheets.find(object => tab === object.properties.title).properties.sheetId;

        const genderRange = getRange(1, 1, tables["gender"]);
        const ageRange = getRange(genderRange.row.end + 2, 1, tables["age"]);
        const regionRange = getRange(ageRange.row.end + 2, 1, tables["region"])

        const segmentRange = getRange(regionRange.row.end + 3, 1, tables["mainSegment"]);
        const segmentByOwnerRange = getRange(regionRange.row.end + 3, segmentRange.column.end + 1, tables["mainSegmentByOwner"]);
        const segmentByIntenderRange = getRange(regionRange.row.end + 3, segmentByOwnerRange.column.end + 1, tables["mainSegmentByIntender"]);

        const brandRange = getRange(1, genderRange.column.end + 2, tables["brand"]);
        const brandByOwnerRange = getRange(segmentRange.row.end + 2, 1, tables["ownerBrand"]);
        const segmentBoosterRange = getRange(brandByOwnerRange.row.end + 2, 1, tables["segmentBooster"]);
        const evBoosterRange = getRange(segmentBoosterRange.row.end + 2, 1, tables["evBooster"]);

        const values = [
            {
                range: `${tab}!${genderRange.toString()}`,
                values: tables["gender"]
            },
            {
                range: `${tab}!${ageRange.toString()}`,
                values: tables["age"]
            },
            {
                range: `${tab}!${brandRange.toString()}`,
                values: tables["brand"]
            },
            {
                range: `${tab}!${regionRange.toString()}`,
                values: tables["region"]
            },
            {
                range: `${tab}!${segmentRange.toString()}`,
                values: tables["mainSegment"]
            },
            {
                range: `${tab}!${segmentByOwnerRange.toString()}`,
                values: tables["mainSegmentByOwner"]
            },
            {
                range: `${tab}!${segmentByIntenderRange.toString()}`,
                values: tables["mainSegmentByIntender"]
            },
            {
                range: `${tab}!F${segmentRange.row.start - 1}`,
                values: [["Owner"]]
            },
            {
                range: `${tab}!J${segmentRange.row.start - 1}`,
                values: [["Intender"]]
            },
            {
                range: `${tab}!${brandByOwnerRange.toString()}`,
                values: tables["ownerBrand"]
            },
            {
                range: `${tab}!${segmentBoosterRange.toString()}`,
                values: tables["segmentBooster"]
            },
            {
                range: `${tab}!${evBoosterRange.toString()}`,
                values: tables["evBooster"]
            },
        ];

        const mergeRequests = [
            {
                mergeCells: {
                    range: {
                        sheetId: tabId,
                        startRowIndex: segmentRange.row.start - 2,
                        endRowIndex: segmentRange.row.start - 1,
                        startColumnIndex: segmentByOwnerRange.column.start - 1,
                        endColumnIndex: segmentByOwnerRange.column.end
                    },
                    mergeType: "MERGE_ROWS"
                }
            },
            {
                mergeCells: {
                    range: {
                        sheetId: tabId,
                        startRowIndex: segmentRange.row.start - 2,
                        endRowIndex: segmentRange.row.start - 1,
                        startColumnIndex: segmentByIntenderRange.column.start - 1,
                        endColumnIndex: segmentByIntenderRange.column.end
                    },
                    mergeType: "MERGE_ROWS"
                }
            },
            {
                mergeCells: {
                    range: {
                        sheetId: tabId,
                        startRowIndex: segmentRange.row.start - 2,
                        endRowIndex: segmentRange.row.start,
                        startColumnIndex: segmentRange.column.start - 1,
                        endColumnIndex: segmentRange.column.end
                    },
                    mergeType: "MERGE_COLUMNS"
                }
            }
        ];

        const styleRequests = [
            {
                repeatCell: {
                    range: {
                        sheetId: tabId,
                        startRowIndex: segmentRange.row.start - 2,
                        endRowIndex: segmentRange.row.start,
                        startColumnIndex: segmentRange.column.start - 1,
                        endColumnIndex: segmentRange.column.end
                    },
                    cell: {
                        userEnteredFormat: {
                            verticalAlignment: "MIDDLE"
                        }
                    },
                    fields: "userEnteredFormat.verticalAlignment"
                }
            },
            // Header
            {
                repeatCell: {
                    range: {
                        sheetId: tabId,
                        startRowIndex: genderRange.row.start - 1,
                        endRowIndex: genderRange.row.start,
                        startColumnIndex: genderRange.column.start - 1,
                        endColumnIndex: genderRange.column.end
                    },
                    ...styles.header
                }
            },
            {
                repeatCell: {
                    range: {
                        sheetId: tabId,
                        startRowIndex: ageRange.row.start - 1,
                        endRowIndex: ageRange.row.start,
                        startColumnIndex: ageRange.column.start - 1,
                        endColumnIndex: ageRange.column.end
                    },
                    ...styles.header
                }
            },
            {
                repeatCell: {
                    range: {
                        sheetId: tabId,
                        startRowIndex: brandRange.row.start - 1,
                        endRowIndex: brandRange.row.start,
                        startColumnIndex: brandRange.column.start - 1,
                        endColumnIndex: brandRange.column.end
                    },
                    ...styles.header
                }
            },
            {
                repeatCell: {
                    range: {
                        sheetId: tabId,
                        startRowIndex: segmentRange.row.start - 2,
                        endRowIndex: segmentRange.row.start,
                        startColumnIndex: segmentRange.column.start - 1,
                        endColumnIndex: segmentByIntenderRange.column.end
                    },
                    ...styles.header
                }
            },
            {
                repeatCell: {
                    range: {
                        sheetId: tabId,
                        startRowIndex: brandByOwnerRange.row.start - 1,
                        endRowIndex: brandByOwnerRange.row.start,
                        startColumnIndex: brandByOwnerRange.column.start - 1,
                        endColumnIndex: brandByOwnerRange.column.end
                    },
                    ...styles.header
                }
            },
            {
                repeatCell: {
                    range: {
                        sheetId: tabId,
                        startRowIndex: segmentBoosterRange.row.start - 1,
                        endRowIndex: segmentBoosterRange.row.start,
                        startColumnIndex: segmentBoosterRange.column.start - 1,
                        endColumnIndex: segmentBoosterRange.column.end
                    },
                    ...styles.header
                }
            },
            {
                repeatCell: {
                    range: {
                        sheetId: tabId,
                        startRowIndex: evBoosterRange.row.start - 1,
                        endRowIndex: evBoosterRange.row.start,
                        startColumnIndex: evBoosterRange.column.start - 1,
                        endColumnIndex: evBoosterRange.column.end
                    },
                    ...styles.header
                }
            },
            // First column except header and footer
            {
                repeatCell: {
                    range: {
                        sheetId: tabId,
                        startRowIndex: genderRange.row.start,
                        endRowIndex: genderRange.row.end - 1,
                        startColumnIndex: genderRange.column.start - 1,
                        endColumnIndex: genderRange.column.start
                    },
                    ...styles.firstColumn
                }
            },
            {
                repeatCell: {
                    range: {
                        sheetId: tabId,
                        startRowIndex: ageRange.row.start,
                        endRowIndex: ageRange.row.end - 1,
                        startColumnIndex: ageRange.column.start - 1,
                        endColumnIndex: ageRange.column.start
                    },
                    ...styles.firstColumn
                }
            },
            {
                repeatCell: {
                    range: {
                        sheetId: tabId,
                        startRowIndex: brandRange.row.start,
                        endRowIndex: brandRange.row.end - 1,
                        startColumnIndex: brandRange.column.start - 1,
                        endColumnIndex: brandRange.column.start
                    },
                    ...styles.firstColumn
                }
            },
            {
                repeatCell: {
                    range: {
                        sheetId: tabId,
                        startRowIndex: segmentRange.row.start,
                        endRowIndex: segmentRange.row.end - 1,
                        startColumnIndex: segmentRange.column.start - 1,
                        endColumnIndex: segmentRange.column.start
                    },
                    ...styles.firstColumn
                }
            },
            {
                repeatCell: {
                    range: {
                        sheetId: tabId,
                        startRowIndex: brandByOwnerRange.row.start,
                        endRowIndex: brandByOwnerRange.row.end - 1,
                        startColumnIndex: brandByOwnerRange.column.start - 1,
                        endColumnIndex: brandByOwnerRange.column.start
                    },
                    ...styles.firstColumn
                }
            },
            {
                repeatCell: {
                    range: {
                        sheetId: tabId,
                        startRowIndex: segmentBoosterRange.row.start,
                        endRowIndex: segmentBoosterRange.row.end - 1,
                        startColumnIndex: segmentBoosterRange.column.start - 1,
                        endColumnIndex: segmentBoosterRange.column.start
                    },
                    ...styles.firstColumn
                }
            },
            {
                repeatCell: {
                    range: {
                        sheetId: tabId,
                        startRowIndex: evBoosterRange.row.start,
                        endRowIndex: evBoosterRange.row.end - 1,
                        startColumnIndex: evBoosterRange.column.start - 1,
                        endColumnIndex: evBoosterRange.column.start
                    },
                    ...styles.firstColumn
                }
            },
            // Footer
            {
                repeatCell: {
                    range: {
                        sheetId: tabId,
                        startRowIndex: genderRange.row.end - 1,
                        endRowIndex: genderRange.row.end,
                        startColumnIndex: genderRange.column.start - 1,
                        endColumnIndex: genderRange.column.end
                    },
                    ...styles.footer
                }
            },
            {
                repeatCell: {
                    range: {
                        sheetId: tabId,
                        startRowIndex: ageRange.row.end - 1,
                        endRowIndex: ageRange.row.end,
                        startColumnIndex: ageRange.column.start - 1,
                        endColumnIndex: ageRange.column.end
                    },
                    ...styles.footer
                }
            },
            {
                repeatCell: {
                    range: {
                        sheetId: tabId,
                        startRowIndex: brandRange.row.end - 1,
                        endRowIndex: brandRange.row.end,
                        startColumnIndex: brandRange.column.start - 1,
                        endColumnIndex: brandRange.column.end
                    },
                    ...styles.footer
                }
            },
            {
                repeatCell: {
                    range: {
                        sheetId: tabId,
                        startRowIndex: segmentRange.row.end - 1,
                        endRowIndex: segmentRange.row.end,
                        startColumnIndex: segmentRange.column.start - 1,
                        endColumnIndex: segmentByIntenderRange.column.end
                    },
                    ...styles.footer
                }
            },
            {
                repeatCell: {
                    range: {
                        sheetId: tabId,
                        startRowIndex: brandByOwnerRange.row.end - 1,
                        endRowIndex: brandByOwnerRange.row.end,
                        startColumnIndex: brandByOwnerRange.column.start - 1,
                        endColumnIndex: brandByOwnerRange.column.end
                    },
                    ...styles.footer
                }
            },
            {
                repeatCell: {
                    range: {
                        sheetId: tabId,
                        startRowIndex: segmentBoosterRange.row.end - 1,
                        endRowIndex: segmentBoosterRange.row.end,
                        startColumnIndex: segmentBoosterRange.column.start - 1,
                        endColumnIndex: segmentBoosterRange.column.end
                    },
                    ...styles.footer
                }
            },
            {
                repeatCell: {
                    range: {
                        sheetId: tabId,
                        startRowIndex: evBoosterRange.row.end - 1,
                        endRowIndex: evBoosterRange.row.end,
                        startColumnIndex: evBoosterRange.column.start - 1,
                        endColumnIndex: evBoosterRange.column.end
                    },
                    ...styles.footer
                }
            },
            // Borders
            {
                updateBorders: {
                    range: {
                        sheetId: tabId,
                        startRowIndex: genderRange.row.start - 1,
                        endRowIndex: genderRange.row.end,
                        startColumnIndex: genderRange.column.start - 1,
                        endColumnIndex: genderRange.column.end
                    },
                    ...styles.borders
                }
            },
            {
                updateBorders: {
                    range: {
                        sheetId: tabId,
                        startRowIndex: ageRange.row.start - 1,
                        endRowIndex: ageRange.row.end,
                        startColumnIndex: ageRange.column.start - 1,
                        endColumnIndex: ageRange.column.end
                    },
                    ...styles.borders
                }
            },
            {
                updateBorders: {
                    range: {
                        sheetId: tabId,
                        startRowIndex: brandRange.row.start - 1,
                        endRowIndex: brandRange.row.end,
                        startColumnIndex: brandRange.column.start - 1,
                        endColumnIndex: brandRange.column.end
                    },
                    ...styles.borders
                }
            },
            {
                updateBorders: {
                    range: {
                        sheetId: tabId,
                        startRowIndex: segmentRange.row.start - 2,
                        endRowIndex: segmentRange.row.end,
                        startColumnIndex: segmentRange.column.start - 1,
                        endColumnIndex: segmentByIntenderRange.column.end
                    },
                    ...styles.borders
                }
            },
            {
                updateBorders: {
                    range: {
                        sheetId: tabId,
                        startRowIndex: brandByOwnerRange.row.start - 1,
                        endRowIndex: brandByOwnerRange.row.end,
                        startColumnIndex: brandByOwnerRange.column.start - 1,
                        endColumnIndex: brandByOwnerRange.column.end
                    },
                    ...styles.borders
                },
            },
            {
                updateBorders: {
                    range: {
                        sheetId: tabId,
                        startRowIndex: segmentBoosterRange.row.start - 1,
                        endRowIndex: segmentBoosterRange.row.end,
                        startColumnIndex: segmentBoosterRange.column.start - 1,
                        endColumnIndex: segmentBoosterRange.column.end
                    },
                    ...styles.borders
                }
            },
            {
                updateBorders: {
                    range: {
                        sheetId: tabId,
                        startRowIndex: evBoosterRange.row.start - 1,
                        endRowIndex: evBoosterRange.row.end,
                        startColumnIndex: evBoosterRange.column.start - 1,
                        endColumnIndex: evBoosterRange.column.end
                    },
                    ...styles.borders
                }
            },
        ];

        if (tables["region"].length > 0) {
            const regionStyleRequests = [
                {
                    repeatCell: {
                        range: {
                            sheetId: tabId,
                            startRowIndex: regionRange.row.start - 1,
                            endRowIndex: regionRange.row.start,
                            startColumnIndex: regionRange.column.start - 1,
                            endColumnIndex: regionRange.column.end
                        },
                        ...styles.header
                    }
                },
                {
                    repeatCell: {
                        range: {
                            sheetId: tabId,
                            startRowIndex: regionRange.row.start,
                            endRowIndex: regionRange.row.end - 1,
                            startColumnIndex: regionRange.column.start - 1,
                            endColumnIndex: regionRange.column.start
                        },
                        ...styles.firstColumn
                    }
                },
                {
                    repeatCell: {
                        range: {
                            sheetId: tabId,
                            startRowIndex: regionRange.row.end - 1,
                            endRowIndex: regionRange.row.end,
                            startColumnIndex: regionRange.column.start - 1,
                            endColumnIndex: regionRange.column.end
                        },
                        ...styles.footer
                    }
                },
                {
                    updateBorders: {
                        range: {
                            sheetId: tabId,
                            startRowIndex: regionRange.row.start - 1,
                            endRowIndex: regionRange.row.end,
                            startColumnIndex: regionRange.column.start - 1,
                            endColumnIndex: regionRange.column.end
                        },
                        ...styles.borders
                    }
                }
            ];
            allCountriesStyleRequests.push(...regionStyleRequests);
        }
        allCountriesValues.push(...values);
        allCountriesMergeRequests.push(...mergeRequests);
        allCountriesStyleRequests.push(...styleRequests);
    }

    const [valueUpdateError] = await handler(api.spreadsheets.values.batchUpdate({
        spreadsheetId: SPREADSHEET_ID,
        requestBody: {valueInputOption: "USER_ENTERED", data: allCountriesValues}
    }));
    if (valueUpdateError) console.error(`Value update error: ${valueUpdateError}`);

    const [mergeUpdateError] = await handler(api.spreadsheets.batchUpdate({
        spreadsheetId: SPREADSHEET_ID,
        resource: {requests: allCountriesMergeRequests}
    }));
    if (mergeUpdateError) console.error(`Merge update error: ${mergeUpdateError}`);

    const [styleUpdateError] = await handler(api.spreadsheets.batchUpdate({
        spreadsheetId: SPREADSHEET_ID,
        resource: {requests: allCountriesStyleRequests}
    }));
    if (styleUpdateError) console.error(`Style update error: ${styleUpdateError}`);

    console.log("Finished updating tabs!");
}