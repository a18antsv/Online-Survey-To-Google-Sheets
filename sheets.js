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

        const genderRangeStart = 1;
        const genderRangeEnd = genderRangeStart + tables["gender"].length - 1
        const genderRange = `A${genderRangeStart}:E${genderRangeEnd}`;

        const ageRangeStart = genderRangeEnd + 2;
        const ageRangeEnd = ageRangeStart + tables["age"].length - 1;
        const ageRange = `A${ageRangeStart}:E${ageRangeEnd}`;

        const regionRangeStart = ageRangeEnd + 2;
        const regionRangeEnd = regionRangeStart + (tables["region"].length === 0 ? -2 : tables["region"].length - 1);
        const regionRange = `A${regionRangeStart}:E${regionRangeEnd}`;

        const segmentRangeStart = regionRangeEnd + 2 + 1;
        const segmentRangeEnd = segmentRangeStart + tables["mainSegment"].length - 1;
        const segmentRange = `A${segmentRangeStart}:E${segmentRangeEnd}`;
        const segmentByOwnerRange = `F${segmentRangeStart}:I${segmentRangeEnd}`;
        const segmentByIntenderRange = `J${segmentRangeStart}:M${segmentRangeEnd}`;

        const brandRangeStart = 1;
        const brandRangeEnd = brandRangeStart + tables["brand"].length - 1;
        const brandRange = `G${brandRangeStart}:K${brandRangeEnd}`;

        const brandByOwnerRangeStart = segmentRangeEnd + 2;
        const brandByOwnerRangeEnd = brandByOwnerRangeStart + tables["ownerBrand"].length - 1;
        const brandByOwnerColumnCount = tables["ownerBrand"]?.[0]?.length ?? 0;
        const brandByOwnerRange = `A${brandByOwnerRangeStart}:${String.fromCharCode(64 + brandByOwnerColumnCount)}${brandByOwnerRangeEnd}`;

        const segmentBoosterRangeStart = brandByOwnerRangeEnd + 2;
        const segmentBoosterRangeEnd = segmentBoosterRangeStart + tables["segmentBooster"].length - 1;
        const segmentBoosterRange = `A${segmentBoosterRangeStart}:E${segmentBoosterRangeEnd}`;

        const evBoosterRangeStart = segmentBoosterRangeEnd + 2;
        const evBoosterRangeEnd = evBoosterRangeStart + tables["evBooster"].length - 1;
        const evBoosterRange = `A${evBoosterRangeStart}:E${evBoosterRangeEnd}`;

        const values = [
            {
                range: `${tab}!${genderRange}`,
                values: tables["gender"]
            },
            {
                range: `${tab}!${ageRange}`,
                values: tables["age"]
            },
            {
                range: `${tab}!${brandRange}`,
                values: tables["brand"]
            },
            {
                range: `${tab}!${regionRange}`,
                values: tables["region"]
            },
            {
                range: `${tab}!${segmentRange}`,
                values: tables["mainSegment"]
            },
            {
                range: `${tab}!${segmentByOwnerRange}`,
                values: tables["mainSegmentByOwner"]
            },
            {
                range: `${tab}!${segmentByIntenderRange}`,
                values: tables["mainSegmentByIntender"]
            },
            {
                range: `${tab}!F${segmentRangeStart - 1}`,
                values: [["Owner"]]
            },
            {
                range: `${tab}!J${segmentRangeStart - 1}`,
                values: [["Intender"]]
            },
            {
                range: `${tab}!${brandByOwnerRange}`,
                values: tables["ownerBrand"]
            },
            {
                range: `${tab}!${segmentBoosterRange}`,
                values: tables["segmentBooster"]
            },
            {
                range: `${tab}!${evBoosterRange}`,
                values: tables["evBooster"]
            },
        ];

        const mergeRequests = [
            {
                mergeCells: {
                    range: {
                        sheetId: tabId,
                        startRowIndex: segmentRangeStart - 2,
                        endRowIndex: segmentRangeStart - 1,
                        startColumnIndex: 5,
                        endColumnIndex: 9
                    },
                    mergeType: "MERGE_ROWS"
                }
            },
            {
                mergeCells: {
                    range: {
                        sheetId: tabId,
                        startRowIndex: segmentRangeStart - 2,
                        endRowIndex: segmentRangeStart - 1,
                        startColumnIndex: 9,
                        endColumnIndex: 13
                    },
                    mergeType: "MERGE_ROWS"
                }
            },
            {
                mergeCells: {
                    range: {
                        sheetId: tabId,
                        startRowIndex: segmentRangeStart - 2,
                        endRowIndex: segmentRangeStart,
                        startColumnIndex: 0,
                        endColumnIndex: 5
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
                        startRowIndex: segmentRangeStart - 2,
                        endRowIndex: segmentRangeStart,
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
            },
            // Header
            {
                repeatCell: {
                    range: {
                        sheetId: tabId,
                        startRowIndex: genderRangeStart - 1,
                        endRowIndex: genderRangeStart,
                        startColumnIndex: 0,
                        endColumnIndex: 5
                    },
                    ...styles.header
                }
            },
            {
                repeatCell: {
                    range: {
                        sheetId: tabId,
                        startRowIndex: ageRangeStart - 1,
                        endRowIndex: ageRangeStart,
                        startColumnIndex: 0,
                        endColumnIndex: 5
                    },
                    ...styles.header
                }
            },
            {
                repeatCell: {
                    range: {
                        sheetId: tabId,
                        startRowIndex: brandRangeStart - 1,
                        endRowIndex: brandRangeStart,
                        startColumnIndex: 6,
                        endColumnIndex: 11
                    },
                    ...styles.header
                }
            },
            {
                repeatCell: {
                    range: {
                        sheetId: tabId,
                        startRowIndex: segmentRangeStart - 2,
                        endRowIndex: segmentRangeStart,
                        startColumnIndex: 0,
                        endColumnIndex: 13
                    },
                    ...styles.header
                }
            },
            {
                repeatCell: {
                    range: {
                        sheetId: tabId,
                        startRowIndex: brandByOwnerRangeStart - 1,
                        endRowIndex: brandByOwnerRangeStart,
                        startColumnIndex: 0,
                        endColumnIndex: brandByOwnerColumnCount
                    },
                    ...styles.header
                }
            },
            {
                repeatCell: {
                    range: {
                        sheetId: tabId,
                        startRowIndex: segmentBoosterRangeStart - 1,
                        endRowIndex: segmentBoosterRangeStart,
                        startColumnIndex: 0,
                        endColumnIndex: 5
                    },
                    ...styles.header
                }
            },
            {
                repeatCell: {
                    range: {
                        sheetId: tabId,
                        startRowIndex: evBoosterRangeStart - 1,
                        endRowIndex: evBoosterRangeStart,
                        startColumnIndex: 0,
                        endColumnIndex: 5
                    },
                    ...styles.header
                }
            },
            // First column except header and footer
            {
                repeatCell: {
                    range: {
                        sheetId: tabId,
                        startRowIndex: genderRangeStart,
                        endRowIndex: genderRangeEnd - 1,
                        startColumnIndex: 0,
                        endColumnIndex: 1
                    },
                    ...styles.firstColumn
                }
            },
            {
                repeatCell: {
                    range: {
                        sheetId: tabId,
                        startRowIndex: ageRangeStart,
                        endRowIndex: ageRangeEnd - 1,
                        startColumnIndex: 0,
                        endColumnIndex: 1
                    },
                    ...styles.firstColumn
                }
            },
            {
                repeatCell: {
                    range: {
                        sheetId: tabId,
                        startRowIndex: brandRangeStart,
                        endRowIndex: brandRangeEnd - 1,
                        startColumnIndex: 6,
                        endColumnIndex: 7
                    },
                    ...styles.firstColumn
                }
            },
            {
                repeatCell: {
                    range: {
                        sheetId: tabId,
                        startRowIndex: segmentRangeStart,
                        endRowIndex: segmentRangeEnd - 1,
                        startColumnIndex: 0,
                        endColumnIndex: 1
                    },
                    ...styles.firstColumn
                }
            },
            {
                repeatCell: {
                    range: {
                        sheetId: tabId,
                        startRowIndex: brandByOwnerRangeStart,
                        endRowIndex: brandByOwnerRangeEnd - 1,
                        startColumnIndex: 0,
                        endColumnIndex: 1
                    },
                    ...styles.firstColumn
                }
            },
            {
                repeatCell: {
                    range: {
                        sheetId: tabId,
                        startRowIndex: segmentBoosterRangeStart,
                        endRowIndex: segmentBoosterRangeEnd - 1,
                        startColumnIndex: 0,
                        endColumnIndex: 1
                    },
                    ...styles.firstColumn
                }
            },
            {
                repeatCell: {
                    range: {
                        sheetId: tabId,
                        startRowIndex: evBoosterRangeStart,
                        endRowIndex: evBoosterRangeEnd - 1,
                        startColumnIndex: 0,
                        endColumnIndex: 1
                    },
                    ...styles.firstColumn
                }
            },
            // Footer
            {
                repeatCell: {
                    range: {
                        sheetId: tabId,
                        startRowIndex: genderRangeEnd - 1,
                        endRowIndex: genderRangeEnd,
                        startColumnIndex: 0,
                        endColumnIndex: 5
                    },
                    ...styles.footer
                }
            },
            {
                repeatCell: {
                    range: {
                        sheetId: tabId,
                        startRowIndex: ageRangeEnd - 1,
                        endRowIndex: ageRangeEnd,
                        startColumnIndex: 0,
                        endColumnIndex: 5
                    },
                    ...styles.footer
                }
            },
            {
                repeatCell: {
                    range: {
                        sheetId: tabId,
                        startRowIndex: brandRangeEnd - 1,
                        endRowIndex: brandRangeEnd,
                        startColumnIndex: 6,
                        endColumnIndex: 11
                    },
                    ...styles.footer
                }
            },
            {
                repeatCell: {
                    range: {
                        sheetId: tabId,
                        startRowIndex: segmentRangeEnd - 1,
                        endRowIndex: segmentRangeEnd,
                        startColumnIndex: 0,
                        endColumnIndex: 13
                    },
                    ...styles.footer
                }
            },
            {
                repeatCell: {
                    range: {
                        sheetId: tabId,
                        startRowIndex: brandByOwnerRangeEnd - 1,
                        endRowIndex: brandByOwnerRangeEnd,
                        startColumnIndex: 0,
                        endColumnIndex: brandByOwnerColumnCount
                    },
                    ...styles.footer
                }
            },
            {
                repeatCell: {
                    range: {
                        sheetId: tabId,
                        startRowIndex: segmentBoosterRangeEnd - 1,
                        endRowIndex: segmentBoosterRangeEnd,
                        startColumnIndex: 0,
                        endColumnIndex: 5
                    },
                    ...styles.footer
                }
            },
            {
                repeatCell: {
                    range: {
                        sheetId: tabId,
                        startRowIndex: evBoosterRangeEnd - 1,
                        endRowIndex: evBoosterRangeEnd,
                        startColumnIndex: 0,
                        endColumnIndex: 5
                    },
                    ...styles.footer
                }
            },
            // Borders
            {
                updateBorders: {
                    range: {
                        sheetId: tabId,
                        startRowIndex: genderRangeStart - 1,
                        endRowIndex: genderRangeEnd,
                        startColumnIndex: 0,
                        endColumnIndex: 5
                    },
                    ...styles.borders
                }
            },
            {
                updateBorders: {
                    range: {
                        sheetId: tabId,
                        startRowIndex: ageRangeStart - 1,
                        endRowIndex: ageRangeEnd,
                        startColumnIndex: 0,
                        endColumnIndex: 5
                    },
                    ...styles.borders
                }
            },
            {
                updateBorders: {
                    range: {
                        sheetId: tabId,
                        startRowIndex: brandRangeStart - 1,
                        endRowIndex: brandRangeEnd,
                        startColumnIndex: 6,
                        endColumnIndex: 11
                    },
                    ...styles.borders
                }
            },
            {
                updateBorders: {
                    range: {
                        sheetId: tabId,
                        startRowIndex: segmentRangeStart - 2,
                        endRowIndex: segmentRangeEnd,
                        startColumnIndex: 0,
                        endColumnIndex: 13
                    },
                    ...styles.borders
                }
            },
            {
                updateBorders: {
                    range: {
                        sheetId: tabId,
                        startRowIndex: brandByOwnerRangeStart - 1,
                        endRowIndex: brandByOwnerRangeEnd,
                        startColumnIndex: 0,
                        endColumnIndex: brandByOwnerColumnCount
                    },
                    ...styles.borders
                },
            },
            {
                updateBorders: {
                    range: {
                        sheetId: tabId,
                        startRowIndex: segmentBoosterRangeStart - 1,
                        endRowIndex: segmentBoosterRangeEnd,
                        startColumnIndex: 0,
                        endColumnIndex: 5
                    },
                    ...styles.borders
                }
            },
            {
                updateBorders: {
                    range: {
                        sheetId: tabId,
                        startRowIndex: evBoosterRangeStart - 1,
                        endRowIndex: evBoosterRangeEnd,
                        startColumnIndex: 0,
                        endColumnIndex: 5
                    },
                    ...styles.borders
                }
            },
        ];

        // Update region table only for countries that have regions
        if (tables["region"].length > 0) {
            const regionStyleRequests = [
                {
                    repeatCell: {
                        range: {
                            sheetId: tabId,
                            startRowIndex: regionRangeStart - 1,
                            endRowIndex: regionRangeStart,
                            startColumnIndex: 0,
                            endColumnIndex: 5
                        },
                        ...styles.header
                    }
                },
                {
                    repeatCell: {
                        range: {
                            sheetId: tabId,
                            startRowIndex: regionRangeStart,
                            endRowIndex: regionRangeEnd - 1,
                            startColumnIndex: 0,
                            endColumnIndex: 1
                        },
                        ...styles.firstColumn
                    }
                },
                {
                    repeatCell: {
                        range: {
                            sheetId: tabId,
                            startRowIndex: regionRangeEnd - 1,
                            endRowIndex: regionRangeEnd,
                            startColumnIndex: 0,
                            endColumnIndex: 5
                        },
                        ...styles.footer
                    }
                },
                {
                    updateBorders: {
                        range: {
                            sheetId: tabId,
                            startRowIndex: regionRangeStart - 1,
                            endRowIndex: regionRangeEnd,
                            startColumnIndex: 0,
                            endColumnIndex: 5
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