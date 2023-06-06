import "dotenv/config"
import {apiCall, BasicData, CountryData, EvData, MainData} from "./api.js";
import {updateSheets} from "./sheets.js";
import {sheets_v4} from "googleapis";
import Schema$MergeCellsRequest = sheets_v4.Schema$MergeCellsRequest;
import Schema$UpdateBordersRequest = sheets_v4.Schema$UpdateBordersRequest;

const {
    FILTER_IN,
} = process.env;

const TEST = '1A08T3GA';
const pKeysToInclude: string[] = FILTER_IN?.split(",").map(countryCode => `2023_${countryCode}`) ?? [];
const CAPITAL_ASCII_START = 65;

const indexToChar = (index: number): string => String.fromCharCode(CAPITAL_ASCII_START - 1 + index);

interface Range {
    start: number;
    end: number;
}

interface Area {
    row: Range;
    column: Range;
    toString: () => string
}

export interface SheetData {
    sheetName: string,
    tables: TableData[]
}

interface TableData {
    area: Area,
    rows: (string | number)[][],
    mergeRequests?: Schema$MergeCellsRequest[]
    updateBorderRequests?: Schema$UpdateBordersRequest[]
}

const getArea = (
    rowStartIndex: number,
    columnStartIndex: number,
    rowCount: number,
    columnCount: number,
): Area => {
    const rowEndIndex = rowCount === 0 ? rowStartIndex : rowStartIndex - 1 + rowCount;
    const columnEndIndex = columnStartIndex - 1 + columnCount;
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

const toPercent = (part: number, whole: number): string => {
    if (whole === 0) return "0.00%";

    return `${((part / whole) * 100).toFixed(2)}%`;
}

const getBasicTableRows = (name: string, rows: BasicData[]): (string | number)[][] => {
    return [
        [name, 'Quota', 'Achie.(N)', 'Achie.(%)', 'Remaining'],
        ...rows.map((row) => [
            row.QSLABEL,
            row.Total_Cnt,
            row.Total_Complete,
            `${row.Total_Percent}%`,
            Math.max(0, row.Total_Cnt - row.Total_Complete),
        ]),
    ];
}

const getEvTableRows = (name: string, rows: EvData[]): (string | number)[][] => {
    return [
        [name, 'Total', 'Total', 'Total', 'Total', 'Owner', 'Owner', 'Intender', 'Intender'],
        [name, 'Quota', 'Achie.(N)', 'Achie.(%)', 'Remaining', 'Achie.(N)', 'Achie.(%)', 'Achie.(N)', 'Achie.(%)'],
        ...rows.map(row => {
            const totalComplete = row.Ownner_Complete + row.Intender_Complete;
            return [
                row.QSLABEL,
                row.Total_cnt,
                totalComplete,
                toPercent(totalComplete, row.Total_cnt),
                Math.max(0, row.Total_cnt - totalComplete),
                row.Ownner_Complete,
                `${row.Ownner_Percent}%`,
                row.Intender_Complete,
                `${row.Intender_Percent}%`,
            ]
        }),
    ];
}

const getMainTableRows = (name: string, rows: MainData[]): (string | number)[][] => {
    return [
        [name, 'Total', 'Total', 'Total', 'Total', 'Owner', 'Owner', 'Owner', 'Owner', 'Intender', 'Intender', 'Intender', 'Intender'],
        [name, 'Quota', 'Achie.(N)', 'Achie.(%)', 'Remaining', 'Quota', 'Achie.(N)', 'Achie.(%)', 'Remaining', 'Quota', 'Achie.(N)', 'Achie.(%)', 'Remaining'],
        ...rows.map(row => [
            row.QSLABEL,
            row.Total_cnt,
            row.Total_Complete,
            `${row.Total_Percent}%`,
            Math.max(0, row.Total_cnt - row.Total_Complete),
            row.Ownner_Cnt,
            row.Ownner_Complete,
            `${row.Ownner_Percent}%`,
            Math.max(0, row.Ownner_Cnt - row.Ownner_Complete),
            row.Intender_Cnt,
            row.Intender_Complete,
            `${row.Intender_Percent}%`,
            Math.max(0, row.Intender_Cnt - row.Intender_Complete),
        ]),
    ];
}

const getEvMergeRequests = (area: Area): Schema$MergeCellsRequest[] => {
    return [
        {
            mergeType: 'MERGE_COLUMNS',
            range: {
                startColumnIndex: area.column.start - 1,
                endColumnIndex: area.column.start,
                startRowIndex: area.row.start - 1,
                endRowIndex: area.row.start + 1,
            }
        },
        {
            mergeType: 'MERGE_ROWS',
            range: {
                startColumnIndex: area.column.start,
                endColumnIndex: area.column.start + 4,
                startRowIndex: area.row.start - 1,
                endRowIndex: area.row.start,
            }
        },
        {
            mergeType: 'MERGE_ROWS',
            range: {
                startColumnIndex: area.column.start + 4,
                endColumnIndex: area.column.start + 6,
                startRowIndex: area.row.start - 1,
                endRowIndex: area.row.start,
            }
        },
        {
            mergeType: 'MERGE_ROWS',
            range: {
                startColumnIndex: area.column.start + 6,
                endColumnIndex: area.column.start + 8,
                startRowIndex: area.row.start - 1,
                endRowIndex: area.row.start,
            }
        },
    ];
}

const getMainMergeRequests = (area: Area): Schema$MergeCellsRequest[] => {
    return [
        {
            mergeType: 'MERGE_COLUMNS',
            range: {
                startColumnIndex: area.column.start - 1,
                endColumnIndex: area.column.start,
                startRowIndex: area.row.start - 1,
                endRowIndex: area.row.start + 1,
            }
        },
        {
            mergeType: 'MERGE_ROWS',
            range: {
                startColumnIndex: area.column.start,
                endColumnIndex: area.column.start + 4,
                startRowIndex: area.row.start - 1,
                endRowIndex: area.row.start,
            }
        },
        {
            mergeType: 'MERGE_ROWS',
            range: {
                startColumnIndex: area.column.start + 4,
                endColumnIndex: area.column.start + 8,
                startRowIndex: area.row.start - 1,
                endRowIndex: area.row.start,
            }
        },
        {
            mergeType: 'MERGE_ROWS',
            range: {
                startColumnIndex: area.column.start + 8,
                endColumnIndex: area.column.start + 12,
                startRowIndex: area.row.start - 1,
                endRowIndex: area.row.start,
            }
        },
    ];
}

const main = async (): Promise<void> => {
    const countries = (await apiCall<CountryData[]>("/subCountry", new Map([
        ['isTest', TEST],
        ['isCapi', '0'],
        ['ccode', '999'],
    ])))
        .filter(country => pKeysToInclude.includes(country.PKEY));

    const sheets: SheetData[] = [];

    for (const country of countries) {
        const sheetName = `${country.country_code}. ${country.PKEY.split("_")[1]}`;
        console.log(`Fetching data for ${sheetName}...`);

        const genderData = (await apiCall<BasicData[]>("/basicTable", new Map([
            ['isTest', TEST],
            ['pkey', country.PKEY],
            ['qsString', 'Gender'],
        ])));
        const genderTableRows = getBasicTableRows('Gender', genderData);
        const genderArea = getArea(1, 1, genderTableRows.length, 5);

        const ageData = (await apiCall<BasicData[]>("/basicTable", new Map([
            ['isTest', TEST],
            ['pkey', country.PKEY],
            ['qsString', 'Age'],
        ])));
        const ageTableRows = getBasicTableRows('Age', ageData);
        const ageArea = getArea(genderArea.row.end + 2, 1, ageTableRows.length, 5);

        const regionData = (await apiCall<BasicData[]>("/basicTable", new Map([
            ['isTest', TEST],
            ['pkey', country.PKEY],
            ['qsString', 'Region'],
        ])));
        const regionTableRows = getBasicTableRows('Region', regionData);
        const regionArea = getArea(ageArea.row.end + 2, 1, regionTableRows.length, 5);

        const mainBrandData = await apiCall<MainData[]>("/subOwnTable", new Map([
            ['isTest', TEST],
            ['pkey', country.PKEY],
            ['qsString', 'Brand'],
        ]));
        const mainBrandTableRows = getMainTableRows('Main - Brand', mainBrandData);
        const mainBrandArea = getArea(1, genderArea.column.end + 2, mainBrandTableRows.length, 13);

        const mainSegmentData = (await apiCall<MainData[]>("/subOwnTable", new Map([
            ['isTest', TEST],
            ['pkey', country.PKEY],
            ['qsString', 'Seg_Quota'],
        ])));
        const mainSegmentTableRows = getMainTableRows('Main - Segment', mainSegmentData);
        const mainSegmentArea = getArea(mainBrandArea.row.end + 2, genderArea.column.end + 2, mainSegmentTableRows.length, 13);

        const mainEngineData = (await apiCall<MainData[]>("/subOwnTable", new Map([
            ['isTest', TEST],
            ['pkey', country.PKEY],
            ['qsString', 'Engine'],
        ])));
        const mainEngineTableRows = getMainTableRows('Main - Engine', mainEngineData);
        const mainEngineArea = getArea(mainSegmentArea.row.end + 2, genderArea.column.end + 2, mainEngineTableRows.length, 13);

        const teslaData = (await apiCall<EvData[]>("/evTable", new Map([
            ['isTest', TEST],
            ['pkey', country.PKEY],
            ['qsString', 'Tesla'],
        ])));
        const evTeslaTableRows = getEvTableRows('Tesla Count', teslaData);
        const evTeslaArea = getArea(mainEngineArea.row.end + 2, 1, evTeslaTableRows.length, 9);

        const evData = (await apiCall<EvData[]>("/evTable", new Map([
            ['isTest', TEST],
            ['pkey', country.PKEY],
            ['qsString', 'EV'],
        ])));
        const evTableRows = getEvTableRows('EV Count', evData);
        const evArea = getArea(evTeslaArea.row.end + 2, 1, evTableRows.length, 9);

        const sheet: SheetData = {
            sheetName: sheetName,
            tables: [
                {
                    area: genderArea,
                    rows: genderTableRows,
                    updateBorderRequests: [
                        {
                            range: {
                                startColumnIndex: genderArea.column.start - 1,
                                endColumnIndex: genderArea.column.end,
                                startRowIndex: genderArea.row.start - 1,
                                endRowIndex: genderArea.row.end,
                            },
                        }
                    ]
                },
                {
                    area: ageArea,
                    rows: ageTableRows,
                },
                {
                    area: regionArea,
                    rows: regionTableRows,
                },
                {
                    area: mainBrandArea,
                    rows: mainBrandTableRows,
                    mergeRequests: getMainMergeRequests(mainBrandArea),
                },
                {
                    area: mainSegmentArea,
                    rows: mainSegmentTableRows,
                    mergeRequests: getMainMergeRequests(mainSegmentArea),
                },
                {
                    area: mainEngineArea,
                    rows: mainEngineTableRows,
                    mergeRequests: getMainMergeRequests(mainEngineArea),
                },
                {
                    area: evTeslaArea,
                    rows: evTeslaTableRows,
                    mergeRequests: getEvMergeRequests(evTeslaArea),
                },
                {
                    area: evArea,
                    rows: evTableRows,
                    mergeRequests: getEvMergeRequests(evArea),
                }
            ],
        };

        sheets.push(sheet);
    }

    updateSheets(sheets)
}


main()
