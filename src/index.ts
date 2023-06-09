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

const indexToChar = (index: number): string => String.fromCharCode(CAPITAL_ASCII_START + index);

interface Range {
    start: number;
    end: number;
}

interface Area {
    row: Range;
    column: Range;
    a1Notation: string,
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
    {
        startColumnIndex,
        startRowIndex,
        columnCount,
        rowCount,
    }: {
        startColumnIndex: number,
        startRowIndex: number,
        columnCount: number,
        rowCount: number,
    }
): Area => {
    const column: Range = {
        start: startColumnIndex,
        end: startColumnIndex + columnCount,
    };

    const row: Range = {
        start: startRowIndex,
        end: startRowIndex + rowCount,
    };

    return {
        column,
        row,
        a1Notation: `${indexToChar(column.start)}${row.start + 1}:${indexToChar(column.end - 1)}${row.end}`
    };
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
            row.Total_Cnt - row.Total_Complete,
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
                row.Total_cnt - totalComplete,
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
            row.Total_cnt - row.Total_Complete,
            row.Ownner_Cnt,
            row.Ownner_Complete,
            `${row.Ownner_Percent}%`,
            row.Ownner_Cnt - row.Ownner_Complete,
            row.Intender_Cnt,
            row.Intender_Complete,
            `${row.Intender_Percent}%`,
            row.Intender_Cnt - row.Intender_Complete,
        ]),
    ];
}

const getEvMergeRequests = (area: Area): Schema$MergeCellsRequest[] => {
    return [
        {
            mergeType: 'MERGE_COLUMNS',
            range: {
                startColumnIndex: area.column.start,
                endColumnIndex: area.column.start + 1,
                startRowIndex: area.row.start,
                endRowIndex: area.row.start + 2,
            }
        },
        {
            mergeType: 'MERGE_ROWS',
            range: {
                startColumnIndex: area.column.start + 1,
                endColumnIndex: area.column.start + 5,
                startRowIndex: area.row.start,
                endRowIndex: area.row.start + 1,
            }
        },
        {
            mergeType: 'MERGE_ROWS',
            range: {
                startColumnIndex: area.column.start + 5,
                endColumnIndex: area.column.start + 7,
                startRowIndex: area.row.start,
                endRowIndex: area.row.start + 1,
            }
        },
        {
            mergeType: 'MERGE_ROWS',
            range: {
                startColumnIndex: area.column.start + 7,
                endColumnIndex: area.column.start + 9,
                startRowIndex: area.row.start,
                endRowIndex: area.row.start + 1,
            }
        },
    ];
}

const getMainMergeRequests = (area: Area): Schema$MergeCellsRequest[] => {
    return [
        {
            mergeType: 'MERGE_COLUMNS',
            range: {
                startColumnIndex: area.column.start,
                endColumnIndex: area.column.start + 1,
                startRowIndex: area.row.start,
                endRowIndex: area.row.start + 2,
            }
        },
        {
            mergeType: 'MERGE_ROWS',
            range: {
                startColumnIndex: area.column.start + 1,
                endColumnIndex: area.column.start + 5,
                startRowIndex: area.row.start,
                endRowIndex: area.row.start + 1,
            }
        },
        {
            mergeType: 'MERGE_ROWS',
            range: {
                startColumnIndex: area.column.start + 5,
                endColumnIndex: area.column.start + 9,
                startRowIndex: area.row.start,
                endRowIndex: area.row.start + 1,
            }
        },
        {
            mergeType: 'MERGE_ROWS',
            range: {
                startColumnIndex: area.column.start + 9,
                endColumnIndex: area.column.start + 13,
                startRowIndex: area.row.start,
                endRowIndex: area.row.start + 1,
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

        const genderArea = getArea({
            startColumnIndex: 0,
            startRowIndex: 0,
            columnCount: genderTableRows[0]?.length ?? 0,
            rowCount: genderTableRows.length,
        });

        const ageData = (await apiCall<BasicData[]>("/basicTable", new Map([
            ['isTest', TEST],
            ['pkey', country.PKEY],
            ['qsString', 'Age'],
        ])));
        const ageTableRows = getBasicTableRows('Age', ageData);
        const ageArea = getArea({
            startColumnIndex: 0,
            startRowIndex: genderArea.row.end + 1,
            columnCount: ageTableRows[0]?.length ?? 0,
            rowCount: ageTableRows.length,
        });

        const regionData = (await apiCall<BasicData[]>("/basicTable", new Map([
            ['isTest', TEST],
            ['pkey', country.PKEY],
            ['qsString', 'Region'],
        ])));
        const regionTableRows = getBasicTableRows('Region', regionData);
        const regionArea = getArea({
            startColumnIndex: 0,
            startRowIndex: ageArea.row.end + 1,
            columnCount: regionTableRows[0]?.length ?? 0,
            rowCount: regionTableRows.length,
        });

        const teslaData = (await apiCall<EvData[]>("/evTable", new Map([
            ['isTest', TEST],
            ['pkey', country.PKEY],
            ['qsString', 'Tesla'],
        ])));
        const evTeslaTableRows = getEvTableRows('Tesla Count', teslaData);
        const evTeslaArea = getArea({
            startColumnIndex: genderArea.column.end + 1,
            startRowIndex: 0,
            columnCount: evTeslaTableRows[0]?.length ?? 0,
            rowCount: evTeslaTableRows.length,
        });

        const evData = (await apiCall<EvData[]>("/evTable", new Map([
            ['isTest', TEST],
            ['pkey', country.PKEY],
            ['qsString', 'EV'],
        ])));
        const evTableRows = getEvTableRows('EV Count', evData);
        const evArea = getArea({
            startColumnIndex: evTeslaArea.column.start,
            startRowIndex: evTeslaArea.row.end + 1,
            columnCount: evTableRows[0]?.length ?? 0,
            rowCount: evTableRows.length,
        });

        const mainBrandData = await apiCall<MainData[]>("/subOwnTable", new Map([
            ['isTest', TEST],
            ['pkey', country.PKEY],
            ['qsString', 'Brand'],
        ]));
        const mainBrandTableRows = getMainTableRows('Main - Brand', mainBrandData);

        const mainBrandArea = getArea({
            startColumnIndex: 0,
            startRowIndex: Math.max(regionArea.row.end, evArea.row.end) + 1,
            columnCount: mainBrandTableRows[0]?.length ?? 0,
            rowCount: mainBrandTableRows.length,
        });

        const mainSegmentData = (await apiCall<MainData[]>("/subOwnTable", new Map([
            ['isTest', TEST],
            ['pkey', country.PKEY],
            ['qsString', 'Seg_Quota'],
        ])));
        const mainSegmentTableRows = getMainTableRows('Main - Segment', mainSegmentData);
        const mainSegmentArea = getArea({
            startColumnIndex: 0,
            startRowIndex: mainBrandArea.row.end + 1,
            columnCount: mainSegmentTableRows[0]?.length ?? 0,
            rowCount: mainSegmentTableRows.length,
        });

        const mainEngineData = (await apiCall<MainData[]>("/subOwnTable", new Map([
            ['isTest', TEST],
            ['pkey', country.PKEY],
            ['qsString', 'Engine'],
        ])));
        const mainEngineTableRows = getMainTableRows('Main - Engine', mainEngineData);
        const mainEngineArea = getArea({
            startColumnIndex: 0,
            startRowIndex: mainSegmentArea.row.end + 1,
            columnCount: mainEngineTableRows[0]?.length ?? 0,
            rowCount: mainEngineTableRows.length,
        });

        const sheet: SheetData = {
            sheetName: sheetName,
            tables: [
                {
                    area: genderArea,
                    rows: genderTableRows,
                    // updateBorderRequests: [
                    //     {
                    //         range: {
                    //             startColumnIndex: genderArea.column.start - 1,
                    //             endColumnIndex: genderArea.column.end,
                    //             startRowIndex: genderArea.row.start - 1,
                    //             endRowIndex: genderArea.row.end,
                    //         },
                    //     }
                    // ]
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
