import "dotenv/config";
import {request} from "./request.js";
import {readCredentialsAndAuthorize} from "./auth.js";
import {updateTabs} from "./sheets.js";

const {
    FILTER_IN,
} = process.env;

const bodies = {
    getCountries: "CRUD=SELECT&COMMAND=GET_TOTAL_QUOTA",
    getCountryQuota: "CRUD=SELECT&COMMAND=GET_COUNTRY_QUOTA&PKEY=",
    getCountryQuotaCount: "CRUD=SELECT&COMMAND=GET_COUNTRY_QUOTA_COM_CNT&PKEY=",
    getCountryQuotaSeries: "CRUD=SELECT&COMMAND=GET_COUNTRY_QUOTA_SERIES&PKEY=",
    getCountryQuotaSeriesCount: "CRUD=SELECT&COMMAND=GET_COUNTRY_QUOTA_COM_CNT_SERIES&PKEY=",
    getConvertAge: "CRUD=SELECT&COMMAND=GET_CONVERT_AGE&PKEY=",
    getConvertRegion: "CRUD=SELECT&COMMAND=GET_CONVERT_REGION&PKEY=",
    ownerBrand: "CRUD=SELECT&COMMAND=GET_OWN_DATA&OWN=1&PKEY=",
};

const CAPITAL_ASCII_START = 65;
const includeCountries = FILTER_IN.split(",").map(countryCode => `2022_${countryCode}`);

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

const main = async auth => {
    const countries = (await request(bodies.getCountries))
        .filter(country => includeCountries.includes(country["PKEY"]));

    const countryDataByKey = {};

    for (const country of countries) {
        console.log(`Fetching data for ${country["country_name"]}...`);

        const key = country["PKEY"];
        const tab = `${country["country_code"]}. ${key.split("_")[1]}`;
        const countryQuota = await getCountryWithCount(key);
        const countryQuotaSeries = await getCountryQuotaSeriesWithCount(key);
        const regionData = await getRegionData(key, countryQuota);
        const ageData = await getAgeData(key, countryQuota);
        
        const genderRows = getTableRows("Gender", getGenderData(countryQuota));
        const ageRows = getTableRows("Age", ageData);
        const regionRows = getTableRows("Region", regionData);
        const brandRows = getTableRows("Brand", getBrandData(countryQuota));
        const ownerBrandRows = await getOwnerBrandTableRows(key, countryQuotaSeries);
        const segmentBoosterRows = getTableRows("Seg Booster", getSegmentBooster(countryQuotaSeries));
        const evBoosterRows = getTableRows("EV Booster", getEvBooster(countryQuotaSeries));
        const mainSegmentRows = getTableRows("Main Segment", getMainSegmentTotalData(countryQuota), {addExtraHeader: true, extraHeaderTitle: "Total"});
        const mainSegmentByOwnerRows = getTableRows("Owner", getMainSegmentOwnerData(countryQuotaSeries), {removeFirst: true, addExtraHeader: true, extraHeaderTitle: "Owner"})
        const mainSegmentByIntenderRows = getTableRows("Intender", getMainSegmentIntenderData(countryQuotaSeries), {removeFirst: true, addExtraHeader: true, extraHeaderTitle: "Intender"});

        const genderRange = getRange(1, 1, genderRows);
        const ageRange = getRange(genderRange.row.end + 2, 1, ageRows);
        const regionRange = getRange(ageRange.row.end + 2, 1, regionRows)
        const brandRange = getRange(1, genderRange.column.end + 2, brandRows);
        const mainSegmentRange = getRange(regionRange.row.end + 2, 1, mainSegmentRows);
        const mainSegmentByOwnerRange = getRange(mainSegmentRange.row.start, mainSegmentRange.column.end + 1, mainSegmentByOwnerRows);
        const mainSegmentByIntenderRange = getRange(mainSegmentRange.row.start, mainSegmentByOwnerRange.column.end + 1, mainSegmentByIntenderRows);
        const ownerBrandRange = getRange(mainSegmentRange.row.end + 2, 1, ownerBrandRows);
        const segmentBoosterRange = getRange(ownerBrandRange.row.end + 2, 1, segmentBoosterRows);
        const evBoosterRange = getRange(segmentBoosterRange.row.end + 2, 1, evBoosterRows);

        countryDataByKey[key] = {
            tab,
            tables: {
                region: {
                    range: regionRange,
                    rows: regionRows,
                },
                age: {
                    range: ageRange,
                    rows: ageRows,
                },
                gender: {
                    range: genderRange,
                    rows: genderRows,
                },
                brand: {
                    range: brandRange,
                    rows: brandRows,
                },
                segmentBooster: {
                    range: segmentBoosterRange,
                    rows: segmentBoosterRows,
                },
                evBooster: {
                    range: evBoosterRange,
                    rows: evBoosterRows,
                },
                mainSegment: {
                    range: mainSegmentRange,
                    rows: mainSegmentRows,
                    styleOptions: {
                        top: {
                            size: 2,
                        }
                    },
                    merges: [{
                        mergeType: "MERGE_ROWS",
                        range: {
                            startRowIndex: mainSegmentRange.row.start - 1,
                            endRowIndex: mainSegmentRange.row.start,
                            startColumnIndex: mainSegmentRange.column.start - 1,
                            endColumnIndex: mainSegmentRange.column.end,                            
                        }
                    }]
                },
                mainSegmentByOwner: {
                    range: mainSegmentByOwnerRange,
                    rows: mainSegmentByOwnerRows,
                    styleOptions: {
                        top: {
                            size: 2,
                        },
                        left: {
                            hasStyle: false,
                        }
                    },
                    merges: [{
                        mergeType: "MERGE_ROWS",
                        range: {
                            startRowIndex: mainSegmentByOwnerRange.row.start - 1,
                            endRowIndex: mainSegmentByOwnerRange.row.start,
                            startColumnIndex: mainSegmentByOwnerRange.column.start - 1,
                            endColumnIndex: mainSegmentByOwnerRange.column.end,                            
                        }
                    }]
                },
                mainSegmentByIntender: {
                    range: mainSegmentByIntenderRange,
                    rows: mainSegmentByIntenderRows,
                    styleOptions: {
                        top: {
                            size: 2,
                        },
                        left: {
                            hasStyle: false,
                        }
                    },
                    merges: [{
                        mergeType: "MERGE_ROWS",
                        range: {
                            startRowIndex: mainSegmentByIntenderRange.row.start - 1,
                            endRowIndex: mainSegmentByIntenderRange.row.start,
                            startColumnIndex: mainSegmentByIntenderRange.column.start - 1,
                            endColumnIndex: mainSegmentByIntenderRange.column.end,                            
                        }
                    }]
                },
                ownerBrand: {
                    range: ownerBrandRange,
                    rows: ownerBrandRows,
                },
            }
        };
    }

    updateTabs(auth, countryDataByKey);
}

readCredentialsAndAuthorize(main);

async function getCountryWithCount(key) {
    const countryQuota = await request(bodies.getCountryQuota + key);
    const countryQuotaCount = await request(bodies.getCountryQuotaCount + key);

    return countryQuota.map(r1 => {
        const r2Matched = countryQuotaCount.find(r2 => (
            r1["PKEY"] === r2["PKEY"] &&
            r1["QID"].toLowerCase() === r2["QID"].toLowerCase() &&
            r1["QID_2"].toLowerCase() === r2["QID_2"].toLowerCase() &&
            r1["ANS"] === r2["ANS"]
        ));

        r1["COM_CNT"] = r2Matched?.["COM_CNT"] ?? "0";

        return r1;
    });
}

async function getCountryQuotaSeriesWithCount(key) {
    const quotaSeriesRows = await request(bodies.getCountryQuotaSeries + key);
    const quotaSeriesCountRows = await request(bodies.getCountryQuotaSeriesCount + key);

    return quotaSeriesRows.map(r1 => {
        const r2Matched = quotaSeriesCountRows.find(r2 => (
            r1["PKEY"] === r2["PKEY"] &&
            r1["QID"].toLowerCase() === r2["QID"].toLowerCase() &&
            r1["Q83_VAL"].toLowerCase() === r2["Q83_VAL"].toLowerCase() &&
            r1["QID_2"].toLowerCase() === r2["QID_2"].toLowerCase() &&
            r1["ANS"] === r2["ANS"]
        ));

        r1["COM_CNT"] = r2Matched?.["COM_CNT"] ?? "0";

        return r1;
    });
}

function getBrandData(countryQuota) {
    return countryQuota.filter(row => row["QID_2"] === "Q_Brand");
}

function getGenderData(countryQuota) {
    return countryQuota.filter(row => row["QID"] === "Q3");
}

async function getAgeData(key, countryQuota) {
    const convertAgeRows = await request(bodies.getConvertAge + key);

    return countryQuota
        .filter(row => row["QID"] === "Q6")
        .map(row => {
            row["COM_CNT"] = convertAgeRows.reduce((previous, current) => {
                return row["ANS"].includes(current["AGE"])
                    ? previous + parseInt(current["COM_CNT"])
                    : previous;
            }, 0);

            return row;
        });
}

async function getRegionData(key, countryQuota) {
    const convertRegionRows = await request(bodies.getConvertRegion + key);

    return countryQuota
        .filter(row => row["QID"] === "Q7")
        .map(row => {
            row["COM_CNT"] = convertRegionRows.reduce((previous, current) => {
                return row["ANS"].includes(current["REGION"])
                    ? previous + parseInt(current["COM_CNT"])
                    : previous;
            }, 0);

            return row;
        });
}

function getMainSegmentTotalData(countryQuota) {
    return countryQuota.filter(row => row["QID_2"] === "Q_Seg_Quota");
}

function getMainSegmentOwnerData(countryQuotaSeries) {
    return countryQuotaSeries.filter(row => row["QID_2"] === "Q_Seg_Quota" && row["Q83_VAL"] === "1");
}

function getMainSegmentIntenderData(countryQuotaSeries) {
    return countryQuotaSeries.filter(row => row["QID_2"] === "Q_Seg_Quota" && row["Q83_VAL"] === "2");
}

function getSegmentBooster(countryQuotaSeries) {
    return countryQuotaSeries
        .filter(row => row["QID_2"] === "Q_Seg_Booster")
        .map(row => {
            if (row["Q83_VAL"] === "6") {
                row["LABEL"] += " - Owner"
            } else {
                row["LABEL"] += " - Intender"
            }
            return row;
        });
}

function getEvBooster(countryQuotaSeries) {
    return countryQuotaSeries.filter(row => row["QSORDER"] === "6" || row["QSORDER"] === "7");
}

async function getOwnerBrandTableRows(key, countryQuotaSeries) {
    if (countryQuotaSeries.length === 0) return [];

    const ownerBrand = await request(bodies.ownerBrand + key);

    const segmentQuota = countryQuotaSeries.filter(row => row["QID_2"] === "Q_Seg_Quota" && row["Q83_VAL"] === "1");
    const labelsRow = segmentQuota.map(row => row["LABEL"]);
    const achievedRow = segmentQuota.map(row => row["COM_CNT"]);

    let tableRows = [["- Owner Brand", "Quota", ...labelsRow, "Total", "Total(%)"]];

    const ownerBrandQuotas = countryQuotaSeries.filter(row => row["QID_2"] === "Q_Brand" && row["Q83_VAL"] === "1");

    const rowsByBrandOwner = ownerBrand.reduce((previous, current) => {
        const key = `${current["BRAND_CD"]}. ${current["LABEL"]}`;

        if (!previous[key]) {
            const ownerBrandQuota = ownerBrandQuotas.find(row => row["ANS"] === current["BRAND_CD"]);
            previous[key] = {
                data: [],
                quota: parseInt(ownerBrandQuota["ORI_CNT"]),
                achieved: parseInt(ownerBrandQuota["COM_CNT"]),
            };
        }

        previous[key]["data"].push(current);

        return previous;
    }, {});

    let totalQuota = 0;
    let totalAchieved = 0;
    for (const [key, {data, quota, achieved}] of Object.entries(rowsByBrandOwner)) {
        const achievedCount = data.map(row => row["CNT"]);
        tableRows.push([key, quota, ...achievedCount, achieved, toPercent(achieved, quota)]);
        totalQuota += quota;
        totalAchieved += achieved;
    }

    tableRows.push(["Total", totalQuota, ...achievedRow, totalAchieved, toPercent(totalAchieved, totalQuota)]);

    return tableRows;
}

function getTableRows(name, rows, options = {}) {
    if (rows.length === 0) return [];

    options = {
        removeFirst: options?.removeFirst ?? false,
        addExtraHeader: options?.addExtraHeader ?? false,
        extraHeaderTitle: options?.extraHeaderTitle ?? "Title",
    };

    let quotaSum = 0;
    let achievedSum = 0;

    const tableRows = [
        ...options.addExtraHeader ? [Array.from({length: 5}).fill(options.extraHeaderTitle)] : [],
        [`- ${name}`, "Quota", "Achie.(N)", "Achie.(%)", "Remaining"],
        ...rows.map(row => {
            const {ANS, LABEL, COM_CNT, ORI_CNT} = row;
            const key = `${ANS}. ${LABEL}`;
            const quota = parseFloat(ORI_CNT);
            const achieved = parseFloat(COM_CNT);

            quotaSum += quota;
            achievedSum += achieved;

            return [key, quota, achieved, toPercent(achieved, quota), quota - achieved];
        }),
        ["Total", quotaSum, achievedSum, toPercent(achievedSum, quotaSum), quotaSum - achievedSum],
    ];

    if (options.removeFirst) {
        tableRows.forEach(row => row.shift());
    }

    return tableRows;
}

function toPercent(part, whole) {
    if (whole === 0) return "0.00%";
    return `${((part / whole) * 100).toFixed(2)}%`;
}
