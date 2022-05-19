import "dotenv/config";
import {request} from "./request.js";
import {readCredentialsAndAuthorize} from "./auth.js";
import {updateTabs} from "./sheets.js";

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

const excludeCountries = [
    "2022_SA",
    "2022_EG",
    "2022_AE",
    "2022_IL",
    "2022_KZ",
    "2022_MA",
];

(async () => {
    const countries = (await request(bodies.getCountries))
        .filter(country => !excludeCountries.includes(country["PKEY"]));

    const countryDataByKey = {};

    for (const country of countries) {
        console.log(`Fetching data for ${country["country_name"]}...`);

        const key = country["PKEY"];
        const countryQuota = await getCountryWithCount(key);
        const countryQuotaSeries = await getCountryQuotaSeriesWithCount(key);
        const ownerBrand = await getOwnerBrandTableRows(key, countryQuotaSeries);

        countryDataByKey[key] = {
            tab: `${country["country_code"]}. ${key.split("_")[1]}`,
            tables: {
                region: getTableRows("Region", await getRegionData(key, countryQuota)),
                age: getTableRows("Age", await getAgeData(key, countryQuota)),
                gender: getTableRows("Gender", getGenderData(countryQuota)),
                brand: getTableRows("Brand", getBrandData(countryQuota)),
                mainSegment: getTableRows("Main Segment", getMainSegmentTotalData(countryQuota)),
                mainSegmentByOwner: getTableRows("Owner", getMainSegmentOwnerData(countryQuotaSeries), true),
                mainSegmentByIntender: getTableRows("Intender", getMainSegmentIntenderData(countryQuotaSeries), true),
                segmentBooster: getTableRows("Seg Booster", getSegmentBooster(countryQuotaSeries)),
                evBooster: getTableRows("EV Booster", getEvBooster(countryQuotaSeries)),
                ownerBrand,
            }
        };
    }

    readCredentialsAndAuthorize(auth => {
        updateTabs(auth, countryDataByKey);
    });
})();

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
    return countryQuotaSeries.filter(row => row["QSORDER"] === "6");
}

async function getOwnerBrandTableRows(key, countryQuotaSeries) {
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

function getTableRows(name, rows, removeFirst = false) {
    let quotaSum = 0;
    let achievedSum = 0;

    const tableRows = [
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

    if (removeFirst) {
        tableRows.forEach(row => row.shift());
    }

    return tableRows;
}

function toPercent(part, whole) {
    if (whole === 0) return "0.00%";
    return `${((part / whole) * 100).toFixed(2)}%`;
}
