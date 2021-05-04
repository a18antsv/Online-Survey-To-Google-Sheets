import { readFile, writeFile } from 'fs';
import { createInterface } from 'readline';
import { google } from 'googleapis';
import fetch from "node-fetch";

const SPREADSHEET_ID = "13oAA2yRDtfDAUCJIVSXnfTk7L40KMOJeeHEu9Tw8KQg";
const FETCH_URL = "http://w3.onlinesurvey.kr/project/quota_ajax.asp";
const SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]; // If modifying these scopes, delete token.json.
const TOKEN_PATH = "./token.json"; // The file token.json stores the user's access and refresh tokens, and is created automatically when the authorization flow completes for the first time.
const CREDENTIALS_PATH = "./credentials.json";
let countriesByKey = [];

const fetchBodies = {
  countries: "CRUD=SELECT&COMMAND=GET_TOTAL_QUOTA&ISCAPI=0",
  country: "CRUD=SELECT&COMMAND=GET_COUNTRY_QUOTA&PKEY=", // Add country key at the end
  segment: "CRUD=SELECT&COMMAND=GET_GROUPBY_SEGMENT_OWN&PKEY=", // Add country key at the end
  brandOwner: "CRUD=SELECT&COMMAND=GET_OWN_DATA&OWN=1&PKEY=" // Add country key at the end
}

const fetchOptions = {
  "headers": {
    "accept": "*/*",
    "accept-language": "en-GB,en;q=0.9,sv;q=0.8,en-US;q=0.7",
    "content-type": "application/x-www-form-urlencoded; charset=UTF-8",
    "x-requested-with": "XMLHttpRequest"
  },
  "referrer": "http://w3.onlinesurvey.kr/project/quota_index.asp?CID=83ex9w1qe2",
  "referrerPolicy": "strict-origin-when-cross-origin",
  "body": fetchBodies.countries,
  "method": "POST",
  "mode": "cors",
  "credentials": "include"
};

(async () => {
  countriesByKey = await getCountries();  

  // Load client secrets from a local file.
  readFile(CREDENTIALS_PATH, (error, content) => {
    if (error) {
      return console.error(`Error loading client secret file: ${error}`);
    }

    const credentials = JSON.parse(content);

    // Authorize a client with credentials, then call the Google Sheets API.
    authorize(credentials, updateTabs)
  });
})();


/**
 * Create an OAuth2 client with the given credentials, and then execute the given callback function.
 * @param {Object} credentials The authorization client credentials.
 * @param {function} callback The callback to call with the authorized client.
 */
function authorize(credentials, callback) {
  const { client_secret, client_id, redirect_uris } = credentials.installed;
  const oAuth2Client = new google.auth.OAuth2(client_id, client_secret, redirect_uris[0]);

  // Check if we have previously stored a token.
  readFile(TOKEN_PATH, (error, token) => {
    if (error) {
      return getNewToken(oAuth2Client, callback);
    }
    oAuth2Client.setCredentials(JSON.parse(token));
    callback(oAuth2Client);
  });
}

/**
 * Get and store new token after prompting for user authorization, and then execute the given callback with the authorized OAuth2 client.
 * @param {google.auth.OAuth2} oAuth2Client The OAuth2 client to get token for.
 * @param {getEventsCallback} callback The callback for the authorized client.
 */
function getNewToken(oAuth2Client, callback) {
  const authUrl = oAuth2Client.generateAuthUrl({ access_type: 'offline', scope: SCOPES });
  console.log('Authorize this app by visiting this url:', authUrl);
  const rl = createInterface({ input: process.stdin, output: process.stdout });
  rl.question('Enter the code from that page here: ', code => {
    rl.close();
    oAuth2Client.getToken(code, (err, token) => {
      if (err) return console.error('Error while trying to retrieve access token', err);
      oAuth2Client.setCredentials(token);
      // Store the token to disk for later program executions
      writeFile(TOKEN_PATH, JSON.stringify(token), (err) => {
        if (err) return console.error(err);
        console.log('Token stored to', TOKEN_PATH);
      });
      callback(oAuth2Client);
    });
  });
}

function handler(promise) {
  return promise
    .then(response => [undefined, response])
    .catch(error => [error, undefined]);
 }

async function getCountries() {
  const countriesByKey = {};
  
  console.log("Fetching all countries data...");
  fetchOptions.body = fetchBodies.countries;
  const [fetchCountriesError, countriesResponse] = await handler(fetch(FETCH_URL, fetchOptions));
  if(fetchCountriesError) {
    return console.error("Could not fetch countries");
  }

  const countries = await countriesResponse.json();
  //countries = countries.filter(obj => obj["PKEY"] === "2103022_SG"); // Filter the countries for testing
  for(const country of countries) {
    console.log(`Fetching data for country: ${country["country_name"]}...`);
    const key = country["PKEY"];
    const countryData = await getCountryData(key);
    if(!countryData) {
      console.error(`Could not fetch data for key ${key}.`);
      continue;
    }
    countriesByKey[key] = {
      number: country["country_code"],
      code: key.slice(key.lastIndexOf("_") + 1),
      name: country["country_name"],
      quota: country["TOTAL_CNT"],
      count: country["CNT"],
      data: countryData
    };
    console.log(`Finished fetching data for country: ${country["country_name"]}!`);
  }
  console.log("Finished fetching all countries data!");
  return countriesByKey;
}

async function getCountryData(key) {
  fetchOptions.body = fetchBodies.country + key;
  const [countryFetchError, countryResponse] = await handler(fetch(FETCH_URL, fetchOptions));
  if(countryFetchError) return;
  const countryData = await countryResponse.json();
  
  fetchOptions.body = fetchBodies.segment + key;
  const [segmentFetchError, segmentResponse] = await handler(fetch(FETCH_URL, fetchOptions));
  if(segmentFetchError) return;
  const segmentData = await segmentResponse.json();

  fetchOptions.body = fetchBodies.brandOwner + key;
  const [brandOwnerError, brandOwnerResponse] = await handler(fetch(FETCH_URL, fetchOptions));
  if(brandOwnerError) return;
  const brandOwnerData = await brandOwnerResponse.json();

  const ownerSegments = countryData.filter(object => object["grp"] === "Segment by OWNER");
  const intenderSegments = countryData.filter(object => object["grp"] === "Segment by INTENDER");
  const ownerSegmentsToUpdate = segmentData.filter(segment => segment.OWN == "1");
  const intenderSegmentsToUpdate = segmentData.filter(segment => segment.OWN == "2");

  for(const ownerSegmentToUpdate of ownerSegmentsToUpdate) {
    for(const ownerSegment of ownerSegments) {
      if(ownerSegmentToUpdate["Q_Seg_Quota"] === ownerSegment["ANS"]) {
        ownerSegment["com_cnt"] = ownerSegmentToUpdate["cnt"];
      }
    }
  }
  for(const intenderSegmentToUpdate of intenderSegmentsToUpdate) {
    for(const intenderSegment of intenderSegments) {
      if(intenderSegmentToUpdate["Q_Seg_Quota"] === intenderSegment["ANS"]) {
        intenderSegment["com_cnt"] = intenderSegmentToUpdate["cnt"];
      }
    }
  }

  const countryDataByGroup = countryData.reduce((accumulator, currentObject) => {
    if(!accumulator[currentObject["grp"]]) {
      accumulator[currentObject["grp"]] = [];
    }
    accumulator[currentObject["grp"]].push({
      ...currentObject,
      percent: ((currentObject["com_cnt"] / currentObject["cnt"]) * 100).toFixed(2),
      remaining: currentObject["cnt"] - currentObject["com_cnt"]
    });
    return accumulator;
  }, {}); 

  countryDataByGroup["BrandOwner"] = brandOwnerData;

  return countryDataByGroup;
}

function getTableRows(tableName, tableRowsObject, removeFirst = false) {
  const rows = [];
  const headerRow = [tableName, "Quota", "Achieved (N)", "Achieved (%)", "Remaining"];

  for(const tableRowObject of tableRowsObject) {
    const { ANS, label, cnt: quota, com_cnt: achievedNumber, percent: achievedPercent, remaining } = tableRowObject;
    rows.push([`${ANS}. ${label}`, parseInt(quota), parseInt(achievedNumber), parseFloat(achievedPercent), parseInt(remaining)]);
  }

  const footerRow = rows.reduce((accumulator, currentArray) => {
    for(let i = 1; i < currentArray.length; i++) {
      const value = currentArray[i];
      accumulator[i] = (accumulator[i] || 0) + value;
    }
    return accumulator;
  }, []);
  footerRow[0] = "Total";
  footerRow[3] = +((footerRow[2] / footerRow[1]) * 100).toFixed(2);

  rows.push(footerRow);
  rows.unshift(headerRow);

  if(removeFirst) {
    rows.forEach(row => row.shift());
  }

  return rows;
}

function getBrandByOwnerRows(tableName, countryData) {
  const rows = [];
  const headerRow = [tableName, "Quota"]

  const brandByOwnerByBrandLabel = countryData["BrandOwner"].reduce((accumulator, currentObject) => {
    const key = `${currentObject["BRAND_CD"]}. ${currentObject["LABEL"]}`;
    if(!accumulator[key]) {
      accumulator[key] = {
        data: [],
        quota: 0
      };
    }
    accumulator[key]["data"].push(currentObject);
    return accumulator;
  }, {});

  for(const brandLabel of Object.keys(brandByOwnerByBrandLabel)) {
    const brandObject = brandByOwnerByBrandLabel[brandLabel];
    const brandObjectData = brandObject.data;
    for(const brandObjectDataObject of brandObjectData) {
      for(const ownerSegment of countryData["Segment by OWNER"]) {
        if(brandObjectDataObject["SEG_CD"] === ownerSegment["ANS"]) {
          brandObjectDataObject["label"] = ownerSegment["label"];
        }
      }
      for(const brandByOwner of countryData["Brand by OWNER"]) {
        if(brandObjectDataObject["LABEL"] === brandByOwner["label"]) {
          brandByOwnerByBrandLabel[brandLabel]["quota"] = brandByOwner["cnt"];
        }
      }
    }

    const row = [brandLabel, parseInt(brandObject.quota)];
    let sum = 0;
    for(const brandObject of brandObjectData) {
      const cnt = parseInt(brandObject["CNT"])
      row.push(cnt);
      sum += cnt;
    }
    row.push(sum);
    rows.push(row);
  }

  const footerRow = rows.reduce((accumulator, currentArray) => {
    for(let i = 1; i < currentArray.length; i++) {
      const value = currentArray[i];
      accumulator[i] = (accumulator[i] || 0) + value;
    }
    return accumulator;
  }, []);
  footerRow[0] = "Total";

  brandByOwnerByBrandLabel[Object.keys(brandByOwnerByBrandLabel)[0]].data.forEach(object => {
    headerRow.push(object.label);
  });
  headerRow.push("Total");

  rows.unshift(headerRow);
  rows.push(footerRow);

  return rows;
}

async function addTabs(api, spreadsheet) {
  const addSheetRequests = [];

  for(const key of Object.keys(countriesByKey)) {
    const country = countriesByKey[key];
    const tabName = `${country["number"]}. ${country["code"]}`;
    const exists = spreadsheet.data.sheets.some(sheet => sheet.properties.title === tabName);
    if(!exists) {
      addSheetRequests.push({
        addSheet: {
          properties: {
            title: tabName
          }
        }
      });
    }
  }

  if(addSheetRequests.length > 0) {
    console.log("Adding tabs...");
    const [sheetUpdateError, sheetUpdateResponse] = await handler(api.spreadsheets.batchUpdate({
      spreadsheetId: SPREADSHEET_ID,
      resource: { requests: addSheetRequests }
    }));
    if(sheetUpdateError) {
      console.error(`Error: ${sheetUpdateError.errors[0].message}`);
    }
    console.log("Finished adding tabs!");
  }

  return addSheetRequests.length;
}

async function updateTabs(auth) {
  console.log("Updating tabs...");
  const api = google.sheets({ version: 'v4', auth });
  let [getSpreadsheetError, spreadsheet] = await handler(api.spreadsheets.get({ spreadsheetId: SPREADSHEET_ID }));
  if(getSpreadsheetError) {
    return console.error(`Error: ${getSpreadsheetError.errors[0].message}`);
  }

  const numberOfAddedSheets = await addTabs(api, spreadsheet);

  if(numberOfAddedSheets > 0) {
    [getSpreadsheetError, spreadsheet] = await handler(api.spreadsheets.get({ spreadsheetId: SPREADSHEET_ID }));
    if(getSpreadsheetError) {
      return console.error(`Error: ${getSpreadsheetError.errors[0].message}`);
    }
  }

  const allCountriesValues = [];
  const allCountriesMergeRequests = [];
  const allCountriesStyleRequests = [];

  const borderStyle = {
    style: "SOLID",
    color: {
      red: 0.0,
      green: 0.0,
      blue: 0.0,
      alpha: 1.0
    }
  };
  const borders = { top: borderStyle, right: borderStyle, bottom: borderStyle, left: borderStyle };
  const headerStyle = {
    cell: {
      userEnteredFormat: {
        backgroundColor: {
          red: 239 / 255,
          green: 239 / 255,
          blue: 239 / 255
        },
        horizontalAlignment: "CENTER",
        textFormat: {
          bold: true,
        }
      }
    },
    fields: "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)"
  };
  const firstColumnStyle = {
    cell: {
      userEnteredFormat: {
        backgroundColor: {
          red: 239 / 255,
          green: 239 / 255,
          blue: 239 / 255
        }
      }
    },
    fields: "userEnteredFormat(backgroundColor)"
  };
  const footerStyle = { ...firstColumnStyle };

  for(const key of Object.keys(countriesByKey)) {
    const country = countriesByKey[key];
    const tabName = `${country["number"]}. ${country["code"]}`;
    const tabId = spreadsheet.data.sheets.find(object => tabName === object.properties.title).properties.sheetId;

    const genderRows = getTableRows("Gender", country.data["Gender"]);
    const ageRows = getTableRows("Age", country.data["Age"]);
    const brandRows = getTableRows("Brand", country.data["Brand"]);
    const segmentRows = getTableRows("Segment", country.data["Segment"]);
    const segmentByOwnerRows = getTableRows("Owner", country.data["Segment by OWNER"], true);
    const segmentByIntenderRows = getTableRows("Intender", country.data["Segment by INTENDER"], true);
    const brandByOwnerRows = getBrandByOwnerRows("Brand Owner", country.data);
    
    let regionRows = getTableRows("Region", country.data["Region"] || []);
    if(regionRows.length <= 2) {
      regionRows = [];
    }

    const genderRangeStart = 1;
    const genderRangeEnd = genderRangeStart + genderRows.length - 1
    const genderRange = `A${genderRangeStart}:E${genderRangeEnd}`;

    const ageRangeStart = genderRangeEnd + 2;
    const ageRangeEnd = ageRangeStart + ageRows.length - 1;
    const ageRange = `A${ageRangeStart}:E${ageRangeEnd}`;

    const regionRangeStart = ageRangeEnd + 2;
    const regionRangeEnd = regionRangeStart + (regionRows.length === 0 ? -2 : regionRows.length - 1);
    const regionRange = `A${regionRangeStart}:E${regionRangeEnd}`;

    const segmentRangeStart = regionRangeEnd + 2 + 1;
    const segmentRangeEnd = segmentRangeStart + segmentRows.length - 1;
    const segmentRange = `A${segmentRangeStart}:E${segmentRangeEnd}`;
    const segmentByOwnerRange = `F${segmentRangeStart}:I${segmentRangeEnd}`;
    const segmentByIntenderRange = `J${segmentRangeStart}:M${segmentRangeEnd}`;

    const brandRangeStart = 1;
    const brandRangeEnd = brandRangeStart + brandRows.length - 1;
    const brandRange = `G${brandRangeStart}:K${brandRangeEnd}`;

    const brandByOwnerRangeStart = segmentRangeEnd + 2;
    const brandByOwnerRangeEnd = brandByOwnerRangeStart + brandByOwnerRows.length - 1;
    const brandByOwnerColumnCount = brandByOwnerRows[0].length;
    const brandByOwnerRange = `A${brandByOwnerRangeStart}:${String.fromCharCode(64 + brandByOwnerColumnCount)}${brandByOwnerRangeEnd}`;

    
    const values = [
      {
        range: `${tabName}!${genderRange}`,
        values: genderRows
      },
      {
        range: `${tabName}!${ageRange}`,
        values: ageRows
      },
      {
        range: `${tabName}!${brandRange}`,
        values: brandRows
      },
      {
        range: `${tabName}!${regionRange}`,
        values: regionRows
      },
      {
        range: `${tabName}!${segmentRange}`,
        values: segmentRows
      },
      {
        range: `${tabName}!${segmentByOwnerRange}`,
        values: segmentByOwnerRows
      },
      {
        range: `${tabName}!${segmentByIntenderRange}`,
        values: segmentByIntenderRows
      },
      {
        range: `${tabName}!F${segmentRangeStart - 1}`,
        values: [["Owner"]]
      },
      {
        range: `${tabName}!J${segmentRangeStart - 1}`,
        values: [["Intender"]]
      },
      {
        range: `${tabName}!${brandByOwnerRange}`,
        values: brandByOwnerRows
      }
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
          ...headerStyle
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
          ...headerStyle
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
          ...headerStyle
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
          ...headerStyle
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
          ...headerStyle
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
          ...firstColumnStyle
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
          ...firstColumnStyle
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
          ...firstColumnStyle
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
          ...firstColumnStyle
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
          ...firstColumnStyle
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
          ...footerStyle
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
          ...footerStyle
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
          ...footerStyle
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
          ...footerStyle
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
          ...footerStyle
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
          ...borders
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
          ...borders
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
          ...borders
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
          ...borders
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
          ...borders
        },
      }
    ];

    // Update region table only for countries that have regions
    if(regionRows.length > 0) {
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
            ...headerStyle
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
            ...firstColumnStyle
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
            ...footerStyle
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
            ...borders
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
    requestBody: { valueInputOption: "USER_ENTERED", data: allCountriesValues }
  }));
  if(valueUpdateError) {
    console.error(`Error: ${valueUpdateError.errors[0].message}`);
  }

  const [mergeUpdateError] = await handler(api.spreadsheets.batchUpdate({
    spreadsheetId: SPREADSHEET_ID,
    resource: { requests: allCountriesMergeRequests }
  }));
  if(mergeUpdateError) {
    console.error(`Error: ${mergeUpdateError.errors[0].message}`);
  }

  const [styleUpdateError] = await handler(api.spreadsheets.batchUpdate({
    spreadsheetId: SPREADSHEET_ID,
    resource: { requests: allCountriesStyleRequests }
  }));
  if(styleUpdateError) {
    console.error(`Error: ${styleUpdateError.errors[0].message}`);
  }

  console.log("Finished updating tabs!");
}
