function removeAdTemplates(spreadsheetId, rowData, sheetId) {
  // Remove ad templates

  let removeRequests = [];

  for(let i = 1; i < rowData.length; i++) {
    const row = rowData[i];

    if (row.values && row.values.length > 5 && row.values.length !== 10) {
      removeRequests.push(createDeleteDimensionRequest(sheetId, i - removeRequests.length, 1, "ROWS"));
    } 
  }
  if(removeRequests.length > 0) {
    batchUpdate(spreadsheetId, removeRequests);
  }
}
  
function removeEmptyRows(spreadsheetId, rowData, sheetId) {
  let removeRequests = [];

  for(let i = 1; i < rowData.length; i++) {
    const row = rowData[i];

    if (rowIsEmpty(row)) {
      removeRequests.push(createDeleteDimensionRequest(sheetId, i - removeRequests.length, 1, "ROWS"));
    } 
  }
  if(removeRequests.length > 0) {
    batchUpdate(spreadsheetId, removeRequests);
  }
}

function removeExtraCols(spreadsheetId, rowData, sheetId) {
  let removeRequests = [];

  // deletes extra columns. deletes column 5, 5 times because batchUpdate processes requests in order.
  removeRequests.push(createDeleteDimensionRequest(sheetId, 9, rowData.length, "COLUMNS"));
  removeRequests.push(createDeleteDimensionRequest(sheetId, 8, rowData.length, "COLUMNS"));
  removeRequests.push(createDeleteDimensionRequest(sheetId, 7, rowData.length, "COLUMNS"));
  removeRequests.push(createDeleteDimensionRequest(sheetId, 6, rowData.length, "COLUMNS"));
  removeRequests.push(createDeleteDimensionRequest(sheetId, 5, rowData.length, "COLUMNS"));
  
  if(removeRequests.length > 0) {
    batchUpdate(spreadsheetId, removeRequests);
  }
}

function getAccountsFromManagerSheet(sheet) {
  if (!sheet || !sheet.data) {
    console.error('getAccountsFromManagerSheet: Invalid sheet:', sheet);
    alert('Account data is not available. Please try again.');
    return [];
  }
  let accounts = [];
  //skip header
  for(let i = 1; i < sheet.data[0].rowData.length; i++) {
    const row = sheet.data[0].rowData[i];
    if(rowIsEmpty(row)) continue;
    const accountTitle = row.values[1].userEnteredValue.stringValue;

    if(!accounts.includes(accountTitle)) {
      accounts.push(accountTitle);
    }
  }

  return accounts;
}

function insertAds(spreadsheetId, sheetId, rowData) {
  /**
   * 
   * adGroupIndex Obj {
   * 
   *  campaign,
   *  adGroup,
   *  A1index,
   * 
   * }
   * 
   */
    let adGroupIndexes = [];
  
    for (let i = 1; i < rowData.length; i++) {
      const row = rowData[i];
                
      //skips empty rows
      if (!rowIsEmpty(row)) {
        const campaign = row.values[0].userEnteredValue.stringValue;
        const adGroup = row.values[1].userEnteredValue.stringValue;
        const index = i;
  
        const adGroupIndex = {campaign, adGroup, index}
        if (!indexIncluded(adGroupIndexes, adGroupIndex)){
          adGroupIndexes.push(adGroupIndex);
        } 
      }
    }
  
    let requests = [];
  
    for (let i = 0; i < adGroupIndexes.length; i++) {         
      const index = adGroupIndexes[i].index;
      requests.push(createInsertDimensionRequest(sheetId, index + i, 1));
      requests.push(createUpdateCellsRequest(sheetId, index + i, ["test", "test2", "test3", "test4", "test5" ] ))
    }
    
    batchUpdate(spreadsheetId, requests);
}

function rowIsEmpty(row) {
  if(Object.keys(row).length === 0) {
    return true;
  }

  for(let i = 1; i < row.values.length; i++) {
    if(row.values[i].hasOwnProperty('userEnteredValue')) {
      return false;
    }
  }

  return true;
}

function getAdCopyRowData(sheet, account, language, type) {
  let adCopyRows = [];
  const rows = sheet.data[0].rowData;
  for(let i = 0; i < rows.length; i++) {
    if(!rowIsEmpty(rows[i]) && rows[i].values[1].hasOwnProperty('userEnteredValue')) {
      const adAccount = rows[i].values[1].userEnteredValue.stringValue;
      const adLanguage = rows[i].values[2].userEnteredValue.stringValue;
      const adType = rows[i].values[3].userEnteredValue.stringValue;
      if(adAccount === account && adLanguage === language && adType === type)  {
        adCopyRows.push(rows[i])
      }
    } 
  }

  return adCopyRows;
}

async function processRequest(buildoutSpreadsheet, accountDataSpreadsheet, accounts) {
  if (!buildoutSpreadsheet || !buildoutSpreadsheet.sheets) {
    alert('Failed to load the brand buildout spreadsheet. Please check your access and try again.');
    return;
  }
  if (!accountDataSpreadsheet || !accountDataSpreadsheet.sheets) {
    alert('Failed to load the account data spreadsheet. Please check your access and try again.');
    return;
  }
  const urlDataSheet = getUrlDataSheet(accountDataSpreadsheet);
  let spreadsheets = [];

  const managerSheet = getManagerSheet(accountDataSpreadsheet, MANAGER);

  for(let i = 0; i < accounts.length; i++) {
    const accountBuildoutSpreadsheet = await createAccountBuildoutSpreadsheet(buildoutSpreadsheet, managerSheet, urlDataSheet, accounts[i]);
    spreadsheets.push(accountBuildoutSpreadsheet);
  }

  for(let i = 0; i < spreadsheets.length; i++) {
    const newSpreadsheet = await createNewDocument(spreadsheets[i]);
    if (newSpreadsheet) {
      const url = newSpreadsheet.spreadsheetUrl;
      window.open(url, '_blank');
    } else {
      console.warn("createNewDocument returned null (or undefined) for spreadsheet index " + i + ". Skipping open new window.");
    }
  }

  location.reload();
}

function getUrlDataSheet(accountDataSpreadsheet) {
  for(let i = 0; i < accountDataSpreadsheet.sheets.length; i++) {
    const sheet = accountDataSpreadsheet.sheets[i];

    if(sheet.properties.title === "URL Data") {
      return sheet;
    }
  }
  
  return null;
}


function getAccountAdCopySheet(accountDataSpreadsheet, account) {
  for(let i = 0; i < accountDataSpreadsheet.sheets.length; i++) {
    const sheet = accountDataSpreadsheet.sheets[i];
    if(sheet.properties.title === account || sheet.properties.title === account.accountTitle) {
      return sheet;
    }
  }
  
  return null;
}

function getAccountsFromAccountSheet(sheet) {
  let accounts = [];
  //skip header
  for(let i = 1; i < sheet.data[0].rowData.length; i++) {
    const row = sheet.data[0].rowData[i];
    if(rowIsEmpty(row)) continue;

    const accountTitle = row.values[0].userEnteredValue.stringValue;
    
    if(accountTitle !== "URL DATA"){
      accounts.push(accountTitle);
    }
  }

  return accounts;
}

function getAccountCampaignsFromSheet(sheet, account) {
  let campaigns = [];

  //skip header
  for(let i = 1; i < sheet.data[0].rowData.length; i++) {
    const row = sheet.data[0].rowData[i];
    if(rowIsEmpty(row)){
      continue;
    }
    const campaign = row.values[3].userEnteredValue.stringValue;
    const rowAccount = row.values[1].userEnteredValue.stringValue;
    if(rowAccount === account && !campaigns.includes(campaign)) {
      campaigns.push(campaign);
    }
  }

  return campaigns;
}

function getAccountLanguagesFromSheet(sheet, account) { 
  let languages = [];
  //skip header
  for(let i = 1; i < sheet.data[0].rowData.length; i++) {
    const row = sheet.data[0].rowData[i];
    if(rowIsEmpty(row)){
      continue;
    }
    const language = row.values[2].userEnteredValue.stringValue;
    const rowAccount = row.values[1].userEnteredValue.stringValue;

    if(rowAccount === account && !languages.includes(language)) {
      languages.push(language);
    }
  }

  return languages;
}

function getPostfixFromSheet(sheet, account, language) {
  //skip header
  for(let i = 1; i < sheet.data[0].rowData.length; i++) {
    const row = sheet.data[0].rowData[i];
    if(rowIsEmpty(row)){
      continue;
    } 
    const rowLanguage = row.values[1].userEnteredValue.stringValue;
    const rowAccount = row.values[0].userEnteredValue.stringValue;
    const postfix = row.values[2].userEnteredValue.stringValue;

    if(rowAccount === account && rowLanguage === language) {
      return postfix;
    }
  }

  return "ERROR";
}

function getThirdLevelDomainFromSheet(sheet, account) {
  //skip header
  for(let i = 1; i < sheet.data[0].rowData.length; i++) {
    const row = sheet.data[0].rowData[i];
    if(rowIsEmpty(row)){
      continue;
    } 
    const rowAccount = row.values[0].userEnteredValue.stringValue;
    const thirdLevelDomain = row.values[3].userEnteredValue.stringValue;

    if(rowAccount === account ) {
      return thirdLevelDomain;
    }
  }

  return "ERROR";
}

async function createAccountBuildoutSpreadsheet(keywordSpreadsheet, adCopySheet, urlDataSheet, account) {
  const rawHeaderRow = [
    "Campaign", "Ad Group", "Keyword", "Criterion Type", "Final URL", "Labels", "Ad type", "Status", "Description Line 1", "Description Line 2",
    "Headline 1", "Headline 1 position", "Headline 2", "Headline 3", "Path 1",
    "Headline 4", "Headline 5", "Headline 6", "Headline 7",
    "Headline 8", "Headline 9", "Headline 10", "Headline 11", "Headline 12", "Headline 13", "Headline 14", "Headline 15", "Description 1", "Description 1 position", "Description 2", "Description 3", "Description 4", "Max CPC", "Flexible Reach"
  ];
  //adds header row 
  let masterSpreadsheet = createSpreadSheet(account+ " Buildout", rawHeaderRow);
  const languages = getAccountLanguagesFromSheet(adCopySheet, account);
  const campaigns = getAccountCampaignsFromSheet(adCopySheet, account);

  //for every brand
  for (let i = 0; i < keywordSpreadsheet.sheets.length; i++) {
    for (let j = 0; j < languages.length; j++) {
      for (let k = 0; k < campaigns.length; k++) {
        const language = languages[j];
        const campaign = campaigns[k]
        const sheet = keywordSpreadsheet.sheets[i];
        const rawRowData = sheet.data[0].rowData;
        const rowData = rawRowData.slice(1); //removes header from each brand buildout
        
        //create new copy that we can edit
        let keywordRowData = [];
        //Update third level domain and campaign title
        for (let i = 0; i < rowData.length; i++) {
          //Campaign title
          if(isCellEmpty(rowData[i].values[0])) break;
          let newRowData = copyRowData(rowData[i])
          const rawCampaignTitle = newRowData.values[0].userEnteredValue.stringValue;
          
          const accountCampaignTitle = (rawCampaignTitle + " > " + account + " > " + campaign + " (" + language + ")");
          newRowData.values[0].userEnteredValue.stringValue = accountCampaignTitle;
          keywordRowData.push(newRowData);
        }

        const thirdLevelDomain = getThirdLevelDomainFromSheet(urlDataSheet, account)
        for(let i = 0; i < keywordRowData.length; i++) {
          //Final Url
          if (keywordRowData[i].values.length > 4 && keywordRowData[i].values[4].hasOwnProperty('userEnteredValue')) {
            const rawFinalURL = keywordRowData[i].values[4].userEnteredValue.stringValue;

            const accountFinalURL = rawFinalURL.replace("www.", thirdLevelDomain + ".");
            keywordRowData[i].values[4].userEnteredValue.stringValue = accountFinalURL;
          }
        }

        //Ad Groups
        const campaignTitle = keywordRowData[0].values[0].userEnteredValue.stringValue;
        const adGroupTitle = keywordRowData[0].values[1].userEnteredValue.stringValue;
        const finalURL = keywordRowData[0].values[4].userEnteredValue.stringValue;

        const adGroupRowData = createAdGroupRowData(campaignTitle, adGroupTitle, "Active", "Active", campaign)
        masterSpreadsheet.sheets[0].data[0].rowData.push(adGroupRowData);

        //Ads
        //TODO
        const adCopyRowData = getAdCopyRowData(adCopySheet, account, language, campaign);
        const brandTitle = sheet.properties.title;
        const path = createPath(brandTitle);
        
        for (let i = 0; i < adCopyRowData.length; i++) {
          const campaign = campaignTitle;
          const adGroup = adGroupTitle;

          const labels = !isCellEmpty(adCopyRowData[i].values[4]) ? adCopyRowData[i].values[4].userEnteredValue.stringValue : "";
          const adType = !isCellEmpty(adCopyRowData[i].values[5]) ? adCopyRowData[i].values[5].userEnteredValue.stringValue : "";
          const status = !isCellEmpty(adCopyRowData[i].values[6]) ? adCopyRowData[i].values[6].userEnteredValue.stringValue : "";
          const descriptionLine1 = !isCellEmpty(adCopyRowData[i].values[7]) ? adCopyRowData[i].values[7].userEnteredValue.stringValue : "";
          const descriptionLine2 = !isCellEmpty(adCopyRowData[i].values[8]) ? adCopyRowData[i].values[8].userEnteredValue.stringValue : "";
        
          const headline1 = !isCellEmpty(adCopyRowData[i].values[9]) ? createHeadline1(brandTitle, adCopyRowData[i].values[9].userEnteredValue.stringValue) : "";
          const headline1Position = !isCellEmpty(adCopyRowData[i].values[10])
            ? (adCopyRowData[i].values[10].userEnteredValue.stringValue ?? adCopyRowData[i].values[10].userEnteredValue.numberValue ?? "")
            : "";
        
          const headline2 = !isCellEmpty(adCopyRowData[i].values[11]) ? adCopyRowData[i].values[11].userEnteredValue.stringValue : "";
          const headline3 = !isCellEmpty(adCopyRowData[i].values[12]) ? adCopyRowData[i].values[12].userEnteredValue.stringValue : "";
          const headline4 = !isCellEmpty(adCopyRowData[i].values[13]) ? adCopyRowData[i].values[13].userEnteredValue.stringValue : "";
          const headline5 = !isCellEmpty(adCopyRowData[i].values[14]) ? adCopyRowData[i].values[14].userEnteredValue.stringValue : "";
          const headline6 = !isCellEmpty(adCopyRowData[i].values[15]) ? adCopyRowData[i].values[15].userEnteredValue.stringValue : "";
          const headline7 = !isCellEmpty(adCopyRowData[i].values[16]) ? adCopyRowData[i].values[16].userEnteredValue.stringValue : "";
          const headline8 = !isCellEmpty(adCopyRowData[i].values[17]) ? adCopyRowData[i].values[17].userEnteredValue.stringValue : "";
          const headline9 = !isCellEmpty(adCopyRowData[i].values[18]) ? adCopyRowData[i].values[18].userEnteredValue.stringValue : "";
          const headline10 = !isCellEmpty(adCopyRowData[i].values[19]) ? adCopyRowData[i].values[19].userEnteredValue.stringValue : "";
          const headline11 = !isCellEmpty(adCopyRowData[i].values[20]) ? adCopyRowData[i].values[20].userEnteredValue.stringValue : "";
          const headline12 = !isCellEmpty(adCopyRowData[i].values[21]) ? adCopyRowData[i].values[21].userEnteredValue.stringValue : "";
          const headline13 = !isCellEmpty(adCopyRowData[i].values[22]) ? adCopyRowData[i].values[22].userEnteredValue.stringValue : "";
          const headline14 = !isCellEmpty(adCopyRowData[i].values[23]) ? adCopyRowData[i].values[23].userEnteredValue.stringValue : "";
          const headline15 = !isCellEmpty(adCopyRowData[i].values[24]) ? adCopyRowData[i].values[24].userEnteredValue.stringValue : "";
        
          const description1 = !isCellEmpty(adCopyRowData[i].values[25]) ? adCopyRowData[i].values[25].userEnteredValue.stringValue : "";
          const description1Position = !isCellEmpty(adCopyRowData[i].values[26])
            ? (adCopyRowData[i].values[26].userEnteredValue.stringValue ?? adCopyRowData[i].values[26].userEnteredValue.numberValue ?? "")
            : "";
          const description2 = !isCellEmpty(adCopyRowData[i].values[27]) ? adCopyRowData[i].values[27].userEnteredValue.stringValue : "";
          const description3 = !isCellEmpty(adCopyRowData[i].values[28]) ? adCopyRowData[i].values[28].userEnteredValue.stringValue : "";
          const description4 = !isCellEmpty(adCopyRowData[i].values[29]) ? adCopyRowData[i].values[29].userEnteredValue.stringValue : "";
        
          const adRowValues = [
            campaign,
            adGroup,
            "", "", // keyword, criterion
            finalURL,
            labels,
            adType,
            status,
            descriptionLine1,
            descriptionLine2,
            headline1,
            headline1Position,
            headline2,
            headline3,
            path, // still using 'path' as you did before
            headline4,
            headline5,
            headline6,
            headline7,
            headline8,
            headline9,
            headline10,
            headline11,
            headline12,
            headline13,
            headline14,
            headline15,
            description1,
            description1Position,
            description2,
            description3,
            description4,
            "", // Max CPC
            "" // flexibleReach
          ];
        
          const adRow = createRowData(adRowValues);
          handleFieldLengthLimits(adRow);
          masterSpreadsheet.sheets[0].data[0].rowData.push(adRow);
        }        
      
        
        //Keywords  
        // add country specific keyword postfix
        for(let i = 0; i < keywordRowData.length; i++) { 
          if (!isCellEmpty(keywordRowData[i].values[4])) {
            const rawFinalURL = keywordRowData[i].values[4].userEnteredValue.stringValue;
            
            const rawPostfix = getPostfixFromSheet(urlDataSheet, account, language);
            const postfix = (rawFinalURL.indexOf('?') >= 0 ? "&" + rawPostfix : "?" + rawPostfix);
            const accountFinalURL = rawFinalURL.length < 5 ? " " : rawFinalURL + postfix;
            keywordRowData[i].values[4].userEnteredValue.stringValue = accountFinalURL; 
          } 

          masterSpreadsheet.sheets[0].data[0].rowData.push(keywordRowData[i]);
        }
      }
    }
  }
  return masterSpreadsheet;
}

function createHeadline1(brandTitle, headline1) {
  const index = headline1.indexOf('Brand');
  const newString = headline1.substr(0, index) + brandTitle + headline1.substr(index + 5);
  return newString;
}

function createPath(brandTitle) {
  let path = "";
  if(brandTitle.indexOf(" ") > -1) {
    // TODOL remove special charachters
    const tokens = brandTitle.split(/(\s+)/);
    
    //ads token with a "-" delimeter
    for (let i = 0; i < tokens.length - 1; i++) {
      if (tokens[i] === " ") continue;
      path += tokens[i] + "-"
    }

    //adds final token without delimiter
    path += tokens[tokens.length - 1];
  } else {
    path = brandTitle;
  }
  if(path.length > 15) {
    console.log("ERROR: PATH OVER 15 CHARACHTERS PLEASE CHECK")
  }

  return path;
}

function isCellEmpty(cell) {
  return typeof cell === "undefined" || !cell.hasOwnProperty('userEnteredValue')
}

/**potentially hazasrdous */
function createAdGroupRowData(campaign, adGroup, campaignStatus, adGroupStatus, campaignType) {
  const flexibleReach = (campaignType === "Acquisition" || campaignType === "Broad")
    ? "Audience segments;Genders;Ages;Parental status;Household incomes"
    : "Genders;Ages;Parental status;Household incomes";
  return createRowData([
    campaign, adGroup,
    "", "", // keyword, criterion
    "", "", // Final URL, Labels
    "", "", // Ad type, Status
    "", "", // Description Line 1, Description Line 2
    "", "", // Headline 1, Headline 1 position
    "", "", // Headline 2, Headline 3
    "", // Path 1
    "", "", "", "", "", "", "", "", "", "", // Headline 4-15
    "", "", // Description 1, Description 1 position
    "", "", // Description 2, Description 3
    "", "", "", // Description 4, (add one more empty string)
    "0.5", flexibleReach // Max CPC, Flexible Reach
  ]);
}

function indexIncluded (indexes, index) {
  for (let i = 0; i < indexes.length; i++) {
    if (indexes[i].campaign === index.campaign && indexes[i].adGroup === index.adGroup) {
      return true;
    }
  }

  return false;
}

function indexToA1(row, col) {
  const rowNum = row + 1;
  const colLetter = toLetters(col + 1) ;

  return colLetter + String(rowNum);
}

function toLetters(num) {
  "use strict";
  var mod = num % 26,
      pow = num / 26 | 0,
      out = mod ? String.fromCharCode(64 + mod) : (--pow, 'Z');
  return pow ? toLetters(pow) + out : out;
}

function getDocumentIdFromUrl (url) {
  const start = url.search("/spreadsheets/d/") + 16;
  const end = start + 44;
  const documentId = url.slice(start, end);

  return documentId;
}

// Helper to fetch Google Sheets data using access token
async function fetchSpreadsheetNoGridData(url) {
  const spreadsheetId = getDocumentIdFromUrl(url);
  const apiUrl = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}?includeGridData=false`;
  if (!window.accessToken) {
    alert("Google access token not available. Please sign in again.");
    return null;
  }
  const res = await fetch(apiUrl, {
    headers: {
      'Authorization': `Bearer ${window.accessToken}`
    }
  });
  if (!res.ok) {
    const err = await res.text();
    console.error('Sheets API error:', err);
    alert('Sheets API error: ' + err);
    return null;
  }
  return await res.json();
}

// Helper to fetch a spreadsheet with grid data using access token
async function fetchSpreadsheet(url) {
  const spreadsheetId = getDocumentIdFromUrl(url);
  const apiUrl = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}?includeGridData=true`;
  if (!window.accessToken) {
    alert("Google access token not available. Please sign in again.");
    return null;
  }
  const res = await fetch(apiUrl, {
    headers: {
      'Authorization': `Bearer ${window.accessToken}`
    }
  });
  if (!res.ok) {
    const err = await res.text();
    console.error('Sheets API error:', err);
    alert('Sheets API error: ' + err);
    return null;
  }
  return await res.json();
}

// Helper to fetch a single manager's sheet and URL Data using access token
async function fetchSpreadsheetSingleManager(url, manager) {
  const spreadsheetId = getDocumentIdFromUrl(url);
  const apiUrl = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}?ranges=${encodeURIComponent(manager + '!A1:AD150')}&ranges=URL%20Data!A1:D150&includeGridData=true`;
  if (!window.accessToken) {
    alert("Google access token not available. Please sign in again.");
    return null;
  }
  const res = await fetch(apiUrl, {
    headers: {
      'Authorization': `Bearer ${window.accessToken}`
    }
  });
  if (!res.ok) {
    const err = await res.text();
    console.error('Sheets API error:', err);
    alert('Sheets API error: ' + err);
    return null;
  }
  return await res.json();
}

function transformFetchedSpreadsheetToCreationFormat(fetchedSpreadsheet) {
  if (!fetchedSpreadsheet || !fetchedSpreadsheet.sheets) {
    console.error("transformFetchedSpreadsheetToCreationFormat: Invalid fetched spreadsheet (missing .sheets).");
    return null;
  }
  // Create a new object with properties and a sheets array (each sheet has a .data property).
  const creationSpreadsheet = {
    properties: { title: fetchedSpreadsheet.properties?.title || "Untitled" },
    sheets: fetchedSpreadsheet.sheets.map(sheet => ({
      properties: { title: sheet.properties?.title || "Sheet1" },
      data: sheet.data || []
    }))
  };
  return creationSpreadsheet;
}

function createInsertDimensionRequest(sheetId, index, numRows) {
  const request = {
    insertDimension: {
      range: {
        sheetId,
        dimension: "ROWS",
        startIndex: index,
        endIndex: index + numRows
      },
      inheritFromBefore: true
    }
  }

  return request;
}

function createAppendCellsRequest(sheetId, rows) {
  return {
    appendCells: {
      sheetId,
      rows,
      fields:"userEnteredValue/stringValue"
    }
  }
}

function createDeleteDimensionRequest(sheetId, index, numCells, dimension) {
  return {
    deleteDimension: {
      range: {
        sheetId,
        dimension,
        startIndex: index,
        endIndex: index + numCells
      }
    }
  }
}

function createUpdateCellsRequest(sheetId, index, values) {
  const row = createRowData(values);
  return {
    updateCells: {
      rows: [
        row
      ],
      fields: "userEnteredValue/stringValue",
      range: {
        sheetId,
        startColumnIndex: 0,
        endColumnIndex: values.length,
        startRowIndex: index,
        endRowIndex: index + 1
      }
    }
  }
}

/**
 * Batch update wrapper
 */
function batchUpdate(spreadsheetId, requests) {
  gapi.client.sheets.spreadsheets.batchUpdate({
    spreadsheetId                        
  }, {
    requests
  }).then((response) => {
    console.log(response);
  });
}
function copyRowData(row) {
  let values = [];
  for(let i = 0; i < row.values.length; i++) {
    if(row.values[i].hasOwnProperty("userEnteredValue")) {
      values.push(row.values[i].userEnteredValue.stringValue);
    }
  }

  return createRowData(values);
}
//Takes an array of Row Data
function createRowData(values){
  var rowData = {
    values: []
  } 
  for (value in values){
    if(typeof values[value] === "string") {
      rowData.values.push({
        userEnteredValue: {
          stringValue: values[value]
        },
        userEnteredFormat: {
          backgroundColor: {
            red: 1,
            green: 1,
            blue: 1
          }
        } 
      })
    } else if(typeof values[value] === "number") {
      rowData.values.push({
        userEnteredValue: {
          numberValue: values[value]
        },
        userEnteredFormat: {
          backgroundColor: {
            red: 1,
            green: 1,
            blue: 1
          }
        } 
      })
    }
    
  }

  return rowData;
}

function handleFieldLengthLimits(rowData) {
  //10 and 13
  const headline1 = rowData.values[10].userEnteredValue.stringValue;
  if(headline1.length > 40) {
    markCellRed(rowData.values[10])
  }

  const path1 = rowData.values[14].userEnteredValue.stringValue;
  if(path1.length > 15) {
    markCellRed(rowData.values[14])
  }
  
}

//objects are passed by reference
function markCellRed(cell) {
  cell.userEnteredFormat.backgroundColor.green = 0;
  cell.userEnteredFormat.backgroundColor.blue = 0;
}

async function createNewDocument(spreadsheet) {
  console.log('createNewDocument received:', spreadsheet);
  if (!spreadsheet) {
    alert("createNewDocument: spreadsheet is null or undefined.");
    console.error("createNewDocument: spreadsheet is null or undefined.");
    return null;
  }
  // Transform if not in creation format
  if (!Array.isArray(spreadsheet.sheets) || (spreadsheet.sheets.length > 0 && !spreadsheet.sheets[0].data)) {
    console.warn("createNewDocument: spreadsheet does not have a .sheets array (or its sheets do not have a .data property). Transforming using transformFetchedSpreadsheetToCreationFormat.");
    spreadsheet = transformFetchedSpreadsheetToCreationFormat(spreadsheet);
    if (!spreadsheet) {
      alert("createNewDocument: transformation failed. Cannot create new document.");
      console.error("createNewDocument: transformation failed.");
      return null;
    }
  }
  const firstSheet = spreadsheet.sheets[0];
  console.log('First sheet object (after transformation if needed):', firstSheet);
  if (!firstSheet) {
    alert("createNewDocument: first sheet is undefined. Cannot create new document.");
    console.error("createNewDocument: first sheet is undefined.");
    return null;
  }
  let res = null;
  await gapi.client.sheets.spreadsheets.create(spreadsheet)
    .then((response) => { res = response.result; });
  return res;
}

function createSpreadSheet(title, headers) {
  var spreadSheet = {
      properties: {
        title
      },
      sheets: [{
        data: [{
          //header row
          rowData: [createRowData(headers)]
        }]
      }]
    }

  return spreadSheet;
}

function createBrandBuildoutTemplateSpreadsheet(campaign, adGroup, baseKeyword, finalUrl){
var spreadSheet = createSpreadSheet(baseKeyword, ["Campaign", "Ad Group", "Keyword", "Match Type", "Final URL"]);

//Product Keywords
for(keyword in PRODUCT_KEYWORDS) {
  var keywordString = PRODUCT_KEYWORDS[keyword].prefix + baseKeyword + PRODUCT_KEYWORDS[keyword].suffix;
  var row = createRowData([campaign, adGroup, keywordString, "Phrase", finalUrl]);
  spreadSheet.sheets[0].data[0].rowData.push(row);
}

//Domain Keywords
//check for simple domain
if (baseKeyword.indexOf(" ") < 0) {
  for(keyword in SIMPLE_DOMAIN_KEYWORDS) {
    var keywordString = SIMPLE_DOMAIN_KEYWORDS[keyword].prefix + baseKeyword + SIMPLE_DOMAIN_KEYWORDS[keyword].suffix;
    var row = createRowData([campaign, adGroup, keywordString, "Phrase", finalUrl]);
    spreadSheet.sheets[0].data[0].rowData.push(row);
  }
} else {
  for(keyword in COMPOUND_DOMAIN_KEYWORDS) {
    var keywordString = COMPOUND_DOMAIN_KEYWORDS[keyword].prefix + baseKeyword + COMPOUND_DOMAIN_KEYWORDS[keyword].suffix;
    var row = createRowData([campaign, adGroup, keywordString, "Phrase", finalUrl]);
    spreadSheet.sheets[0].data[0].rowData.push(row);
  }

  //remove whitespace
  var baseKeywordNoSpace = baseKeyword.replace(/\s+/g, '');

  for(keyword in SIMPLE_DOMAIN_KEYWORDS) {
    var keywordString = SIMPLE_DOMAIN_KEYWORDS[keyword].prefix + baseKeywordNoSpace + SIMPLE_DOMAIN_KEYWORDS[keyword].suffix;
    var row = createRowData([campaign, adGroup, keywordString, "Phrase", finalUrl]);
    spreadSheet.sheets[0].data[0].rowData.push(row);
  }
}

//Empty Brand Keywords
for (var i =0; i < 50; i++) {
  var row = createRowData([campaign, adGroup, baseKeyword, "Phrase", finalUrl]);
  spreadSheet.sheets[0].data[0].rowData.push(row);
}

//Negative Keywords

for (keyword in NEGATIVE_KEYWORDS) {
  var row = createRowData([campaign, adGroup, NEGATIVE_KEYWORDS[keyword].keyword, NEGATIVE_KEYWORDS[keyword].match_type, ""])
  spreadSheet.sheets[0].data[0].rowData.push(row);
}

//Empty Negative Keywords

for (var i =0; i < 5; i++) {
  var row = createRowData([campaign, adGroup, "", "Negative Broad", ""]);
  spreadSheet.sheets[0].data[0].rowData.push(row);
}

return spreadSheet;
}

function consoleLogSpreadSheet(spreadSheet) {
  const rowData = spreadSheet.sheets[0].data[0].rowData;
  for (const rowIndex in rowData) {
    const row = rowData[rowIndex].values;
    var rowText = "";  
    for (cellIndex in row) {
      const cell = row[cellIndex];

      if (!cell.hasOwnProperty('userEnteredValue')) {
        rowText += " ";
      } else {
        rowText += cell.userEnteredValue.stringValue + " ";
      }
    }
    console.log(rowText)
  }
}

function getManagersFromDataSpreadsheet(dataSpreadsheet) {
  if (!dataSpreadsheet || !dataSpreadsheet.sheets) {
    alert('Failed to load spreadsheet data. Please check your access and try again.');
    return [];
  }
  let managers = [];
  for(let i = 0; i < dataSpreadsheet.sheets.length; i++) {
    const sheet = dataSpreadsheet.sheets[i];
    if(sheet.properties.title !== "URL Data") managers.push(sheet.properties.title);
  }
  return managers;
}

function getManagerSheet(dataSpreadsheet, manager) {
  if (!dataSpreadsheet || !Array.isArray(dataSpreadsheet.sheets)) {
    console.error('getManagerSheet: Invalid dataSpreadsheet:', dataSpreadsheet);
    alert('Manager sheet data is not available. Please try again.');
    return null;
  }
  const availableSheetNames = dataSpreadsheet.sheets.map(s => s.properties.title);
  console.log('getManagerSheet: available sheets:', availableSheetNames, 'looking for manager:', manager);
  let managerSheet = null;
  for(let i = 0; i < dataSpreadsheet.sheets.length; i++) {
    if(dataSpreadsheet.sheets[i].properties.title === manager) {
      managerSheet = dataSpreadsheet.sheets[i];
      break;
    }
  }
  if (!managerSheet) {
    console.warn('getManagerSheet: Manager sheet not found for manager:', manager);
    alert('Manager sheet not found: ' + manager + '. Available sheets: ' + availableSheetNames.join(', '));
  }
  return managerSheet;
}

function createManagerHtml(managers) {
  let managerHtml = '<select class="form-select" id="manager-select">';
  for (let i = 0; i < managers.length; i++) {        
    const template = `<option value="${managers[i]}">${managers[i]}</option>`;
    managerHtml += template;
  }
  managerHtml += '</select>';
  return managerHtml;
}

function createAccountHtml(accounts) {
  let html = "";
  for (let i = 0; i < accounts.length; i++) {
    const account = accounts[i].accountTitle || accounts[i];
    const id = account.split(" ").join("") + "-checkbox";
    const template =
      `<div class="input-group mb-1">
        <div class="input-group-text">
          <input id="${id}" class="form-check-input mt-0" type="checkbox" value="" aria-label="Checkbox for following text input">
        </div>
        <span class="input-group-text">${account}</span>
      </div>`;
    html += template;
  }
  return html;
}