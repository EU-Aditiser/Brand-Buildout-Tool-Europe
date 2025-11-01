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
  // Validate input spreadsheets
  const buildoutValidation = validateSpreadsheetData(buildoutSpreadsheet, 'Brand Buildout Spreadsheet');
  const accountDataValidation = validateSpreadsheetData(accountDataSpreadsheet, 'Account Data Spreadsheet');
  
  const buildoutHasIssues = displayValidationResults(buildoutValidation, 'Brand Buildout Spreadsheet');
  const accountDataHasIssues = displayValidationResults(accountDataValidation, 'Account Data Spreadsheet');
  
  // Ask user if they want to continue with issues
  if (buildoutHasIssues || accountDataHasIssues) {
    // const continueProcessing = confirm('Data validation found issues. Do you want to continue processing anyway? (This may cause errors)');
    // if (!continueProcessing) {
    //   return;
    // }
    // Instead, just stop processing if there are issues
    return;
  }

  if (!buildoutSpreadsheet || !buildoutSpreadsheet.sheets) {
    alert('Failed to load the brand buildout spreadsheet. Please check your access and try again.');
    return;
  }
  if (!accountDataSpreadsheet || !accountDataSpreadsheet.sheets) {
    alert('Failed to load the account data spreadsheet. Please check your access and try again.');
    return;
  }

  const urlDataSheet = getUrlDataSheet(accountDataSpreadsheet);
  if (!urlDataSheet) {
    alert('Failed to find URL Data sheet. Please check the spreadsheet structure.');
    return;
  }

  let spreadsheets = [];
  const managerSheet = getManagerSheet(accountDataSpreadsheet, MANAGER);
  if (!managerSheet) {
    alert('Failed to get manager sheet data. Please try again.');
    return;
  }

  for(let i = 0; i < accounts.length; i++) {
    try {
      const accountBuildoutSpreadsheet = await createAccountBuildoutSpreadsheet(buildoutSpreadsheet, managerSheet, urlDataSheet, accounts[i]);
      if (!accountBuildoutSpreadsheet) {
        continue;
      }
      spreadsheets.push(accountBuildoutSpreadsheet);
    } catch (error) {
      console.error('Error creating buildout spreadsheet for account:', accounts[i], 'Error:', error);
      throw new Error(`Failed to create buildout spreadsheet for account ${accounts[i]}: ${error.message}`);
    }
  }

  if (spreadsheets.length === 0) {
    alert('No spreadsheets could be created. Please check the console for details.');
    return;
  }

  let createdUrls = [];
  for(let i = 0; i < spreadsheets.length; i++) {
    if (!spreadsheets[i]) {
      continue;
    }
    try {
      const spreadsheetUrl = await createNewDocument(spreadsheets[i]);
      if (spreadsheetUrl) {
        createdUrls.push(spreadsheetUrl);
      }
    } catch (error) {
      console.error('Error creating spreadsheet:', error);
      alert('Failed to create spreadsheet: ' + error.message);
    }
  }

  // Open all created spreadsheets in new tabs
  if (createdUrls.length > 0) {
    createdUrls.forEach(url => {
      if (url && url.startsWith('https://')) {
        window.open(url, '_blank', 'noopener,noreferrer');
      }
    });
  } else {
    alert('No spreadsheets could be created. Please check the console for details.');
  }

  // Don't reload the page immediately to allow console logs to be visible
  setTimeout(() => {
    location.reload();
  }, 2000);
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
  // Creating buildout spreadsheet for account: account

  if (!keywordSpreadsheet || !keywordSpreadsheet.sheets) {
    console.error('createAccountBuildoutSpreadsheet: Invalid keywordSpreadsheet:', keywordSpreadsheet);
    return null;
  }
  if (!adCopySheet || !adCopySheet.data) {
    console.error('createAccountBuildoutSpreadsheet: Invalid adCopySheet:', adCopySheet);
    return null;
  }
  if (!urlDataSheet || !urlDataSheet.data) {
    console.error('createAccountBuildoutSpreadsheet: Invalid urlDataSheet:', urlDataSheet);
    return null;
  }

  // Update header row to include all required columns
  const rawHeaderRow = [
    "Campaign", "Ad Group", "Keyword", "Criterion Type", "Final URL", "Labels", "Ad type", "Status",
    "Description Line 1", "Description Line 2",
    "Headline 1", "Headline 1 position", "Headline 2", "Headline 3", "Path 1",
    "Headline 4", "Headline 5", "Headline 6", "Headline 7", "Headline 8", "Headline 9", "Headline 10",
    "Headline 11", "Headline 12", "Headline 13", "Headline 14", "Headline 15",
    "Description 1", "Description 1 position", "Description 2", "Description 3", "Description 4",
    "Max CPC", "Flexible Reach"
  ];
  
  // Create the master spreadsheet with proper structure
  let masterSpreadsheet = {
    properties: {
      title: account + " Buildout"
    },
    sheets: [{
      properties: {
        title: "Sheet1",
        sheetId: 0,
        index: 0,
        sheetType: "GRID",
        gridProperties: {
          rowCount: 1000,
          columnCount: rawHeaderRow.length // Set column count based on actual header length
        }
      },
      data: [{
        rowData: [createRowData(rawHeaderRow)],
        startRow: 0,
        startColumn: 0
      }]
    }]
  };

  const languages = getAccountLanguagesFromSheet(adCopySheet, account);
  const campaigns = getAccountCampaignsFromSheet(adCopySheet, account);

  // Found languages and campaigns for account

  if (!languages.length || !campaigns.length) {
    console.error('createAccountBuildoutSpreadsheet: No languages or campaigns found for account:', account);
    return null;
  }

  //for every brand
  for (let i = 0; i < keywordSpreadsheet.sheets.length; i++) {
    for (let j = 0; j < languages.length; j++) {
      for (let k = 0; k < campaigns.length; k++) {
        try {
        const language = languages[j];
        const campaign = campaigns[k]
        const sheet = keywordSpreadsheet.sheets[i];
        const rawRowData = sheet.data[0].rowData;
        const rowData = rawRowData.slice(1); //removes header from each brand buildout
        
        //create new copy that we can edit
        let keywordRowData = [];
        //Update third level domain and campaign title
        for (let rowIndex = 0; rowIndex < rowData.length; rowIndex++) {
          //Campaign title
          if(isCellEmpty(rowData[rowIndex].values[0])) break;
          let newRowData = copyRowData(rowData[rowIndex])
          
          const rawCampaignTitle = newRowData.values[0].userEnteredValue.stringValue;
          
          const accountCampaignTitle = (rawCampaignTitle + " > " + account + " > " + campaign + " (" + language + ")");
          newRowData.values[0].userEnteredValue.stringValue = accountCampaignTitle;
          keywordRowData.push(newRowData);
        }

        const thirdLevelDomain = getThirdLevelDomainFromSheet(urlDataSheet, account)
        for(let urlIndex = 0; urlIndex < keywordRowData.length; urlIndex++) {
          //Final Url
          if (keywordRowData[urlIndex].values.length > 4 && keywordRowData[urlIndex].values[4].hasOwnProperty('userEnteredValue')) {
            const rawFinalURL = keywordRowData[urlIndex].values[4].userEnteredValue.stringValue;

            // Handle both https://www. and www. formats
            let accountFinalURL;
            if (rawFinalURL.includes('https://www.')) {
              accountFinalURL = rawFinalURL.replace('https://www.', `https://${thirdLevelDomain}.`);
            } else if (rawFinalURL.includes('http://www.')) {
              accountFinalURL = rawFinalURL.replace('http://www.', `http://${thirdLevelDomain}.`);
            } else if (rawFinalURL.includes('www.')) {
              accountFinalURL = rawFinalURL.replace('www.', `${thirdLevelDomain}.`);
            } else {
              // If no www. found, keep the original URL
              accountFinalURL = rawFinalURL;
            }
            
            keywordRowData[urlIndex].values[4].userEnteredValue.stringValue = accountFinalURL;
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
        
        for (let adIndex = 0; adIndex < adCopyRowData.length; adIndex++) {
          const campaign = campaignTitle;
          const adGroup = adGroupTitle;

          const labels = !isCellEmpty(adCopyRowData[adIndex].values[4]) ? adCopyRowData[adIndex].values[4].userEnteredValue.stringValue : "";
          const adType = !isCellEmpty(adCopyRowData[adIndex].values[5]) ? adCopyRowData[adIndex].values[5].userEnteredValue.stringValue : "";
          const status = !isCellEmpty(adCopyRowData[adIndex].values[6]) ? adCopyRowData[adIndex].values[6].userEnteredValue.stringValue : "";
          const descriptionLine1 = !isCellEmpty(adCopyRowData[adIndex].values[7]) ? adCopyRowData[adIndex].values[7].userEnteredValue.stringValue : "";
          const descriptionLine2 = !isCellEmpty(adCopyRowData[adIndex].values[8]) ? adCopyRowData[adIndex].values[8].userEnteredValue.stringValue : "";
        
          // Headline columns in the adCopySheet row (0-based):
          // H1 at index 9, H1 position at 10, H2..H3 at 11..12, H4..H15 at 13..24
          const headline1 = transformHeadlineCell(brandTitle, adCopyRowData[adIndex].values[9]);
          const headline1PositionValue = !isCellEmpty(adCopyRowData[adIndex].values[10])
            ? (adCopyRowData[adIndex].values[10].userEnteredValue.stringValue
               ?? adCopyRowData[adIndex].values[10].userEnteredValue.numberValue
               ?? '')
            : '';
        
          // Build Headline 2..15 using the same rule
          const headlinesRest = [];
          for (let srcCol = 11; srcCol <= 24; srcCol++) {
            headlinesRest.push(transformHeadlineCell(brandTitle, adCopyRowData[adIndex].values[srcCol]));
          }
        
          // Unpack for readability
          const [headline2, headline3, headline4, headline5, headline6, headline7, headline8, headline9,
                 headline10, headline11, headline12, headline13, headline14, headline15] = headlinesRest;
        
          const description1 = !isCellEmpty(adCopyRowData[adIndex].values[25]) ? adCopyRowData[adIndex].values[25].userEnteredValue.stringValue : "";
          const description1Position = !isCellEmpty(adCopyRowData[adIndex].values[26])
            ? (adCopyRowData[adIndex].values[26].userEnteredValue.stringValue ?? adCopyRowData[adIndex].values[26].userEnteredValue.numberValue ?? "")
            : "";
          const description2 = !isCellEmpty(adCopyRowData[adIndex].values[27]) ? adCopyRowData[adIndex].values[27].userEnteredValue.stringValue : "";
          const description3 = !isCellEmpty(adCopyRowData[adIndex].values[28]) ? adCopyRowData[adIndex].values[28].userEnteredValue.stringValue : "";
          const description4 = !isCellEmpty(adCopyRowData[adIndex].values[29]) ? adCopyRowData[adIndex].values[29].userEnteredValue.stringValue : "";
        
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
            headline1PositionValue,
            headline2,
            headline3,
            path,
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
            "" // Flexible Reach
          ];
        
          const adRow = createRowData(adRowValues);
          handleFieldLengthLimits(adRow);
          masterSpreadsheet.sheets[0].data[0].rowData.push(adRow);
        }        
      
        
        //Keywords  
        // add country specific keyword postfix
        for(let keywordIndex = 0; keywordIndex < keywordRowData.length; keywordIndex++) { 
          if (!isCellEmpty(keywordRowData[keywordIndex].values[4])) {
            const rawFinalURL = keywordRowData[keywordIndex].values[4].userEnteredValue.stringValue;
            
            const rawPostfix = getPostfixFromSheet(urlDataSheet, account, language);
            const postfix = (rawFinalURL.indexOf('?') >= 0 ? "&" + rawPostfix : "?" + rawPostfix);
            const accountFinalURL = rawFinalURL.length < 5 ? " " : rawFinalURL + postfix;
            keywordRowData[keywordIndex].values[4].userEnteredValue.stringValue = accountFinalURL; 
          } 

          masterSpreadsheet.sheets[0].data[0].rowData.push(keywordRowData[keywordIndex]);
        }
        } catch (error) {
          console.error('createAccountBuildoutSpreadsheet: Error processing brand:', sheet.properties.title, 'campaign:', campaign, 'language:', language, 'Error:', error);
          // Continue with next iteration instead of failing completely
        }
      }
    }
  }
  // Final spreadsheet structure created
  return masterSpreadsheet;
}

// DEPRECATED: Use transformHeadlineCell instead
// function createHeadline1(brandTitle, headline1) {
//   const index = headline1.indexOf('Brand');
//   const newString = headline1.substr(0, index) + brandTitle + headline1.substr(index + 5);
//   return newString;
// }

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
  return typeof cell === "undefined" || cell === null || !cell.hasOwnProperty('userEnteredValue')
}

// Exact token we support
const KEYWORD_BRAND_TOKEN = '{KeyWord:Brand at iHerb}';

// Returns the plain string value from a cell ('' if empty)
function readCellString(cell) {
  if (isCellEmpty(cell)) return '';
  const v = cell.userEnteredValue;
  // prefer string, fall back to number coerced to string
  return typeof v.stringValue === 'string'
    ? v.stringValue
    : (typeof v.numberValue === 'number' ? String(v.numberValue) : '');
}

// Replace only when the cell equals the exact token. Otherwise pass through unchanged.
function transformHeadlineCell(brandTitle, cell) {
  const s = readCellString(cell).trim();
  if (s === KEYWORD_BRAND_TOKEN) {
    return `{KeyWord:${brandTitle} at iHerb}`;
  }
  return s; // unchanged
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
  console.log('fetchSpreadsheetNoGridData called with URL:', url);
  
  const spreadsheetId = getDocumentIdFromUrl(url);
  console.log('Extracted spreadsheet ID:', spreadsheetId);
  
  const apiUrl = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}?includeGridData=false`;
  console.log('API URL:', apiUrl);
  
  // Get current access token with multiple fallbacks
  let token = window.accessToken;
  if (!token) {
    token = localStorage.getItem('accessToken');
  }
  if (!token) {
    // Try to get from global function if available
    if (typeof getCurrentAccessToken === 'function') {
      token = await getCurrentAccessToken();
    }
  }
  
  if (!token) {
    console.error('No access token available');
    alert("Google access token not available. Please sign in again.");
    return null;
  }
  
  console.log('Token available for API call:', token ? 'Yes' : 'No');
  console.log('Token length:', token ? token.length : 0);
  
  try {
    console.log('Making API request...');
    const res = await fetch(apiUrl, {
      headers: {
        'Authorization': `Bearer ${token}`
      }
    });
    
    console.log('API response status:', res.status);
    console.log('API response ok:', res.ok);
    
    if (!res.ok) {
      const err = await res.text();
      console.error('Sheets API error:', err);
      alert('Sheets API error: ' + err);
      return null;
    }
    
    const result = await res.json();
    console.log('API response result:', result);
    console.log('Result has sheets property:', result && result.sheets ? 'Yes' : 'No');
    
    return result;
  } catch (error) {
    console.error('Error in fetchSpreadsheetNoGridData:', error);
    alert('Error fetching spreadsheet: ' + error.message);
    return null;
  }
}

// Helper to fetch a spreadsheet with grid data using access token
async function fetchSpreadsheet(url) {
  if (!url || url.trim() === '') {
    alert('Please provide a valid spreadsheet URL.');
    return null;
  }
  
  const spreadsheetId = getDocumentIdFromUrl(url);
  if (!spreadsheetId) {
    alert('Invalid spreadsheet URL. Please check the URL and try again.');
    return null;
  }
  
  const apiUrl = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}?includeGridData=true`;
  
  // Get current access token with multiple fallbacks
  let token = window.accessToken;
  if (!token) {
    token = localStorage.getItem('accessToken');
  }
  if (!token) {
    // Try to get from global function if available
    if (typeof getCurrentAccessToken === 'function') {
      token = await getCurrentAccessToken();
    }
  }
  
  if (!token) {
    alert("Google access token not available. Please sign in again.");
    return null;
  }
  
  console.log('Using token for API call:', token ? 'Token available' : 'No token');
  
  try {
    const res = await fetch(apiUrl, {
      headers: {
        'Authorization': `Bearer ${token}`
      }
    });
    
    if (!res.ok) {
      const err = await res.text();
      console.error('Sheets API error:', err);
      alert('Sheets API error: ' + err);
      return null;
    }
    
    const result = await res.json();
    
    // Validate the fetched data
    if (result && result.properties && result.properties.title) {
      const validation = validateSpreadsheetData(result, result.properties.title);
      if (validation.errors.length > 0 || validation.warnings.length > 0) {
        displayValidationResults(validation, result.properties.title);
      }
    }
    
    return result;
  } catch (error) {
    console.error('Failed to fetch spreadsheet:', error);
    alert('Failed to fetch spreadsheet: ' + error.message);
    return null;
  }
}

// Helper to fetch a single manager's sheet and URL Data using access token
async function fetchSpreadsheetSingleManager(url, manager) {
  if (!url || url.trim() === '') {
    alert('Please provide a valid spreadsheet URL.');
    return null;
  }
  
  if (!manager || manager.trim() === '') {
    alert('Please select a manager.');
    return null;
  }
  
  const spreadsheetId = getDocumentIdFromUrl(url);
  if (!spreadsheetId) {
    alert('Invalid spreadsheet URL. Please check the URL and try again.');
    return null;
  }
  
  // Update range to include all columns up to AH (column 34) to cover all headlines and descriptions
  const apiUrl = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}?ranges=${encodeURIComponent(manager + '!A1:AH100')}&ranges=URL%20Data!A1:D150&includeGridData=true`;
  
  // Get current access token with multiple fallbacks
  let token = window.accessToken;
  if (!token) {
    token = localStorage.getItem('accessToken');
  }
  if (!token) {
    // Try to get from global function if available
    if (typeof getCurrentAccessToken === 'function') {
      token = await getCurrentAccessToken();
    }
  }
  
  if (!token) {
    alert("Google access token not available. Please sign in again.");
    return null;
  }
  
  console.log('Using token for API call:', token ? 'Token available' : 'No token');
  
  try {
    const res = await fetch(apiUrl, {
      headers: {
        'Authorization': `Bearer ${token}`
      }
    });
    
    if (!res.ok) {
      const err = await res.text();
      console.error('Sheets API error:', err);
      alert('Sheets API error: ' + err);
      return null;
    }
    
    const result = await res.json();
    return result;
  } catch (error) {
    console.error('Failed to fetch spreadsheet:', error);
    alert('Failed to fetch spreadsheet: ' + error.message);
    return null;
  }
}

function transformFetchedSpreadsheetToCreationFormat(fetchedSpreadsheet) {
  // Transforming fetched spreadsheet to creation format
  
  if (!fetchedSpreadsheet) {
    console.error("transformFetchedSpreadsheetToCreationFormat: Input spreadsheet is null or undefined");
    return null;
  }

  // Validate input structure
  if (!fetchedSpreadsheet.properties || !fetchedSpreadsheet.properties.title) {
    console.error("transformFetchedSpreadsheetToCreationFormat: Missing required properties");
    return null;
  }

  // If it's already in the correct format with all required fields, return it as is
  if (Array.isArray(fetchedSpreadsheet.sheets) && 
      fetchedSpreadsheet.sheets.length > 0 && 
      fetchedSpreadsheet.sheets[0].properties &&
      fetchedSpreadsheet.sheets[0].data &&
      fetchedSpreadsheet.sheets[0].properties.title &&
      fetchedSpreadsheet.sheets[0].properties.sheetId !== undefined) {
    // Spreadsheet already in correct format
    return fetchedSpreadsheet;
  }

  // Create a new object with the required structure
  const creationSpreadsheet = {
    properties: {
      title: fetchedSpreadsheet.properties.title
    },
    sheets: []
  };

  // If we have sheets data, transform it
  if (Array.isArray(fetchedSpreadsheet.sheets)) {
    creationSpreadsheet.sheets = fetchedSpreadsheet.sheets.map((sheet, index) => {
      // Ensure we have valid sheet properties
      const sheetProperties = sheet.properties || {};
      const transformedSheet = {
        properties: {
          title: sheetProperties.title || `Sheet${index + 1}`,
          sheetId: sheetProperties.sheetId || index,
          index: sheetProperties.index || index,
          sheetType: "GRID",
          gridProperties: {
            rowCount: 1000,
            columnCount: 26
          }
        },
        data: []
      };

      // If the sheet has data, include it
      if (Array.isArray(sheet.data)) {
        transformedSheet.data = sheet.data.map(gridData => ({
          rowData: Array.isArray(gridData.rowData) ? gridData.rowData : [],
          startRow: typeof gridData.startRow === 'number' ? gridData.startRow : 0,
          startColumn: typeof gridData.startColumn === 'number' ? gridData.startColumn : 0
        }));
      } else {
        // If no data, create an empty data array with proper structure
        transformedSheet.data = [{
          rowData: [],
          startRow: 0,
          startColumn: 0
        }];
      }

      return transformedSheet;
    });
  } else {
    // If no sheets data, create a default sheet with proper structure
    creationSpreadsheet.sheets = [{
      properties: {
        title: "Sheet1",
        sheetId: 0,
        index: 0,
        sheetType: "GRID",
        gridProperties: {
          rowCount: 1000,
          columnCount: 26
        }
      },
      data: [{
        rowData: [],
        startRow: 0,
        startColumn: 0
      }]
    }];
  }

  // Validate the transformed structure
  if (!Array.isArray(creationSpreadsheet.sheets) || creationSpreadsheet.sheets.length === 0) {
    console.error("transformFetchedSpreadsheetToCreationFormat: Failed to create valid sheet structure");
    return null;
  }

  // Validate each sheet
  for (const sheet of creationSpreadsheet.sheets) {
    if (!sheet.properties || !sheet.properties.title || !Array.isArray(sheet.data)) {
      console.error("transformFetchedSpreadsheetToCreationFormat: Invalid sheet structure in transformed spreadsheet");
      return null;
    }
  }

  // Transformation completed
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
async function batchUpdate(spreadsheetId, requests) {
  // Get the latest access token
  let token = null;
  if (typeof getCurrentAccessToken === 'function') {
    token = await getCurrentAccessToken();
  } else if (window.accessToken) {
    token = window.accessToken;
  } else if (localStorage.getItem('googleAccessToken')) {
    token = localStorage.getItem('googleAccessToken');
  }

  if (!token) {
    alert('Google access token not available. Please sign in again.');
    throw new Error('No access token available for Google Sheets API.');
  }

  const url = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}:batchUpdate`;
  try {
    const res = await fetch(url, {
      method: 'POST',
      headers: {
        'Authorization': `Bearer ${token}`,
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({requests})
    });
    if (!res.ok) {
      const err = await res.text();
      console.error('batchUpdate: Sheets API error:', err);
      throw new Error('Failed to batch update spreadsheet: ' + err);
    }
    // Optionally handle response
    // const result = await res.json();
    // return result;
  } catch (error) {
    console.error('batchUpdate: Error updating spreadsheet:', error);
    throw new Error('Failed to batch update spreadsheet: ' + (error.message || 'Unknown error'));
  }
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
  // Creating new document

  // Validate input structure
  if (!spreadsheet || !spreadsheet.properties || !spreadsheet.properties.title) {
    throw new Error('Invalid spreadsheet structure: missing required properties');
  }

  // Get the number of columns from the header row
  const headerRow = spreadsheet.sheets[0]?.data[0]?.rowData[0]?.values || [];
  const columnCount = headerRow.length;

  // Transform the spreadsheet structure to match Google Sheets API requirements
  const apiSpreadsheet = {
    properties: {
      title: spreadsheet.properties.title
    },
    sheets: spreadsheet.sheets.map(sheet => {
      // Ensure each sheet has the correct structure
      const apiSheet = {
        properties: {
          title: sheet.properties.title,
          sheetId: sheet.properties.sheetId || 0,
          index: sheet.properties.index || 0,
          sheetType: 'GRID',
          gridProperties: {
            rowCount: Math.max(1000, sheet.data?.[0]?.rowData?.length || 1000),
            columnCount: columnCount // Use the actual number of columns from header
          }
        }
      };
      // Add data if it exists, ensuring proper structure
      if (sheet.data && sheet.data.length > 0) {
        apiSheet.data = sheet.data.map(data => {
          // Ensure each row has exactly the right number of columns
          const rowData = data.rowData.map(row => {
            const values = row.values || [];
            // Pad or truncate to match header column count
            while (values.length < columnCount) {
              values.push({
                userEnteredValue: { stringValue: "" },
                userEnteredFormat: {
                  backgroundColor: { red: 1, green: 1, blue: 1 }
                }
              });
            }
            if (values.length > columnCount) {
              values.length = columnCount;
            }
            return { ...row, values };
          });
          return {
            rowData,
            startRow: data.startRow || 0,
            startColumn: data.startColumn || 0
          };
        });
      }
      return apiSheet;
    })
  };

  // --- TOKEN HANDLING ---
  let token = null;
  if (typeof getCurrentAccessToken === 'function') {
    token = await getCurrentAccessToken();
  } else if (window.accessToken) {
    token = window.accessToken;
  } else if (localStorage.getItem('googleAccessToken')) {
    token = localStorage.getItem('googleAccessToken');
  }

  if (!token) {
    alert('Google access token not available. Please sign in again.');
    throw new Error('No access token available for Google Sheets API.');
  }

  // --- API CALL ---
  try {
    // Use fetch to ensure Authorization header is set
    const res = await fetch('https://sheets.googleapis.com/v4/spreadsheets', {
      method: 'POST',
      headers: {
        'Authorization': `Bearer ${token}`,
        'Content-Type': 'application/json'
      },
      body: JSON.stringify(apiSpreadsheet)
    });

    if (!res.ok) {
      const err = await res.text();
      console.error('createNewDocument: Sheets API error:', err);
      throw new Error('Failed to create spreadsheet: ' + err);
    }
    const result = await res.json();
    if (result && result.spreadsheetUrl) {
      return result.spreadsheetUrl;
    } else {
      throw new Error('Failed to create spreadsheet: No spreadsheetUrl in response');
    }
  } catch (error) {
    console.error('createNewDocument: Error creating spreadsheet:', error);
    throw new Error('Failed to create spreadsheet: ' + (error.message || 'Unknown error'));
  }
}

function createSpreadSheet(title, headers) {
  // Creating spreadsheet with title
  var spreadSheet = {
    properties: {
      title: title
    },
    sheets: [{
      properties: {
        title: "Sheet1",
        sheetId: 0,
        index: 0,
        sheetType: "GRID",
        gridProperties: {
          rowCount: 1000,
          columnCount: 26
        }
      },
      data: [{
        rowData: [createRowData(headers)],
        startRow: 0,
        startColumn: 0
      }]
    }]
  };
  // Created spreadsheet structure
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
    // Row data logged
  }
}

function getManagersFromDataSpreadsheet(dataSpreadsheet) {
  if (!dataSpreadsheet) {
    console.error('getManagersFromDataSpreadsheet: dataSpreadsheet is null or undefined');
    alert('Failed to load spreadsheet data. Please check your access and try again.');
    return [];
  }
  
  if (!dataSpreadsheet.sheets || !Array.isArray(dataSpreadsheet.sheets)) {
    console.error('getManagersFromDataSpreadsheet: Invalid sheets structure:', dataSpreadsheet);
    alert('Invalid spreadsheet structure. Please check the spreadsheet format.');
    return [];
  }
  
  let managers = [];
  for(let i = 0; i < dataSpreadsheet.sheets.length; i++) {
    const sheet = dataSpreadsheet.sheets[i];
    if(sheet && sheet.properties && sheet.properties.title && sheet.properties.title !== "URL Data") {
      managers.push(sheet.properties.title);
    }
  }
  
  console.log('Found managers:', managers);
  return managers;
}

function getManagerSheet(dataSpreadsheet, manager) {
  if (!dataSpreadsheet) {
    console.error('getManagerSheet: dataSpreadsheet is null or undefined');
    alert('Manager sheet data is not available. Please try again.');
    return null;
  }
  
  if (!Array.isArray(dataSpreadsheet.sheets)) {
    console.error('getManagerSheet: Invalid sheets structure:', dataSpreadsheet);
    alert('Invalid spreadsheet structure. Please check the spreadsheet format.');
    return null;
  }
  
  const availableSheetNames = dataSpreadsheet.sheets.map(s => s.properties?.title || 'Unknown').filter(Boolean);
  // Looking for manager sheet
  let managerSheet = null;
  for(let i = 0; i < dataSpreadsheet.sheets.length; i++) {
    const sheet = dataSpreadsheet.sheets[i];
    if(sheet && sheet.properties && sheet.properties.title === manager) {
      managerSheet = sheet;
      break;
    }
  }
  
  if (!managerSheet) {
    console.warn('getManagerSheet: Manager sheet not found for manager:', manager);
    alert('Manager sheet not found: ' + manager + '. Available sheets: ' + availableSheetNames.join(', '));
    return null;
  }
  
  console.log('Found manager sheet:', managerSheet.properties.title);
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

// Add validation function at the top of the file
function validateSpreadsheetData(spreadsheet, spreadsheetName) {
  // Validating spreadsheet
  
  if (!spreadsheet || !spreadsheet.sheets) {
    return {
      isValid: false,
      errors: [`${spreadsheetName}: Invalid spreadsheet structure - missing sheets property`]
    };
  }

  const warnings = [];
  const errors = [];
  const problematicSheets = [];
  const problematicCells = [];

  for (let sheetIndex = 0; sheetIndex < spreadsheet.sheets.length; sheetIndex++) {
    const sheet = spreadsheet.sheets[sheetIndex];
    const sheetName = sheet.properties?.title || `Sheet${sheetIndex + 1}`;
    
    if (!sheet.data || !sheet.data[0] || !sheet.data[0].rowData) {
      errors.push(`${spreadsheetName} - ${sheetName}: Missing row data`);
      problematicSheets.push(sheetName);
      continue;
    }

    const rowData = sheet.data[0].rowData;
    
    for (let rowIndex = 0; rowIndex < rowData.length; rowIndex++) {
      const row = rowData[rowIndex];
      
      if (!row || !row.values) {
        warnings.push(`${spreadsheetName} - ${sheetName} Row ${rowIndex + 1}: Missing values array`);
        problematicCells.push(`${sheetName}!A${rowIndex + 1}`);
        continue;
      }

      for (let colIndex = 0; colIndex < row.values.length; colIndex++) {
        const cell = row.values[colIndex];
        
        // Check for formula cells
        if (cell && cell.userEnteredValue && cell.userEnteredValue.formulaValue) {
          warnings.push(`${spreadsheetName} - ${sheetName} ${indexToA1(rowIndex, colIndex)}: Contains formula "${cell.userEnteredValue.formulaValue}" - should use Paste Values Only`);
          problematicCells.push(`${sheetName}!${indexToA1(rowIndex, colIndex)}`);
        }
        
        // Check for cells with unexpected data types
        if (cell && cell.userEnteredValue) {
          const value = cell.userEnteredValue;
          if (value.numberValue !== undefined && value.stringValue !== undefined) {
            warnings.push(`${spreadsheetName} - ${sheetName} ${indexToA1(rowIndex, colIndex)}: Mixed data types detected`);
            problematicCells.push(`${sheetName}!${indexToA1(rowIndex, colIndex)}`);
          }
        }
        
        // Check for cells with special formatting that might cause issues
        if (cell && cell.userEnteredFormat && cell.userEnteredFormat.numberFormat) {
          warnings.push(`${spreadsheetName} - ${sheetName} ${indexToA1(rowIndex, colIndex)}: Has number formatting - may cause parsing issues`);
          problematicCells.push(`${sheetName}!${indexToA1(rowIndex, colIndex)}`);
        }
      }
    }
  }

  return {
    isValid: errors.length === 0,
    errors,
    warnings,
    problematicSheets,
    problematicCells,
    summary: {
      totalSheets: spreadsheet.sheets.length,
      problematicSheetsCount: problematicSheets.length,
      problematicCellsCount: problematicCells.length,
      errorCount: errors.length,
      warningCount: warnings.length
    }
  };
}

function displayValidationResults(validation, spreadsheetName) {
  const hasIssues = validation.errors.length > 0 || validation.warnings.length > 0;
  
  if (hasIssues) {
    console.warn(`=== VALIDATION RESULTS FOR ${spreadsheetName} ===`);
    console.warn(`Summary: ${validation.summary.totalSheets} sheets, ${validation.summary.problematicSheetsCount} problematic sheets, ${validation.summary.problematicCellsCount} problematic cells`);
    
    if (validation.errors.length > 0) {
      console.error('❌ ERRORS:');
      validation.errors.forEach(error => console.error(`  - ${error}`));
    }
    
    if (validation.warnings.length > 0) {
      console.warn('⚠️  WARNINGS:');
      validation.warnings.forEach(warning => console.warn(`  - ${warning}`));
    }
    
    if (validation.problematicSheets.length > 0) {
      console.warn('📋 Problematic Sheets:');
      validation.problematicSheets.forEach(sheet => console.warn(`  - ${sheet}`));
    }
    
    if (validation.problematicCells.length > 0) {
      console.warn('📍 Problematic Cells (first 20):');
      validation.problematicCells.slice(0, 20).forEach(cell => console.warn(`  - ${cell}`));
      if (validation.problematicCells.length > 20) {
        console.warn(`  ... and ${validation.problematicCells.length - 20} more cells`);
      }
    }
    
    // Show user-friendly alert with guidance
    const guidance = getDataIssueGuidance(validation);
    const message = `Data validation issues found in ${spreadsheetName}:\n\n` +
      `• ${validation.summary.errorCount} errors\n` +
      `• ${validation.summary.warningCount} warnings\n` +
      `• ${validation.summary.problematicSheetsCount} problematic sheets\n` +
      `• ${validation.summary.problematicCellsCount} problematic cells\n\n` +
      `${guidance}\n\n` +
      `Check the console for detailed cell locations.`;
    
    alert(message);
  } else {
    // All data validated successfully
  }
  
  return hasIssues;
}

function getDataIssueGuidance(validation) {
  const guidance = [];
  
  guidance.push('⚠️ RECOMMENDATIONS:');
  guidance.push('1. Select all cells in your source spreadsheet');
  guidance.push('2. Copy (Cmd+C)');
  guidance.push('3. In the target spreadsheet, use "Paste Values Only" (Cmd+Shift+V)');
  
  return guidance.join('\n');
}
