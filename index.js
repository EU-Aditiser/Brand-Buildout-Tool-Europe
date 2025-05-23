/** 
 * Initialize Buttons 
 * */
const signoutButton = document.getElementById('signout_button');
const createBrandBuildoutTemplateButton = document.getElementById('create_brand_buildout_template_button');
const accountBuildoutButton = document.getElementById('create_account_buildout_button');

let BUTTON_STATE = "GET_MANAGER_DATA";
let MANAGER = "";
updateButtonState(BUTTON_STATE);

let ACCOUNTS = [];
let accessToken = null;
let tokenClient;


/**
 * 
 * Create Brand Template on Click 
 */
function handleCreateBrandTemplateBuildoutClick(e) {
  const formData = readBrandBuildoutTemplateData();

  const spreadsheet = createBrandBuildoutTemplateSpreadsheet(formData.campaign, formData.adGroup, formData.baseKeyword, formData.finalUrl);

  //create spreadsheet
  const url = createNewDocument(spreadsheet);
  window.open(url, '_blank');
}

function updateButtonState(state) {
  let button_html = ""
  document.getElementById("buildout_button").innerHTML = button_html;

  switch(state) {
    case "GET_MANAGER_DATA":
      button_html = `Get Manager Data`
      break;
    case "GET_DATA":
      button_html = `Get Account Data`
      break;
    case "SELECT_DATA":
      button_html = `Please Select At Least One Account`
      break;
    case "CREATE":
      button_html = `Create`
      break;
    default:
      button_html = `
      <span class="spinner-grow spinner-grow-sm" id="create_account_buildout_loader"></span>
          Working...
      `
  }

  document.getElementById("buildout_button").innerHTML = button_html;
}

function readManagerSelectData() {
  const managerSelect = document.getElementById("manager-select");
  const manager = managerSelect.options[managerSelect.selectedIndex].value;
  return manager;
}

/**
 * 
 * Create Account buildouts on Click 
 */
async function handleAccountBuildoutClick(e) {
  // Always set the access token before any Sheets API call
  if (accessToken) {
    gapi.client.setToken({ access_token: accessToken });
  }
  switch(BUTTON_STATE) {
    case "GET_MANAGER_DATA":
      updateButtonState("")

      const managerFormData = readAccountBuildoutData();
      const managerDataSpreadsheet = await getSpreadsheetNoGridData(managerFormData.accountDataSpreadsheetURL)
      const managers = getManagersFromDataSpreadsheet(managerDataSpreadsheet);
      
      const managerHtml = createManagerHtml(managers)
      document.getElementById("accounts_form").innerHTML = managerHtml;
     
      BUTTON_STATE = "GET_DATA";
      updateButtonState(BUTTON_STATE)

      break;
    case "GET_DATA":
      updateButtonState("");
      const formData = readAccountBuildoutData();

      const manager = readManagerSelectData();
      MANAGER = manager;
      const dataSpreadsheet = await getSpreadsheetSingleManager(formData.accountDataSpreadsheetURL, MANAGER)
      const managerSheet = getManagerSheet(dataSpreadsheet, manager)
      ACCOUNTS = getAccountsFromManagerSheet(managerSheet)
      
      const accounts = ACCOUNTS;

      const accountHtml = createAccountHtml(accounts)
      document.getElementById("accounts_form").innerHTML = accountHtml;
      BUTTON_STATE = "CREATE";
      updateButtonState(BUTTON_STATE)
      break;
    
    case "CREATE":
      updateButtonState("")
      const selectedAccounts = readAccountCheckBoxData();
      
      if(selectedAccounts.length === 0) {
        alert("Please select at least one account.")
        updateButtonState(BUTTON_STATE);
        break;
      }
      
      const formDataCreate = readAccountBuildoutData();
      const brandBuildoutSpreadsheet = await getSpreadsheet(formDataCreate.brandBuildoutSpreadsheetURL)
      const accountDataSpreadsheetCreate = await getSpreadsheetSingleManager(formDataCreate.accountDataSpreadsheetURL, MANAGER)
      
      try {
        const newSheetUrl = await processRequest(brandBuildoutSpreadsheet, accountDataSpreadsheetCreate, selectedAccounts);
        if (newSheetUrl) {
          window.open(newSheetUrl, '_blank');
        }
      } catch (e) {
        console.error("Error in processRequest:", e);
        alert("An error occurred: " + (e.message || e));
        // Optionally: signoutButton.onclick();
      }

      BUTTON_STATE = "GET_DATA";
      updateButtonState(BUTTON_STATE);
      document.getElementById("accounts_form").innerHTML = "";
      ACCOUNTS = [];
    break;

    default:
      console.log("")
  }
}

function readAccountCheckBoxData() {
  let accounts = [];

  for(let i = 0; i < ACCOUNTS.length; i++) {
    const title = ACCOUNTS[i].accountTitle || ACCOUNTS[i]
    const boxdiv = document.getElementById(title.split (" ").join("")+"-checkbox");
    const checked = boxdiv.checked;

    if(checked) {
      accounts.push(ACCOUNTS[i]);
    }
  }

  return accounts;
}

function readBrandBuildoutTemplateData() {
  let formData = {};

  formData.campaign = document.getElementById("campaign_name").value;
  formData.adGroup = document.getElementById("ad_group_name").value;
  formData.baseKeyword = document.getElementById("base_keyword_name").value;
  formData.finalUrl = document.getElementById("final_url_name").value;

  return formData;
}

function readAccountBuildoutData() {
  let formData = {};

  formData.brandBuildoutSpreadsheetURL = document.getElementById("master_brand_buildout_spreadsheet").value;
  formData.accountDataSpreadsheetURL = document.getElementById("account_data_spreadsheet").value;
  
  if(formData.brandBuildoutSpreadsheetURL === "" || formData.accountDataSpreadsheetURL === "") {
    alert("Please enter both a keyword spreadsheet and a manager-style data spreadsheet.")
    location.reload()
  }

  return formData;
}

function readTextFile(file)
{
    var rawFile = new XMLHttpRequest();
    rawFile.open("GET", file, false);
    rawFile.onreadystatechange = function ()
    {
        if(rawFile.readyState === 4)
        {
            if(rawFile.status === 200 || rawFile.status == 0)
            {
                var allText = rawFile.responseText;
            }
        }
    }
    rawFile.send(null);
    return rawFile.responseText;
}

// GIS Authentication Migration
function handleCredentialResponse(response) {
  // Only handle UI, do not set accessToken here
  document.getElementById('g_id_signin').style.display = 'none';
  signoutButton.style.display = 'block';
}

function initGapiClient() {
  gapi.client.init({
    apiKey: API_KEY,
    discoveryDocs: DISCOVERY_DOCS,
  }).then(function () {
    gapi.client.setToken({ access_token: accessToken });
  }, function(error) {
    console.error("Google API client init error:", error);
  });
}

window.onload = function() {
  google.accounts.id.initialize({
    client_id: CLIENT_ID,
    callback: handleCredentialResponse,
    ux_mode: 'popup',
    auto_select: false
  });
  google.accounts.id.renderButton(
    document.getElementById('g_id_signin'),
    { theme: 'outline', size: 'large' }
  );
  google.accounts.id.prompt();
  signoutButton.style.display = 'none';

  // Initialize the token client for OAuth2
  tokenClient = google.accounts.oauth2.initTokenClient({
    client_id: CLIENT_ID,
    scope: 'https://www.googleapis.com/auth/spreadsheets https://www.googleapis.com/auth/drive',
    callback: (tokenResponse) => {
      accessToken = tokenResponse.access_token;
      gapi.client.setToken({ access_token: accessToken });
      gapi.load('client', initGapiClient);
    },
  });
};

signoutButton.onclick = function() {
  google.accounts.id.disableAutoSelect();
  accessToken = null;
  document.getElementById('g_id_signin').style.display = 'block';
  signoutButton.style.display = 'none';
};

// Call this function when you want to request Sheets/Drive access
function requestSheetsAccess() {
  tokenClient.requestAccessToken();
}

accountBuildoutButton.onclick = async function() {
  if (!accessToken) {
    requestSheetsAccess();
    return;
  }
  // Ensure gapi.client is loaded and initialized before setting the token
  if (!gapi.client || !gapi.client.setToken) {
    gapi.load('client', () => {
      gapi.client.init({
        apiKey: API_KEY,
        discoveryDocs: DISCOVERY_DOCS,
      }).then(function () {
        gapi.client.setToken({ access_token: accessToken });
        handleAccountBuildoutClick();
      }, function(error) {
        console.error("Google API client init error:", error);
      });
    });
    return;
  }
  gapi.client.setToken({ access_token: accessToken });
  try {
    await handleAccountBuildoutClick();
  } catch (e) {
    if (e.status === 403 || (e.result && e.result.error && e.result.error.code === 403)) {
      // If a 403 (PERMISSION_DENIED) is returned, re-request the token and re-try (after a small delay)
      requestSheetsAccess();
      setTimeout(() => { if (accessToken) handleAccountBuildoutClick(); }, 500);
    } else {
      console.error("Error in handleAccountBuildoutClick:", e);
    }
  }
};

// Add a global function to reset UI state after spreadsheet creation
function resetUIState() {
  // Clear input fields
  document.getElementById("master_brand_buildout_spreadsheet").value = "";
  document.getElementById("account_data_spreadsheet").value = "";
  document.getElementById("accounts_form").innerHTML = "";
  BUTTON_STATE = "GET_MANAGER_DATA";
  updateButtonState(BUTTON_STATE);
  ACCOUNTS = [];
}