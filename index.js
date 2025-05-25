/** 
 * Initialize Buttons 
 * */
const createBrandBuildoutTemplateButton = document.getElementById('create_brand_buildout_template_button');
const accountBuildoutButton = document.getElementById('create_account_buildout_button');

let BUTTON_STATE = "GET_MANAGER_DATA";
let MANAGER = "";
updateButtonState(BUTTON_STATE);

let ACCOUNTS = [];
let accessToken = null;
let tokenClient;

// Stub for isTokenExpired to prevent ReferenceError
function isTokenExpired() {
  // TODO: Implement real token expiration check if needed
  return false;
}

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
  if (response.credential) {
    // Hide the sign-in button
    document.getElementById('g_id_signin').style.display = 'none';
    
    // Request access token after successful sign-in
    requestSheetsAccess();
  } else {
    console.error("Sign-in failed:", response);
    alert('Sign-in failed. Please try again.');
  }
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
  // Initialize Google Sign-In with updated configuration
  google.accounts.id.initialize({
    client_id: CLIENT_ID,
    callback: handleCredentialResponse,
    ux_mode: 'popup', // Switch from 'redirect' to 'popup'
    auto_select: false,
    cancel_on_tap_outside: false,
    context: 'signin', // Add context
    prompt_parent_id: 'g_id_signin', // Specify parent element
    use_fedcm_for_prompt: true // Use FedCM for better browser compatibility
  });

  // Render the sign-in button with updated configuration
  google.accounts.id.renderButton(
    document.getElementById('g_id_signin'),
    { 
      theme: 'outline', 
      size: 'large',
      type: 'standard',
      text: 'signin_with',
      shape: 'rectangular',
      width: 250, // Specify width
      logo_alignment: 'left'
    }
  );

  // Initialize the token client for OAuth2 with updated configuration
  tokenClient = google.accounts.oauth2.initTokenClient({
    client_id: CLIENT_ID,
    scope: 'https://www.googleapis.com/auth/spreadsheets https://www.googleapis.com/auth/drive',
    callback: async (tokenResponse) => {
      if (tokenResponse && tokenResponse.access_token) {
        accessToken = tokenResponse.access_token;
        try {
          // Initialize GAPI client first
          await new Promise((resolve, reject) => {
            gapi.load('client', () => {
              gapi.client.init({
                apiKey: API_KEY,
                discoveryDocs: DISCOVERY_DOCS,
              }).then(() => {
                gapi.client.setToken({ access_token: accessToken });
                resolve();
              }).catch(reject);
            });
          });
        } catch (error) {
          console.error("Error initializing GAPI client:", error);
          // Show error to user
          alert('Error initializing Google services. Please try refreshing the page.');
        }
      }
    },
    error_callback: (error) => {
      console.error("Token client error:", error);
      alert('Authentication error. Please try signing in again.');
    }
  });

  // Load the GAPI client
  gapi.load('client', initGapiClient);
};

// Update the token refresh function
async function refreshToken() {
  return new Promise((resolve, reject) => {
    try {
      tokenClient.requestAccessToken({
        prompt: 'consent',
        hint: 'select_account' // Add hint for better UX
      });
      
      // Set up a listener for the token response
      const tokenListener = (tokenResponse) => {
        if (tokenResponse && tokenResponse.access_token) {
          accessToken = tokenResponse.access_token;
          gapi.client.setToken({ access_token: accessToken });
          resolve();
        } else {
          reject(new Error('Failed to refresh token'));
        }
      };
      
      // Add the listener
      tokenClient.callback = tokenListener;
    } catch (error) {
      console.error("Error refreshing token:", error);
      reject(error);
    }
  });
}

// Update the request sheets access function
function requestSheetsAccess() {
  try {
    tokenClient.requestAccessToken({
      prompt: 'consent',
      hint: 'select_account'
    });
  } catch (error) {
    console.error("Error requesting sheets access:", error);
    alert('Error accessing Google Sheets. Please try signing in again.');
  }
}

// Update the account buildout button click handler
accountBuildoutButton.onclick = async function() {
  try {
    // Check authentication state
    if (!accessToken || isTokenExpired()) {
      await refreshToken();
    }
    
    // Ensure GAPI client is initialized
    if (!gapi.client || !gapi.client.setToken) {
      await new Promise((resolve, reject) => {
        gapi.load('client', () => {
          gapi.client.init({
            apiKey: API_KEY,
            discoveryDocs: DISCOVERY_DOCS,
          }).then(() => {
            if (accessToken) {
              gapi.client.setToken({ access_token: accessToken });
            }
            resolve();
          }).catch(reject);
        });
      });
    }

    // Proceed with account buildout
    await handleAccountBuildoutClick();
  } catch (error) {
    console.error("Error in account buildout button click:", error);
    if (error.status === 401 || error.status === 403 || 
        (error.result && error.result.error && 
         (error.result.error.code === 401 || error.result.error.code === 403))) {
      try {
        await refreshToken();
        return accountBuildoutButton.onclick();
      } catch (refreshError) {
        console.error("Error refreshing token:", refreshError);
        alert('Authentication error. Please sign in again.');
      }
    } else {
      alert('An error occurred. Please try again or refresh the page.');
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