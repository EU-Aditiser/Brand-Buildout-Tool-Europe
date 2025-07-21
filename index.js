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

// Unified authentication management
function updateAuthState(token) {
  accessToken = token;
  window.accessToken = token;
  
  const signInBtn = document.getElementById('g_id_signin');
  if (!signInBtn) {
    console.error('Sign-in button container not found');
    return;
  }
  
  if (token) {
    localStorage.setItem('accessToken', token);
    // Update button to show "You're signed in" with sign-out option
    signInBtn.innerHTML = `
      <div class="d-flex gap-2 align-items-center">
        <button class="btn btn-success" disabled>✓ You're signed in</button>
        <button id="sign-out-btn" class="btn btn-outline-secondary btn-sm">Sign Out</button>
      </div>
    `;
    
    // Add sign-out functionality
    const signOutBtn = document.getElementById('sign-out-btn');
    if (signOutBtn) {
      signOutBtn.onclick = function() {
        signOut();
      };
    }
    
    // Authentication successful
  } else {
    localStorage.removeItem('accessToken');
    // Update button to show "Sign in with Google"
    signInBtn.innerHTML = '<button id="custom-google-signin" class="btn btn-outline-primary">Sign in with Google</button>';
    const newButton = document.getElementById('custom-google-signin');
    if (newButton) {
      newButton.onclick = function() {
        gisLoaded();
      };
    }
    
    // Authentication cleared
  }
}

// Add sign-out function
function signOut() {
  // Clear all authentication data
  accessToken = null;
  window.accessToken = null;
  localStorage.removeItem('accessToken');
  
  // Update UI to show sign-in button
  updateAuthState(null);
  
  // Reset any error states
  const errorDiv = document.getElementById('signin-error');
  if (errorDiv) {
    errorDiv.style.display = 'none';
  }
  
  // Successfully signed out
}

// Check if user is authenticated
function isAuthenticated() {
  const token = getCurrentAccessToken();
  return !!token;
}

// Get the current access token
function getCurrentAccessToken() {
  // First check memory, then localStorage
  if (accessToken) {
    return accessToken;
  }
  
  const storedToken = localStorage.getItem('accessToken');
  if (storedToken) {
    // Restore token to memory
    accessToken = storedToken;
    window.accessToken = storedToken;
    return storedToken;
  }
  
  return null;
}

// Stub for isTokenExpired to prevent ReferenceError
function isTokenExpired() {
  // TODO: Implement real token expiration check if needed
  return false;
}

/**
 * 
 * Create Brand Template on Click 
 */
async function handleCreateBrandTemplateBuildoutClick(e) {
  try {
    // Check authentication first
    if (!isAuthenticated()) {
      alert('You must sign in with Google before proceeding.');
      return;
    }
    
    // Check if Google API client is initialized
    if (!gapi.client.sheets) {
      await initGoogleApiClient();
    }
    
    // Ensure token is available
    const token = getCurrentAccessToken();
    if (!token) {
      alert('Authentication token not found. Please sign in again.');
      signOut();
      return;
    }
    
    // Update window.accessToken
    window.accessToken = token;
    
    const formData = readBrandBuildoutTemplateData();

    const spreadsheet = createBrandBuildoutTemplateSpreadsheet(formData.campaign, formData.adGroup, formData.baseKeyword, formData.finalUrl);

    //create spreadsheet
    const url = await createNewDocument(spreadsheet);
    if (url) {
      window.open(url, '_blank');
    }
  } catch (error) {
    console.error("Error in brand template creation:", error);
    
    if (error.message && (error.message.includes('access token') || error.message.includes('unauthorized'))) {
      alert('Authentication error. Please sign out and sign in again.');
      signOut();
    } else {
      alert('An error occurred. Please try again.');
    }
  }
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
  if (!managerSelect) {
    alert("Manager select element not found. Please refresh the page and try again.");
    return "";
  }
  const manager = managerSelect.options[managerSelect.selectedIndex].value;
  return manager;
}

/**
 * 
 * Create Account buildouts on Click 
 */
async function handleAccountBuildoutClick(e) {
  // Ensure authentication is still valid before proceeding
  if (!isAuthenticated()) {
    // Authentication lost during account buildout, signing out
    signOut();
    return;
  }
  
  // Ensure token is available for API calls
  const token = getCurrentAccessToken();
  if (token) {
    window.accessToken = token;
  }
  
  switch(BUTTON_STATE) {
    case "GET_MANAGER_DATA":
      updateButtonState("")

      const managerFormData = readAccountBuildoutData();
      const managerDataSpreadsheet = await fetchSpreadsheetNoGridData(managerFormData.accountDataSpreadsheetURL)
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
      if (!manager) {
        alert("Please select a manager before proceeding.");
        updateButtonState(BUTTON_STATE);
        break;
      }
      MANAGER = manager;
      const dataSpreadsheet = await fetchSpreadsheetSingleManager(formData.accountDataSpreadsheetURL, MANAGER)
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
      const brandBuildoutSpreadsheet = await fetchSpreadsheet(formDataCreate.brandBuildoutSpreadsheetURL)
      const accountDataSpreadsheetCreate = await fetchSpreadsheetSingleManager(formDataCreate.accountDataSpreadsheetURL, MANAGER)
      
      // Check if spreadsheets were fetched successfully
      if (!brandBuildoutSpreadsheet) {
        console.error("Failed to fetch brand buildout spreadsheet");
        alert("Failed to fetch brand buildout spreadsheet. Please check the URL and try again.");
        updateButtonState(BUTTON_STATE);
        return;
      }
      
      if (!accountDataSpreadsheetCreate) {
        console.error("Failed to fetch account data spreadsheet");
        alert("Failed to fetch account data spreadsheet. Please check the URL and try again.");
        updateButtonState(BUTTON_STATE);
        return;
      }
      
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

// Initialize Google API client
function initGoogleApiClient() {
  return new Promise((resolve, reject) => {
    gapi.load('client', async () => {
      try {
        await gapi.client.init({
          apiKey: API_KEY,
          discoveryDocs: ['https://sheets.googleapis.com/$discovery/rest?version=v4'],
        });
        // Google API client initialized successfully
        resolve();
      } catch (error) {
        console.error('Error initializing Google API client:', error);
        reject(error);
      }
    });
  });
}

// Implement GIS OAuth 2.0 code flow
async function gisLoaded() {
  try {
    // Initialize the Google API client first
    await initGoogleApiClient();
    
    // Then initialize the OAuth client
    google.accounts.oauth2.initTokenClient({
      client_id: CLIENT_ID,
      scope: SCOPES,
      callback: async (tokenResponse) => {
        if (tokenResponse && tokenResponse.access_token) {
          updateAuthState(tokenResponse.access_token);
        } else {
          console.error('Failed to get access token:', tokenResponse);
          alert('Failed to get access token. Please try again.');
        }
      }
    }).requestAccessToken();
  } catch (error) {
    console.error('Error in gisLoaded:', error);
    alert('Failed to initialize Google API client. Please refresh the page and try again.');
  }
}

window.onload = function() {
  // Always ensure the sign-in container exists and is visible
  const signInBtn = document.getElementById('g_id_signin');
  if (!signInBtn) {
    console.error('Sign-in container not found');
    return;
  }
  
  // Check if user is already logged in
  const existingToken = getCurrentAccessToken();
  
  if (existingToken) {
    // User is already logged in, update the auth state
    updateAuthState(existingToken);
  } else {
    // User is not logged in, show the sign-in button
    signInBtn.innerHTML = '<button id="custom-google-signin" class="btn btn-outline-primary">Sign in with Google</button>';
    const newButton = document.getElementById('custom-google-signin');
    if (newButton) {
      newButton.onclick = function() {
        gisLoaded();
      };
    }
  }

  // Dynamically load patch notes from patchnotes.txt
  fetch('patchnotes.txt')
    .then(response => response.text())
    .then(text => {
      // Replace newline characters with HTML break tags for proper formatting
      const formattedText = text.replace(/\n/g, '<br>');
      document.getElementById('patch_notes').innerHTML = formattedText;
    })
    .catch(err => {
      document.getElementById('patch_notes').textContent = 'Unable to load patch notes.';
      console.error('Failed to load patchnotes.txt:', err);
    });
};

// Update the account buildout button click handler
accountBuildoutButton.onclick = async function() {
  try {
    // Check authentication first
    if (!isAuthenticated()) {
      alert('You must sign in with Google before proceeding.');
      return;
    }
    
    // Check if Google API client is initialized
    if (!gapi.client.sheets) {
      await initGoogleApiClient();
    }
    
    // Ensure token is available in memory and synchronized
    const token = getCurrentAccessToken();
    if (!token) {
      alert('Authentication token not found. Please sign in again.');
      signOut();
      return;
    }
    
    // Update window.accessToken to ensure it's available for API calls
    window.accessToken = token;
    
    // Starting account buildout with authenticated token
    
    // Proceed with account buildout
    await handleAccountBuildoutClick();
  } catch (error) {
    console.error("Error in account buildout button click:", error);
    
    // Check if it's an authentication error
    if (error.message && (error.message.includes('access token') || error.message.includes('unauthorized'))) {
      alert('Authentication error. Please sign out and sign in again.');
      signOut();
    } else if (error.message && error.message.includes('Failed to fetch')) {
      alert('Network error. Please check your internet connection and try again.');
    } else {
      alert('An error occurred: ' + (error.message || 'Unknown error'));
    }
    
    // Don't reset UI state on error - let user try again with same data
    // resetUIState(); // Reset UI so user can try again
  }
};

// Add a global function to reset UI state after spreadsheet creation
function resetUIState() {
  // Don't clear input fields - keep them for repeated use
  // document.getElementById("master_brand_buildout_spreadsheet").value = "";
  // document.getElementById("account_data_spreadsheet").value = "";
  
  // Only clear the accounts form and reset button state
  document.getElementById("accounts_form").innerHTML = "";
  BUTTON_STATE = "GET_MANAGER_DATA";
  updateButtonState(BUTTON_STATE);
  ACCOUNTS = [];
  
  // Ensure authentication is still valid
  if (!isAuthenticated()) {
    // Authentication lost during reset, signing out
    signOut();
  }
};
