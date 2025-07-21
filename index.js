/** 
 * Initialize Buttons 
 * */
const createBrandBuildoutTemplateButton = document.getElementById('create_brand_buildout_template_button');
const accountBuildoutButton = document.getElementById('create_account_buildout_button');

// Global variables
let BUTTON_STATE = "GET_MANAGER_DATA";
let ACCOUNTS = [];
let MANAGER = "";
let accessToken = null;
let gapiInitialized = false;

// DOM elements
const signInBtn = document.getElementById('g_id_signin');
const buildoutButtonSpan = document.getElementById('buildout_button');

// Initialize the application
document.addEventListener('DOMContentLoaded', function() {
  console.log('=== APPLICATION STARTING ===');
  
  // Initialize Google API client
  initGoogleApiClient();
  
  // Check for existing authentication
  checkExistingAuth();
  
  // Set up event listeners
  setupEventListeners();
  
  // Load patch notes
  loadPatchNotes();
});

// Initialize Google API Client
async function initGoogleApiClient() {
  try {
    console.log('Initializing Google API Client...');
    
    await new Promise((resolve, reject) => {
      gapi.load('client', {
        callback: resolve,
        onerror: reject
      });
    });
    
    await gapi.client.init({
      apiKey: API_KEY,
      discoveryDocs: DISCOVERY_DOCS
    });
    
    gapiInitialized = true;
    console.log('Google API Client initialized successfully');
  } catch (error) {
    console.error('Failed to initialize Google API Client:', error);
  }
}

// Check for existing authentication
function checkExistingAuth() {
  console.log('Checking for existing authentication...');
  
  // Check localStorage for token
  const storedToken = localStorage.getItem('googleAccessToken');
  const tokenExpiry = localStorage.getItem('googleTokenExpiry');
  
  if (storedToken && tokenExpiry) {
    const now = Date.now();
    const expiry = parseInt(tokenExpiry);
    
    if (now < expiry) {
      // Token is still valid
      accessToken = storedToken;
      updateAuthUI(true);
      console.log('Found valid stored token');
      return;
    } else {
      // Token expired, clear it
      console.log('Stored token expired, clearing...');
      clearStoredAuth();
    }
  }
  
  // No valid token found
  updateAuthUI(false);
  console.log('No valid authentication found');
}

// Update authentication UI
function updateAuthUI(isAuthenticated) {
  const accountBuildoutButton = document.getElementById('create_account_buildout_button');
  
  if (isAuthenticated) {
    signInBtn.innerHTML = `
      <button class="btn btn-primary btn-sm" onclick="signOut()">Sign Out</button>
    `;
    buildoutButtonSpan.textContent = 'Get Manager Data';
    if (accountBuildoutButton) {
      accountBuildoutButton.disabled = false;
      accountBuildoutButton.className = 'btn btn-primary mt-3';
    }
  } else {
    signInBtn.innerHTML = `
      <button class="btn btn-primary" onclick="signIn()">
        Authorize
      </button>
    `;
    buildoutButtonSpan.textContent = 'Get Manager Data';
    if (accountBuildoutButton) {
      accountBuildoutButton.disabled = true;
      accountBuildoutButton.className = 'btn btn-primary mt-3';
    }
  }
}

// Sign in function
async function signIn() {
  try {
    console.log('Starting sign in process...');
    
    // Initialize Google Identity Services
    if (typeof google === 'undefined' || !google.accounts) {
      console.error('Google Identity Services not loaded');
      alert('Google Identity Services not available. Please refresh the page and try again.');
      return;
    }
    
    // Create token client
    const tokenClient = google.accounts.oauth2.initTokenClient({
      client_id: CLIENT_ID,
      scope: SCOPES,
      callback: handleTokenResponse
    });
    
    // Request token
    tokenClient.requestAccessToken();
    
  } catch (error) {
    console.error('Sign in error:', error);
    alert('Sign in failed: ' + error.message);
  }
}

// Handle token response
function handleTokenResponse(response) {
  console.log('Token response received:', response);
  
  if (response.error) {
    console.error('Token error:', response.error);
    alert('Authentication failed: ' + response.error);
    return;
  }
  
  // Store token and expiry
  accessToken = response.access_token;
  const expiry = Date.now() + (response.expires_in * 1000);
  
  localStorage.setItem('googleAccessToken', accessToken);
  localStorage.setItem('googleTokenExpiry', expiry.toString());
  
  // Update UI
  updateAuthUI(true);
  
  console.log('Authentication successful');
}

// Sign out function
function signOut() {
  console.log('Signing out...');
  
  // Clear stored data
  clearStoredAuth();
  
  // Update UI
  updateAuthUI(false);
  
  // Reset application state
  BUTTON_STATE = "GET_MANAGER_DATA";
  ACCOUNTS = [];
  MANAGER = "";
  
  // Clear forms
  document.getElementById("accounts_form").innerHTML = "";
  updateButtonState(BUTTON_STATE);
  
  console.log('Sign out complete');
}

// Clear stored authentication
function clearStoredAuth() {
  localStorage.removeItem('googleAccessToken');
  localStorage.removeItem('googleTokenExpiry');
  accessToken = null;
}

// Check if user is authenticated
function isAuthenticated() {
  return accessToken !== null;
}

// Get current access token, refresh if expired
async function getCurrentAccessToken() {
  if (!accessToken || isTokenExpired()) {
    // Try to refresh the token
    ensureTokenClient();
    if (tokenClient) {
      console.log('Access token expired or missing, requesting new token...');
      // Wrap in a promise to wait for callback
      await new Promise((resolve, reject) => {
        tokenClient.callback = (response) => {
          if (response && response.access_token) {
            accessToken = response.access_token;
            const expiry = Date.now() + (response.expires_in * 1000);
            localStorage.setItem('googleAccessToken', accessToken);
            localStorage.setItem('googleTokenExpiry', expiry.toString());
            updateAuthUI(true);
            resolve();
          } else {
            alert('Session expired. Please sign in again.');
            signOut();
            reject('No access token received');
          }
        };
        tokenClient.requestAccessToken();
      });
    } else {
      alert('Session expired. Please sign in again.');
      signOut();
      return null;
    }
  }
  return accessToken;
}

function isTokenExpired() {
  const expiry = localStorage.getItem('googleTokenExpiry');
  if (!expiry) return true;
  return Date.now() > parseInt(expiry);
}

// GIS token client for refresh
let tokenClient = null;

function ensureTokenClient() {
  if (!tokenClient && typeof google !== 'undefined' && google.accounts) {
    tokenClient = google.accounts.oauth2.initTokenClient({
      client_id: CLIENT_ID,
      scope: SCOPES,
      callback: handleTokenResponse
    });
  }
}

// Set up event listeners
function setupEventListeners() {
  // Account buildout button
  const accountBuildoutButton = document.getElementById('create_account_buildout_button');
  if (accountBuildoutButton) {
    accountBuildoutButton.addEventListener('click', handleAccountBuildoutClick);
  }
  
  // Create brand template button
  const createButton = document.getElementById('create_brand_buildout_template_button');
  if (createButton) {
    createButton.addEventListener('click', handleCreateBrandTemplateBuildoutClick);
  }
}

// Handle account buildout button click
async function handleAccountBuildoutClick() {
  try {
    console.log('=== ACCOUNT BUILDOUT BUTTON CLICKED ===');
    console.log('Current state:', BUTTON_STATE);
    console.log('Authenticated:', isAuthenticated());
    
    // Check authentication
    if (!isAuthenticated()) {
      alert('Please sign in with Google to continue.');
      return;
    }
    
    // Ensure Google API is initialized
    if (!gapiInitialized) {
      console.log('Initializing Google API...');
      await initGoogleApiClient();
    }
    
    // Process based on current state
    switch (BUTTON_STATE) {
      case "GET_MANAGER_DATA":
        await handleGetManagerData();
        break;
      case "GET_DATA":
        await handleGetData();
        break;
      case "CREATE":
        await handleCreate();
        break;
      default:
        console.error('Unknown button state:', BUTTON_STATE);
    }
    
  } catch (error) {
    console.error('Account buildout error:', error);
    alert('An error occurred: ' + error.message);
  }
}

// Handle get manager data
async function handleGetManagerData() {
  console.log('Getting manager data...');
  
  updateButtonState("");
  
  const formData = readAccountBuildoutData();
  console.log('Form data:', formData);
  
  const spreadsheet = await fetchSpreadsheetNoGridData(formData.accountDataSpreadsheetURL);
  console.log('Spreadsheet response:', spreadsheet);
  
  if (!spreadsheet) {
    alert('Failed to fetch spreadsheet. Please check the URL and try again.');
    updateButtonState(BUTTON_STATE);
    return;
  }
  
  const managers = getManagersFromDataSpreadsheet(spreadsheet);
  console.log('Managers found:', managers);
  
  if (!managers || managers.length === 0) {
    alert('No managers found in the spreadsheet.');
    updateButtonState(BUTTON_STATE);
    return;
  }
  
  const managerHtml = createManagerHtml(managers);
  document.getElementById("accounts_form").innerHTML = managerHtml;
  
  BUTTON_STATE = "GET_DATA";
  updateButtonState(BUTTON_STATE);
}

// Handle get data
async function handleGetData() {
  console.log('Getting data...');
  
  updateButtonState("");
  
  const formData = readAccountBuildoutData();
  const manager = readManagerSelectData();
  
  if (!manager) {
    alert('Please select a manager.');
    updateButtonState(BUTTON_STATE);
    return;
  }
  
  MANAGER = manager;
  console.log('Selected manager:', MANAGER);
  
  const spreadsheet = await fetchSpreadsheetSingleManager(formData.accountDataSpreadsheetURL, MANAGER);
  console.log('Manager spreadsheet:', spreadsheet);
  
  if (!spreadsheet) {
    alert('Failed to fetch manager data.');
    updateButtonState(BUTTON_STATE);
    return;
  }
  
  const managerSheet = getManagerSheet(spreadsheet, manager);
  
  if (!managerSheet) {
    alert('Manager sheet not found.');
    updateButtonState(BUTTON_STATE);
    return;
  }
  
  ACCOUNTS = getAccountsFromManagerSheet(managerSheet);
  console.log('Accounts found:', ACCOUNTS);
  
  if (!ACCOUNTS || ACCOUNTS.length === 0) {
    alert('No accounts found for this manager.');
    updateButtonState(BUTTON_STATE);
    return;
  }
  
  const accountHtml = createAccountHtml(ACCOUNTS);
  document.getElementById("accounts_form").innerHTML = accountHtml;
  
  BUTTON_STATE = "CREATE";
  updateButtonState(BUTTON_STATE);
}

// Handle create
async function handleCreate() {
  console.log('Creating buildouts...');
  
  updateButtonState("");
  
  const selectedAccounts = readAccountCheckBoxData();
  
  if (selectedAccounts.length === 0) {
    alert('Please select at least one account.');
    updateButtonState(BUTTON_STATE);
    return;
  }
  
  const formData = readAccountBuildoutData();
  
  const brandSpreadsheet = await fetchSpreadsheet(formData.brandBuildoutSpreadsheetURL);
  const accountSpreadsheet = await fetchSpreadsheetSingleManager(formData.accountDataSpreadsheetURL, MANAGER);
  
  if (!brandSpreadsheet || !accountSpreadsheet) {
    alert('Failed to fetch spreadsheets.');
    updateButtonState(BUTTON_STATE);
    return;
  }
  
  try {
    await processRequest(brandSpreadsheet, accountSpreadsheet, selectedAccounts);
    
    // Reset to initial state
    BUTTON_STATE = "GET_MANAGER_DATA";
    updateButtonState(BUTTON_STATE);
    document.getElementById("accounts_form").innerHTML = "";
    ACCOUNTS = [];
    
  } catch (error) {
    console.error('Process request error:', error);
    alert('Error creating buildouts: ' + error.message);
    updateButtonState(BUTTON_STATE);
  }
}

// Update button state
function updateButtonState(state) {
  switch (state) {
    case "GET_MANAGER_DATA":
      buildoutButtonSpan.textContent = 'Get Manager Data';
      break;
    case "GET_DATA":
      buildoutButtonSpan.textContent = 'Get Data';
      break;
    case "CREATE":
      buildoutButtonSpan.textContent = 'Create Buildouts';
      break;
    default:
      buildoutButtonSpan.textContent = 'Processing...';
  }
}

// Read form data functions
function readAccountBuildoutData() {
  return {
    brandBuildoutSpreadsheetURL: document.getElementById("master_brand_buildout_spreadsheet").value,
    accountDataSpreadsheetURL: document.getElementById("account_data_spreadsheet").value
  };
}

function readManagerSelectData() {
  const select = document.getElementById("manager-select");
  return select ? select.value : null;
}

function readAccountCheckBoxData() {
  const checkboxes = document.querySelectorAll('input[type="checkbox"]:checked');
  return Array.from(checkboxes).map(cb => cb.value || cb.id.replace('-checkbox', ''));
}

// Load patch notes
function loadPatchNotes() {
  fetch('patchnotes.txt')
    .then(response => response.text())
    .then(text => {
      const formattedText = text.replace(/\n/g, '<br>');
      document.getElementById('patch_notes').innerHTML = formattedText;
    })
    .catch(err => {
      document.getElementById('patch_notes').textContent = 'Unable to load patch notes.';
      console.error('Failed to load patchnotes.txt:', err);
    });
}

// Handle create brand template buildout (existing function)
async function handleCreateBrandTemplateBuildoutClick(e) {
  e.preventDefault();
  
  const formData = readBrandBuildoutTemplateData();
  
  if (!formData.campaign || !formData.adGroup || !formData.baseKeyword || !formData.finalUrl) {
    alert('Please fill in all fields.');
    return;
  }
  
  try {
    const spreadsheet = createBrandBuildoutTemplateSpreadsheet(
      formData.campaign,
      formData.adGroup,
      formData.baseKeyword,
      formData.finalUrl
    );
    
    const url = await createNewDocument(spreadsheet);
    if (url) {
      window.open(url, '_blank');
    }
  } catch (error) {
    console.error('Error creating brand template:', error);
    alert('Error creating spreadsheet: ' + error.message);
  }
}

function readBrandBuildoutTemplateData() {
  return {
    campaign: document.getElementById("campaign_name").value,
    adGroup: document.getElementById("ad_group_name").value,
    baseKeyword: document.getElementById("base_keyword_name").value,
    finalUrl: document.getElementById("final_url_name").value
  };
}
