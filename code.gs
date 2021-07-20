// XERO ACCOUNT ACCESS
var XERO_USER_AGENT = 'TestApp';
var XERO_CLIENT_ID = 'EMPTY'; // "Client id" from your app at https://developer.xero.com/app/manage
var XERO_CLIENT_SECRET = 'EMPTY'; // "Client secret" from your app at https://developer.xero.com/app/manage
var XERO_REDIRECT_URI = 'http://localhost:5000/callback'; // "OAuth 2.0 redirect URIs" from your app at https://developer.xero.com/app/manage
var XERO_AUTHORIZATION_CODE = 'EMPTY'; // get this from the URL you were redirected to

// SHEET NAMES
var SHEET_TECHNICAL = 'Technical';
var SHEET_STATEMENT = 'Statement';

// STATEMENT CONSTANTS
var STATEMENT_FETCH_DAYS = 5; // fetch 5 days back from script run
var STATEMENT_MAX_PERIOD = 180; // fetch initial 6 months

function updateStatement() {
  var active_spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet_technical = active_spreadsheet.getSheetByName(SHEET_TECHNICAL);
  var sheet_statement = active_spreadsheet.getSheetByName(SHEET_STATEMENT);
  if (!validateToken(sheet_technical)) {
    Logger.log('Token validation failed. Statement update cancelled.');
    return;
  }
  fetchStatement(sheet_statement, getAccountID(sheet_technical));
}

// Validates Access Token and inits refreshing if needed
function validateToken(sheet) {
  var response = makeXeroAPIGET('https://api.xero.com/connections');
  if (!response) {
    return;
  }
  if (JSON.stringify(response).includes('TokenExpired')) {
    Logger.log('Access Token expired! Refreshing...');
    return refreshToken(sheet);
  }
  return true;
}

function makeXeroAPIGET(url) {
  var active_spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet_technical = active_spreadsheet.getSheetByName(SHEET_TECHNICAL);
  var access_token = getAccessToken(sheet_technical);
  if (access_token == '' || access_token === undefined) {
    Logger.log('XERO Call - Could not find Access Token. API call cancelled.');
    return;
  }
  var options = {
    'method': 'GET',
    'headers': {
      'User-Agent': XERO_USER_AGENT,
      'Authorization': 'Bearer ' + access_token,
    },
    'muteHttpExceptions': true,
  }
  var response = UrlFetchApp.fetch(url, options);
  return JSON.parse(response.getContentText());
}

function makeXeroAPIGETWithTenantId(url, params) {
  fetch_url = url;
  var params_string = '';
  if (params !== undefined) {
    var first = true;
    for (var key in params) {
      if (!first) params_string = params_string + '&';
      params_string = params_string + key + '=' + encodeURIComponent(params[key]);
      first = false;
    }
    fetch_url = fetch_url + '?' + params_string;
  }

  var active_spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet_technical = active_spreadsheet.getSheetByName(SHEET_TECHNICAL);
  var access_token = getAccessToken(sheet_technical);
  if (access_token == '' || access_token === undefined) {
    Logger.log('XERO Call - Could not find Access Token. API call cancelled.');
    return;
  }
  tenant_id = getTenantId(sheet_technical);
  if (tenant_id == '' || tenant_id === undefined) {
    Logger.log('XERO Call - Could not find Tenant ID. API call cancelled.');
    return;
  }
  var options = {
    'method': 'GET',
    'headers': {
      'User-Agent': XERO_USER_AGENT,
      'Authorization': 'Bearer ' + access_token,
      'xero-tenant-id': tenant_id,
    },
    'muteHttpExceptions': true,
  }
  var response = UrlFetchApp.fetch(fetch_url, options);
  return JSON.parse(response.getContentText());
}

// STEP 2. Get access token using Client id, Client secret and the code from URI you were redirected at first step
function getAccessToken(sheet) {
  var range = sheet.getRange('B2:B2');
  var access_token = range.getValue();
  if (access_token == '') {
    const fetch_url = 'https://identity.xero.com/connect/token';
    const options = {
      'method': 'post',
      'muteHttpExceptions': true,
      'headers': {
        'User-Agent': XERO_USER_AGENT,
        'Authorization': 'Basic ' + Utilities.base64Encode(XERO_CLIENT_ID + ':' + XERO_CLIENT_SECRET),
      },
      'payload': {
        'grant_type': 'authorization_code',
        'code': XERO_AUTHORIZATION_CODE,
        'redirect_uri': XERO_REDIRECT_URI,
      }
    };
    var response = UrlFetchApp.fetch(fetch_url, options);
    var responseJSON = JSON.parse(response.getContentText());
    rangeAccessToken = sheet.getRange('A2:B2');
    rangeAccessToken.setValues([['Access Token', responseJSON.access_token]]);
    rangeRefreshToken = sheet.getRange('A3:B3');
    rangeRefreshToken.setValues([['Refresh Token', responseJSON.refresh_token]]);
    access_token = responseJSON.access_token;
    if (response.getContentText().includes('invalid_grant')) {
      Logger.log('XERO Call - Could not get Access Token. Please check XERO_AUTHORIZATION_CODE value.');
      Logger.log('Use this authorization URL: ' + buildAuthorizingURI());
      return;
    }
    if (response.getContentText().includes('invalid_client')) {
      Logger.log('XERO Call - Could not get Access Token. Please check XERO_CLIENT_ID and XERO_CLIENT_SECRET values.');
      return;
    }
    if (access_token == '' || access_token === undefined) {
      Logger.log('XERO Call - Could not extract Access Token');
      Logger.log(response.getContentText());
    } else {
      Logger.log('XERO Call - Fetched Access Token: ' + access_token);
    }
  } else {
    Logger.log('Use Access Token from the spreadsheet :' + access_token);
  }
  return access_token;
}

// STEP 3. Get Tenant Id using Access Token
function getTenantId(sheet) {
  // first attempt to fetch value from a spreadsheet
  var range = sheet.getRange('B4:B4');
  var tenantId = range.getValue();
  if (tenantId == '') {
    // if tenantId has not been fetched
    var resp = makeXeroAPIGET('https://api.xero.com/connections', undefined);
    Logger.log(resp);
    tenantId = resp[0].tenantId;
    range = sheet.getRange('A4:B4');
    range.setValues([['Tenant ID', tenantId]]);
    Logger.log('XERO Call - Fetched Tenant ID:' + tenantId);
  } else
    Logger.log('Use Tenant ID from the spreadsheet:' + tenantId);

  return tenantId;
}

function getAccountID(sheet) {
  // first attempt to fetch value from a spreadsheet
  var range = sheet.getRange('B1:B1');
  var accountID = range.getValue();
  if (accountID == '') {
    // if accountID has not been fetched
    var resp = makeXeroAPIGETWithTenantId('https://api.xero.com/api.xro/2.0/Reports/BankSummary', undefined);
    Logger.log(resp);
    accountID = resp.Reports[0].Rows[1].Rows[0].Cells[0].Attributes[0].Value;
    range = sheet.getRange('A1:B1');
    range.setValues([['Accound ID', accountID]]);
    Logger.log('XERO Call - Fetched bank accountID:' + accountID);
  } else
    Logger.log('Use accountID from the spreadsheet:' + accountID);

  return accountID;
}

// Refreshes Access Token once it has expired
function refreshToken(sheet) {
  var range = sheet.getRange('B3:B3');
  var refresh_token = range.getValue();
  if (refresh_token != '') {
    const fetch_url = 'https://identity.xero.com/connect/token';
    const options = {
      'method': 'post',
      'muteHttpExceptions': true,
      'headers': {
        'User-Agent': XERO_USER_AGENT,
        'Authorization': 'Basic ' + Utilities.base64Encode(XERO_CLIENT_ID + ':' + XERO_CLIENT_SECRET),
      },
      'payload': {
        'grant_type': 'refresh_token',
        'client_id': XERO_CLIENT_ID,
        'refresh_token': refresh_token,
      }
    };
    var response = UrlFetchApp.fetch(fetch_url, options);
    var responseJSON = JSON.parse(response.getContentText());
    Logger.log('refresh result');
    Logger.log(responseJSON);
    rangeAccessToken = sheet.getRange('A2:B2');
    rangeAccessToken.setValues([['Access Token', responseJSON.access_token]]);
    rangeRefreshToken = sheet.getRange('A3:B3');
    rangeRefreshToken.setValues([['Refresh Token', responseJSON.refresh_token]]);
    Logger.log('XERO Call - Refreshed Access Token:' + responseJSON.access_token);
    return responseJSON.access_token;
  } else {
    Logger.log('Refresh Token not found. Failed to refresh Access Token!');
  }
  return null;
}

// Returns URI for STEP 1
function buildAuthorizingURI() {
  return `https://login.xero.com/identity/connect/authorize?response_type=code&client_id=${XERO_CLIENT_ID}&redirect_uri=${XERO_REDIRECT_URI}&scope=accounting.transactions.read accounting.reports.read accounting.settings.read&state=xero-gsheet`;
}

function fetchStatement(sheet, accountID) {
  // check if statement exists with headers and at least 1 row
  var existing_transactions = [];
  var range;
  var d = new Date();
  if (sheet.getLastRow() == 0) {
    // new sheet, no data
    // add header column
    sheet.appendRow(['Date', 'Description', 'Reference', 'Reconciled', 'Source', 'Amount', 'Balance']);
    d.setDate(d.getDate() - STATEMENT_MAX_PERIOD);
  } else {
    // already fetched before full statement
    range = sheet.getRange(2, 1, sheet.getLastRow() - 1, 7);
    existing_transactions = range.getValues();
    d.setDate(d.getDate() - STATEMENT_FETCH_DAYS);
  }
  /*var params = {'bankAccountID':accountID, 
                'fromDate': Utilities.formatDate(d, 'GMT', 'yyyy-MM-dd'),
                'toDate': Utilities.formatDate(new Date(), 'GMT', 'yyyy-MM-dd')};*/
  var params = {
    'bankAccountID': accountID,
    'fromDate': Utilities.formatDate(d, 'GMT', 'yyyy-MM-dd')
  };
  var stmt = makeXeroAPIGETWithTenantId('https://api.xero.com/api.xro/2.0/Reports/BankStatement', params);
  Logger.log('Fetched Statement from XERO since ' + params['fromDate']);
  Logger.log(stmt);
  var last_row = sheet.getLastRow();
  var push_transactions = [];
  var report = stmt.Reports[0];
  var report_section = report.Rows[1];
  var transactions = report_section.Rows;
  for (var i = 0; i < transactions.length; i++) {
    var val_arr = transactions[i].Cells;
    var row = [];
    for (var j = 0; j < val_arr.length; j++) row.push(val_arr[j].Value);
    // skip opening balance and closing balance records
    // and then if there is no exact match, add records to transaction 
    if ((row[1] != 'Opening Balance') && (row[1] != 'Closing Balance'))
      if (!findRecordInArray(existing_transactions, row)) push_transactions.push(row);
  }
  // add push transactions to the sheet if there are new records
  if (push_transactions.length > 0) {
    range = sheet.getRange(sheet.getLastRow() + 1, 1, push_transactions.length, 7);
    range.setValues(push_transactions);
    Logger.log('Added ' + push_transactions.length + ' new transactions');
  } else
    Logger.log('There are no new transactions on the fetched statement');
}

// seek rows in the de-duplicated array. 
// dedup_array - array in which search is done, copy of a google sheet normally
// fields - array of fields thatwe search a match to
function findRecordInArray(dedup_array, fields) {
  var is_found = false;
  for (var i = 0; (i < dedup_array.length) && (!is_found); i++) {
    var is_match = true;
    for (var j = 0; (j < fields.length) && is_match; j++) {
      if (dedup_array[i][j] != fields[j]) is_match = false;
    }
    is_found = is_match;
  }
  return is_found;
}
