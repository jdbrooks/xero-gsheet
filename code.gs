// XERO ACCOUNT ACCESS
var XERO_CONSUMER_KEY = "AAAAAAAA";
var XERO_USER_AGENT = "TestApp";
var XERO_PEM_KEY = "-----BEGIN RSA PRIVATE KEY----- AAAAAAAA -----END RSA PRIVATE KEY-----";

// SHEET NAMES
var SHEET_TECHNICAL = "Technical";
var SHEET_STATEMENT = "Statement";

// STATEMENT CONSTANTS
var STATEMENT_FETCH_DAYS = 5; // fetch 5 days back from script run
var STATEMENT_MAX_PERIOD = 180; // fetch initial 6 months

function getAccountID(sheet){
  // first attempt to fetch value from a spreadsheet
  var range = sheet.getRange("B1:B1");
  var accountID = range.getValue();
  if ( accountID == ""){    
    // if accountID has not been fetched
    var resp = makeXeroAPIGET("https://api.xero.com/api.xro/2.0/Reports/BankSummary",undefined);
    accountID = resp.Reports[0].Rows[1].Rows[0].Cells[0].Attributes[0].Value;
    range = sheet.getRange("A1:B1");
    range.setValues([["Accound ID",accountID]]);  
    Logger.log("XERO Call - Fetched bank accountID:" + accountID);
  }else
    Logger.log("Use accountID from the spreadsheet:" + accountID);    
  
  return accountID;
}

// seek rows in the de-duplicated array. 
// dedup_array - array in which search is done, copy of a google sheet normally
// fields - array of fields thatwe search a match to
function findRecordInArray(dedup_array, fields){
  var is_found = false;
  for(var i=0; ( i < dedup_array.length ) && ( !is_found ); i++){
    var is_match = true;
    for(var j=0; (j < fields.length) && is_match; j++){
      if( dedup_array[i][j] != fields[j]) is_match =false;
    }
    is_found = is_match;
  }
  return is_found;
}

function updateStatement(){
  var active_spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet_technical = active_spreadsheet.getSheetByName(SHEET_TECHNICAL);      
  var sheet_statement = active_spreadsheet.getSheetByName(SHEET_STATEMENT);    
  fetchStatement(sheet_statement,getAccountID(sheet_technical));
}

function fetchStatement(sheet,accountID){
  // check if statement exists with headers and at least 1 row
  var existing_transactions = [];
  var range;
  var d = new Date();
  if( sheet.getLastRow() == 0){
    // new sheet, no data
    // add header column
    sheet.appendRow(["Date","Description","Reference","Reconciled","Source","Amount","Balance"]);    
    d.setDate(d.getDate() - STATEMENT_MAX_PERIOD);
  }else{
    // already fetched before full statement
    range = sheet.getRange(2,1, sheet.getLastRow() - 1, 7);
    existing_transactions = range.getValues();
    d.setDate(d.getDate() - STATEMENT_FETCH_DAYS);
  }  
  /*var params = {"bankAccountID":accountID, 
                "fromDate": Utilities.formatDate(d, "GMT", "yyyy-MM-dd"),
                "toDate": Utilities.formatDate(new Date(), "GMT", "yyyy-MM-dd")};*/
  var params = {"bankAccountID":accountID, 
                "fromDate": Utilities.formatDate(d, "GMT", "yyyy-MM-dd")};
  var stmt = makeXeroAPIGET("https://api.xero.com/api.xro/2.0/Reports/BankStatement",params);
  Logger.log("Fetched Statement from XERO since " + params["fromDate"]);
  // Logger.log(stmt);
  var last_row = sheet.getLastRow();
  var push_transactions = [];
  var report = stmt.Reports[0];
  var report_section = report.Rows[1];
  var transactions = report_section.Rows;
  for(var i = 0; i < transactions.length; i++){
    var val_arr = transactions[i].Cells;    
    var row = [];
    for(var j = 0; j < val_arr.length; j++) row.push(val_arr[j].Value);
    // skip opening balance and closing balance records
    // and then if there is no exact match, add records to transaction 
    if(( row[1] != "Opening Balance" ) && (row[1] != "Closing Balance"))  
       if( !findRecordInArray(existing_transactions, row))  push_transactions.push(row);
  }
  // add push transactions to the sheet if there are new records
  if (push_transactions.length > 0){
    range = sheet.getRange(sheet.getLastRow()+1,1,push_transactions.length,7);
    range.setValues(push_transactions);
    Logger.log("Added " + push_transactions.length + " new transactions");
  } else
    Logger.log("There are no new transactions on the fetched statement");
}

function makeXeroAPIGET(url,params) {
  
  var oauth_nonce = createGuid();
  var oauth_timestamp = (new Date().valueOf()/1000).toFixed(0);
    
  fetch_url = url; 
  var params_string = "";
  if(params !== undefined) {
    //Logger.log(params);
    var first = true;
    for(var key in params){
      if (!first) params_string = params_string + "&";
      params_string = params_string + key + "=" + encodeURIComponent(params[key]);
      first = false;
    }
    fetch_url = fetch_url + "?" + params_string;
  }
  
  Logger.log("FetchURL: " + fetch_url);
    
  var signatureBase = "GET" + "&" + encodeURIComponent(url) + "&";
  if (params !== undefined) signatureBase = signatureBase + encodeURIComponent(params_string + "&");
  signatureBase = signatureBase + encodeURIComponent("oauth_consumer_key=" + XERO_CONSUMER_KEY + "&oauth_nonce="+oauth_nonce+"&oauth_signature_method=RSA-SHA1&oauth_timestamp="+oauth_timestamp+"&oauth_token=" + XERO_CONSUMER_KEY + "&oauth_version=1.0");

  var rsa = new RSAKey();
  rsa.readPrivateKeyFromPEMString(XERO_PEM_KEY);
  var hashAlg = "sha1";
  var hSig = rsa.signString(signatureBase, hashAlg);
  
  var data = new Array();
  for (var i = 0; i < hSig.length; i += 2) {
    data.push(parseInt("0x" + hSig.substr(i, 2)));
  }
  var oauth_signature = Base64.encode(data);  
  
  var authHeader = "OAuth oauth_token=\"" + XERO_CONSUMER_KEY + "\",oauth_nonce=\"" + oauth_nonce + "\",oauth_consumer_key=\"" + XERO_CONSUMER_KEY + "\",oauth_signature_method=\"RSA-SHA1\",oauth_timestamp=\"" + oauth_timestamp + "\",oauth_version=\"1.0\",oauth_signature=\"" + encodeURIComponent(oauth_signature) + "\"";
  var headers = { "User-Agent": + XERO_USER_AGENT , "Authorization": authHeader, "Accept":"application/json" };
  var options = { "headers": headers, "muteHttpExceptions":true };
  
/*  Logger.log(oauth_signature);
  Logger.log(authHeader);
  Logger.log(signatureBase);*/
  
  var response = UrlFetchApp.fetch(fetch_url, options);
  var responseJSON = response.getContentText();
  //Logger.log(responseJSON);
  return JSON.parse(responseJSON);
}

function createGuid() {
  return 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, function(c) {
    var r = Math.random()*16|0, v = c == 'x' ? r : (r&0x3|0x8);
    return v.toString(16)
      });
}