/**
 * Fill in the following variables
 */
var CLIENT_ID = 'enter this in from Hubspot Developer Account';
var CLIENT_SECRET = 'enter this in from Hubspot Developer Account';
var SCOPE = 'contacts';
var AUTH_URL = "https://app.hubspot.com/oauth/authorize"; //should be the same for all hubspot instances
var TOKEN_URL = "https://api.hubapi.com/oauth/v1/token"; //should be the same for all hubspot instances
var API_URL = "https://api.hubapi.com"; //should be the same for all hubspot instances

/**
 * Create the following sheets in your spreadsheet
 * "Stages"
 * "Deals"
 */
var sheetNameStages = "Stages";
var sheetNameDeals = "Deals";
var sheetNameCustomers = "Customers";
var sheetNameNotes = "Notes";




/**
 * ###########################################################################
 * # ----------------------------------------------------------------------- #
 * # --------------------------- AUTHENTICATION ---------------------------- #
 * # ----------------------------------------------------------------------- #
 * ###########################################################################
 */

/**
 * Authorizes and makes a request to get the deals from Hubspot.
 */
function  getOAuth2Access() {
  var service = getService();
  if (service.hasAccess()) {
    Logger.log('it already had access');
  } else {
    var authorizationUrl = service.getAuthorizationUrl();
    Logger.log('Open the following URL and re-run the script: %s',
        authorizationUrl);
  }
}

/**
 * Reset the authorization state, so that it can be re-tested.
 */
function reset() {
  getService().reset();
}

/**
 * Configures the service.
 */
function getService() {
  return OAuth2.createService('hubspot')
      // Set the endpoint URLs.
      .setTokenUrl(TOKEN_URL)
      .setAuthorizationBaseUrl(AUTH_URL)

      // Set the client ID and secret.
      .setClientId(CLIENT_ID)
      .setClientSecret(CLIENT_SECRET)

      // Set the name of the callback function in the script referenced
      // above that should be invoked to complete the OAuth flow.
      .setCallbackFunction('authCallback')

      // Set the property store where authorized tokens should be persisted.
      .setPropertyStore(PropertiesService.getUserProperties())
      .setScope(SCOPE);
}

/**
 * Handles the OAuth2 callback.
 */
function authCallback(request) {
  var service = getService();
  var authorized = service.handleCallback(request);
  if (authorized) {
    return HtmlService.createHtmlOutput('Success!');
  } else {
    return HtmlService.createHtmlOutput('Denied.');
  }
}

/**
 * Logs the redict URI to register.
 */
function logRedirectUri() {
  Logger.log(getService().getRedirectUri());
}



/**
 * ###########################################################################
 * # ----------------------------------------------------------------------- #
 * # ------------------------------- GET DATA ------------------------------ #
 * # ----------------------------------------------------------------------- #
 * ###########################################################################
 */

/**
 * Get the different stages in your Hubspot pipeline
 * API & Documentation URL: https://developers.hubspot.com/docs/methods/deal-pipelines/get-deal-pipeline
 */
function getStages() {
  // Prepare authentication to Hubspot
  var service = getService();
  var headers = {headers: {'Authorization': 'Bearer ' + service.getAccessToken()}};
  Logger.log(headers);
  
  // API request
  var pipeline_id = 'default'; // enter pipeline id
  var url = API_URL + "/crm-pipelines/v1/pipelines/deals"; //Hubspots V1 pipeline deals endpoint
  var response = UrlFetchApp.fetch(url, headers); //raw JSON object
  var result = JSON.parse(response.getContentText()); // Parse the JSON object to make it ready to process
  var stages = Array();
  
  // Looping through the different pipelines in Hubspot
  result.results.forEach(function(item) {
    if (item.pipelineId == pipeline_id) {
      var result_stages = item.stages;
      // Let's sort the stages by displayOrder
      result_stages.sort(function(a,b) {
        return a.displayOrder-b.displayOrder;
      });
  
      // Let's put all the used stages (id & label) in an array
      result_stages.forEach(function(stage) {
        stages.push([stage.stageId,stage.label]);  
      });
    }
  });
  
  return stages;
}

/**
 * Get the deals from your Hubspot pipeline
 * API & Documentation URL: https://developers.hubspot.com/docs/methods/deals/get-all-deals
 */
function getDeals() {
  // Prepare authentication to Hubspot
  var service = getService();
  var headers = {headers: {'Authorization': 'Bearer ' + service.getAccessToken()}};
  
  // Prepare pagination
  // Hubspot lets you take max 250 deals per request. 
  // make multiple request until we get all the deals.
  var keep_going = true;
  var offset = 0;
  var deals = Array();
  

  while(keep_going)
  {
    // Takes properties from the deals
    var url = API_URL + "/deals/v1/deal/paged?properties=dealstage&properties=deal_contact_source_type&properties=dealname&properties=amount&properties=createdate&properties=expansion_or_renewal&properties=contract_end_date&properties=closedate&includeAssociations=true&limit=250&offset="+offset;
    var response = UrlFetchApp.fetch(url, headers);
    var result = JSON.parse(response.getContentText());
   
    
    // Are there any more results, should we stop the pagination ?
    keep_going = result["has-more"];
    offset = result.offset;
    
    // For each deal, we take the stageId, source & amount
 
    result.deals.forEach(function(deal) {
      //Logger.log(deal.expansion_or_renewal);
 
     // Logger.log('poooooooooooooooooooooooooooooooop');
      var stageId = (deal.properties.hasOwnProperty("dealstage")) ? deal.properties.dealstage.value : "unknown";
      var source = (deal.properties.hasOwnProperty("deal_contact_source_type")) ? deal.properties.deal_contact_source_type.value : "unknown";
      var amount = (deal.properties.hasOwnProperty("amount")) ? deal.properties.amount.value : 0;
      var name = (deal.properties.hasOwnProperty("dealname")) ? deal.properties.dealname.value : "unknown";
      var createdate = (deal.properties.hasOwnProperty("createdate")) ? deal.properties.createdate.value : "unknown";
      var closedate = (deal.properties.hasOwnProperty("closedate")) ? deal.properties.closedate.value : "unknown";
      var associatedCompanyId = deal.associations.associatedCompanyIds[0];
      var expansionRenewal = deal.properties.hasOwnProperty("expansion_or_renewal") ? deal.properties.expansion_or_renewal.value : "unknown"; //these are custome properties so either create them or delete these variables 
      var contractEndDate = deal.properties.hasOwnProperty("contract_end_date") ? deal.properties.contract_end_date.value : "unknown";  //these are custome properties so either create them or delete these variables
      
         
      deals.push([stageId,source,amount,name,createdate,closedate,associatedCompanyId,expansionRenewal,contractEndDate]);
      
    });
  };
  
  return deals;
}

/**
 * Get Customers from companies api which is V2
 * API & Documentation URL: https://developers.hubspot.com/docs/methods/companies/get-all-companies
 */

function getCustomers() {
  // Prepare authentication to Hubspot
  var service = getService();
  var headers = {headers: {'Authorization': 'Bearer ' + service.getAccessToken()}};
  
  // Prepare pagination
  // Hubspot lets you take max 250 companies per request. 
  // make multiple request until we get all the deals.
  var keep_going = true;
  var offset = 0;
  var customers = Array();

  while(keep_going)
  {
    // Takes properties from the companies
    var url = API_URL + "/companies/v2/companies/paged?&properties=lifecyclestage&properties=name&limit=250&offset="+offset;
    var response = UrlFetchApp.fetch(url, headers);
    var result = JSON.parse(response.getContentText());
    
    // Are there any more results, should we stop the pagination ?
    keep_going = result["has-more"];
    offset = result.offset;
    
      // For each deal, we take the stageId, source & amount
    result.companies.forEach(function(company) {
      
      if(company.properties.hasOwnProperty("lifecyclestage") === true){
      
        if(company.properties.lifecyclestage.value === "customer"){
        
          var customerName = company.properties.name.value;
          var customerID = company.companyId;
         
          customers.push([customerID,customerName]);
          
        }
        
      }
          
    });
    
  }
  

  Logger.log(customers);
  return customers;
}


/**
* ###########################################################################
* # ----------------------------------------------------------------------- #
* # -------------------------- WRITE TO SPREADSHEET ----------------------- #
* # ----------------------------------------------------------------------- #
* ###########################################################################
*/

/**
 * Print the different stages in your pipeline to the spreadsheet
 */
function writeStages(stages) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetNameStages);
  
  // Let's put some headers and add the stages to our table
  var matrix = Array(["StageID","StageName"]);
  matrix = matrix.concat(stages);
  
  // Writing the table to the spreadsheet
  var range = sheet.getRange(1,1,matrix.length,matrix[0].length);
  range.setValues(matrix);
}

/**
 * Print the different deals that are in your pipeline to the spreadsheet
 */
function writeDeals(deals) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetNameDeals);
  
  // Let's put some headers and add the deals to our table
  var matrix = Array(["StageID","Source", "Amount", "Name","Createdate","Closedate","associatedCompanyId","expansionRenewal","contractEndDate"]);
  matrix = matrix.concat(deals);
  
  // Writing the table to the spreadsheet
  var range = sheet.getRange(1,1,matrix.length,matrix[0].length);
  range.setValues(matrix);
}

/**
 * Print customers to the spreadsheet
 */
function writeCustomers(customers) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetNameCustomers);
  
  // Let's put some headers and add the deals to our table
  var matrix = Array(["IDs","Customers"]);
  matrix = matrix.concat(customers);
  
  // Writing the table to the spreadsheet
  var range = sheet.getRange(1,1,matrix.length,matrix[0].length);
  range.setValues(matrix);
}



/**
* ###########################################################################
* # ----------------------------------------------------------------------- #
* # -------------------------------- ROUTINE ------------------------------ #
* # ----------------------------------------------------------------------- #
* ###########################################################################
*/

/**
 * This function will update the spreadsheet and is called at Midnight everyday. 
 */
function refresh() {
  var service = getService();
  
  if (service.hasAccess()) {
    var stages = getStages();
    writeStages(stages);
  
    var deals = getDeals();
    writeDeals(deals);
    
     var customers = getCustomers();
    writeCustomers(customers);
    
  } else {
    var authorizationUrl = service.getAuthorizationUrl();
    Logger.log('Open the following URL and re-run the script: %s',
        authorizationUrl);
  }
};

function refreshCustomers() {
  var service = getService();
  
  if (service.hasAccess()) {
    
     var customers = getCustomers();
    writeCustomers(customers);
    
  } else {
    var authorizationUrl = service.getAuthorizationUrl();
    Logger.log('Open the following URL and re-run the script: %s',
        authorizationUrl);
  }
}
