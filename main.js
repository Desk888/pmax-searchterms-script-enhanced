// PMAX SEARCH TERMS SCRIPT ENHANCED - V1.0.2 (Review Readme file for updates)

// ** INSTRUCTIONS **

// - DATE_RANGE: Please specify (in numbers) the days of search term you would like to retrieve.
// - SPREADSHEET_URL: Create an empty spreadsheet and set its restrictions to open with editing mode.
// - EMAIL_ADDRESSES: Enter the email address to send an alert when the script will run.

// ** NOTE **

// - To run this script, you need to make sure the permission of your sheet is set to 'Anyone with the link' and editor mode enabled.
// - This script doesn't work at MCC level.

////////////////////////////////////////////////////////////////////

// CONFIGURATIONS

let config = {
  LOG: true,  // Set to true for debugging
  DATE_RANGE: last_n_days(30), // Last 30 days - choose your date range in numbers.
  SPREADSHEET_URL: "", // Only include URL until /edit.
  EMAIL_ADDRESSES: "", // Separate multiple emails with a comma.
  FORMATTING_RULES: {
    HIGH_CONV_VALUE: 500 // Set the threshold for high conv.value for conditional formatting.
  }
};

////////////////////////////////////////////////////////////////////

// ***DO NOT CHANGE THE CODE BELOW***

function main() {
  
  // Campaign Selection and SQL Query
  let spreadsheet = SpreadsheetApp.openByUrl(config.SPREADSHEET_URL);
  
  Logger.log("Date range: " + config.DATE_RANGE);

  let campaignIterator = AdsApp
    .performanceMaxCampaigns()
    .withCondition("campaign.status = ENABLED")
    .get();

  while (campaignIterator.hasNext()) {
    let campaign = campaignIterator.next();

    let queryString = "SELECT campaign_search_term_insight.category_label, metrics.clicks, metrics.impressions, metrics.conversions, metrics.conversions_value " +
                      "FROM campaign_search_term_insight " +
                      "WHERE campaign_search_term_insight.campaign_id = '" + campaign.getId() + "' "; // +
                      // "DURING " + config.DATE_RANGE;
    
    Logger.log("Query String: " + queryString);
    
    let query = AdsApp.report(queryString);

    if (config.LOG === true) {
      Logger.log("Report " + campaign.getName() + " contains " + query.rows().totalNumEntities() + " rows.");
    }

    let sheet = checkTab(spreadsheet);
    query.exportToSheet(sheet);
    formatSheet(sheet); 
  } // campaignIterator

  // Send Email Functionality
  let recipientEmails = config.EMAIL_ADDRESSES.split(',');
  let subject = "PMAX Search Terms Report [UK]";
  let body =
    "The PMAX Search Terms Report has been generated and is available at: " +
    config.SPREADSHEET_URL +
    "\n\nReport covers the last " +
    config.DATE_RANGE +
    " days." +
    "\n\nThis is an automated email sent by Google Ads Script.";

  MailApp.sendEmail(recipientEmails.join(','), subject, body);
}

////////////////////////////////////////////////////////////////////

// Spreadsheet Creation
function checkTab(file) {
  let spreadsheet = SpreadsheetApp.openById(file.getId());
  let currentDate = new Date();
  let sheetName = currentDate.toISOString().slice(0, 10);

  let tab = spreadsheet.getSheetByName(sheetName);
  if (tab) {
    if (config.LOG === true) {
      Logger.log("Selected tab " + sheetName);
    }
  } else {
    tab = spreadsheet.insertSheet(sheetName);
    if (config.LOG === true) {
      Logger.log("Created tab " + sheetName);
    }
  }
  
  // Remove default tab in English
  let defaultSheetEnglish = spreadsheet.getSheetByName("Sheet1");
  if (defaultSheetEnglish) {
    spreadsheet.deleteSheet(defaultSheetEnglish);
  }

  return tab;
} // function checkTab

////////////////////////////////////////////////////////////////////

// Spreadsheet Formatting
function formatSheet(sheet) {
  Logger.log("Starting formatSheet function");

  if (!sheet) {
    Logger.log("Error: The sheet object is not valid.");
    return;
  }
  Logger.log("Sheet name: " + sheet.getName());
  Logger.log("Sheet rows: " + sheet.getLastRow() + ", columns: " + sheet.getLastColumn());

  let headerRange = sheet.getRange(1, 1, 1, sheet.getLastColumn());
  Logger.log("Header Range: " + headerRange.getA1Notation());
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#4caf50');
  headerRange.setFontColor('white');

  if (sheet.getLastRow() > 1) { 
    let dataRange = sheet.getRange(2, 2, sheet.getLastRow() - 1, sheet.getLastColumn() - 1);
    Logger.log("Data Range: " + dataRange.getA1Notation());
    dataRange.setNumberFormat('#,##0'); 

    let conversionValueRange = sheet.getRange(2, 5, sheet.getLastRow() - 1, 1); 
    Logger.log("Conversion Value Range: " + conversionValueRange.getA1Notation());
    conversionValueRange.setNumberFormat('£#,##0.00'); 

    let range = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn());
    Logger.log("Row Banding Range: " + range.getA1Notation());
    range.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY);
    
    // Apply conditional formatting 
    let avgCPCRule = SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThan(config.FORMATTING_RULES.HIGH_CONV_VALUE)
      .setBackground("#fabf8f")
      .setRanges([conversionValueRange])
      .build();
    let rules = sheet.getConditionalFormatRules();
    rules.push(avgCPCRule);
    sheet.setConditionalFormatRules(rules);
  } else {
    Logger.log("No data rows found to apply formatting.");
  }

  sheet.autoResizeColumns(1, sheet.getLastColumn());
  Logger.log("Completed formatSheet function");
}

////////////////////////////////////////////////////////////////////

// Date Range Logic
function last_n_days(n) {

  let from = new Date();
  let to = new Date();
  to.setDate(to.getDate() - n);
  from.setDate(from.getDate() - 1);

  return google_date_range(to, from);

} // function last_n_days()

function google_date_range(from, to) {

  function google_format(date) {
    let date_array = [
      date.getUTCFullYear(),
      (date.getUTCMonth() + 1).toString().padStart(2, '0'),
      date.getUTCDate().toString().padStart(2, '0')
    ];
    return date_array.join('');
  }

  let fromFormatted = google_format(from);
  let toFormatted = google_format(to);

  Logger.log("Formatted date range: " + fromFormatted + "," + toFormatted);

  return fromFormatted + "," + toFormatted;

} // function google_date_range()
