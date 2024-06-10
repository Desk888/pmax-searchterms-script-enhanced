// This script is almost identical to the main.js with the only difference that is optimised to work
// with data visualisation platforms, in this case Looker Studio.

////////////////////////////////////////////////////////////////////

// CONFIGURATIONS

var config = {
  LOG: false,
  DATE_RANGE: last_n_days(30), // Choose the amount of days for data retrieval
  SPREADSHEET_URL: "", // Only include URL until /edit.
  EMAIL_ADDRESSES: "", // Add email address for email alert.
  SHEET_NAME: "PMAX Search Terms" // !DO NOT CHANGE THIS AS IT WILL BREAK THE LOOKER DASHBOARD!
};

////////////////////////////////////////////////////////////////////

// ***DO NOT CHANGE THE CODE BELOW***

function main() {
  
  // Campaign Selection and SQL Query
  var spreadsheet = SpreadsheetApp.openByUrl(config.SPREADSHEET_URL);

  let campaignIterator = AdsApp
    .performanceMaxCampaigns()
    .withCondition("campaign.status = ENABLED")
    .get();

  while (campaignIterator.hasNext()) {
    let campaign = campaignIterator.next();

    let query = AdsApp.report(
      "SELECT campaign_search_term_insight.category_label, metrics.clicks, metrics.impressions, metrics.conversions, metrics.conversions_value " +
      "FROM campaign_search_term_insight " +
      "WHERE campaign_search_term_insight.campaign_id = '" + campaign.getId() + "' " +
      "AND segments.date BETWEEN '" + config.DATE_RANGE.split(',')[0] + "' AND '" + config.DATE_RANGE.split(',')[1] + "' " +
      "ORDER BY metrics.impressions DESC"
    );

    if (config.LOG === true) {
      Logger.log("Report " + campaign.getName() + " contains " + query.rows().totalNumEntities() + " rows.");
    }

    let sheet = getOrCreateSheet(spreadsheet, config.SHEET_NAME); 
    sheet.clear(); 
    query.exportToSheet(sheet);
  } // campaignIterator

  // Send Email Functionality
  var recipientEmails = config.EMAIL_ADDRESSES.split(',');
  var subject = "PMAX Search Terms Report [UK]";
  var body =
    "The PMAX Search Terms Report has been generated and is available at: " +
    config.SPREADSHEET_URL +
    "\n\nReport covers the last " +
    config.DATE_RANGE +
    " days." +
    "\n\nThis is an automated email sent by Google Ads Script.";

  MailApp.sendEmail(recipientEmails.join(','), subject, body);
}

////////////////////////////////////////////////////////////////////

// Create Spreadsheet
function getOrCreateSheet(spreadsheet, sheetName) {
  var sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(sheetName);
  }
  return sheet;
}

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

  var headerRange = sheet.getRange(1, 1, 1, sheet.getLastColumn());
  Logger.log("Header Range: " + headerRange.getA1Notation());
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#4caf50');
  headerRange.setFontColor('white');

  if (sheet.getLastRow() > 1) { 
    var dataRange = sheet.getRange(2, 2, sheet.getLastRow() - 1, sheet.getLastColumn() - 1);
    Logger.log("Data Range: " + dataRange.getA1Notation());
    dataRange.setNumberFormat('#,##0'); 

    var conversionValueRange = sheet.getRange(2, 5, sheet.getLastRow() - 1, 1); 
    Logger.log("Conversion Value Range: " + conversionValueRange.getA1Notation());
    conversionValueRange.setNumberFormat('£#,##0.00'); 

    var range = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn());
    Logger.log("Row Banding Range: " + range.getA1Notation());
    range.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY);
  } else {
    Logger.log("No data rows found to apply formatting.");
  }

  sheet.autoResizeColumns(1, sheet.getLastColumn());
  Logger.log("Completed formatSheet function");
}

////////////////////////////////////////////////////////////////////

// Date Range Logic
function last_n_days(n) {
  var from = new Date();
  var to = new Date();
  to.setDate(to.getDate() - n);
  from.setDate(from.getDate() - 1);

  return google_date_range(to, from);
} // function last_n_days()

function google_date_range(from, to) {

  function google_format(date) {
    var date_array = [
      date.getUTCFullYear(),
      (date.getUTCMonth() + 1).toString().padStart(2, '0'),
      date.getUTCDate().toString().padStart(2, '0')
    ];
    return date_array.join('');
  }

  var inverse = (from > to);
  from = google_format(from);
  to = google_format(to);
  var result = [from, to];

  if (inverse) {
    result = [to, from];
  }

  return result.join(',');
} // function google_date_range()

////////////////////////////////////////////////////////////////////

