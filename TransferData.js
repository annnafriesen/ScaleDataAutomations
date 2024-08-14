// Adds the tab menu to Master Tracker options bar. Creates tab called "Automation" with one dropdown menu item called "Assign Score"
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Automation')
    .addItem('Assign Score', 'assignScore')
    .addToUi();
}

// Transfers data from the raw application data to the results sheet. Transfers relevant rows to help with scoring.
function transferData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var rawDataSheet = ss.getSheetByName('Raw Application Data'); 
  var resultsSheet = ss.getSheetByName('Application Results');

  // Get the data from the Raw Data sheet
  var dataRange = rawDataSheet.getDataRange();
  var rawData = dataRange.getValues();
  
  // Find the column index for the "Processed" column
  var headers = rawData[0];
  var processedIndex = headers.indexOf('Processed in Results Sheet');

  // Check to see if there is a column called "Processed in Results Sheet"
  if (processedIndex === -1) {
    SpreadsheetApp.getUi().alert("Processed in Results Sheet column not found.");
    return; 
  }

  // Only process rows that have not been marked as processed 
  for (var i = 1; i < rawData.length; i++) {
    var row = rawData[i];
    if (row[processedIndex] === 'Yes' || row[processedIndex] === 'Loading App Data') continue; // Skip if already processed or in the process of being updated through Zapier

    // Extract data from the row [NOTE: if columns are renamed, you must adjust the names here]
    // Searching for row by header name to reduce errors when columns are added to sheet
    var cohort = row[headers.indexOf('Please select the cohort you are applying for: ')];
    var orgName = row[headers.indexOf('What is your organization\'s name?')];
    var location = row[headers.indexOf('Province/State')];
    var budget = row[headers.indexOf('What is your organization\'s current annual budget?')];
    var fullTimeEmployees = row[headers.indexOf('How many full-time staff does your organization employ?')];
    var partTimeEmployees = row[headers.indexOf('How many part-time/contract staff does your organization employ?')];
    var volunteers = row[headers.indexOf('How many people regularly volunteer for your organization? This does not include Board members.')];
    var bursary = row[headers.indexOf('Bursary Needed')];
    var communityFoundation = row[headers.indexOf('Who is your regional Community Foundation?')];
    var banking = row[headers.indexOf('Whom do you bank with?')];
    var participant1Role = row[headers.indexOf('What is your role in your organization?')];
    var participant2Role = row[headers.indexOf('Position (1)')];
    var participant3Role = row[headers.indexOf('Position (2)')];
    var participant4Role = row[headers.indexOf('Position (3)')];
    var webpage = row[headers.indexOf('Organization Website')];
    var mission = row[headers.indexOf('What is your organization\'s mission statement?')];
    var timeCommitment = row[headers.indexOf('Time commitment: Is your team able to set aside two days per month, over five months, for online pre-work, virtual sessions, and organization-specific coaching?')];

    // Append the row to the Application Results sheet
    var lastRow = resultsSheet.getLastRow() + 1;
    resultsSheet.appendRow([
      `=SUM(V${lastRow}:AC${lastRow})`, "Pending", orgName, cohort, location, budget, fullTimeEmployees, partTimeEmployees, volunteers, "", bursary, communityFoundation, banking, participant1Role, participant2Role, participant3Role, participant4Role, webpage, mission, timeCommitment
    ]); // NOTE: if new columns are added to the application results sheet, you must add a placeholder value for the column in this array

    // Mark the row as processed in the Raw Application Sheet
    rawDataSheet.getRange(i + 1, processedIndex + 1).setValue('Yes');
  }
}
