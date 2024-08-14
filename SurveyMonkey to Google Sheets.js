/** CREATED: June 3, 2024 ******* AUTHOR: Anna Friesen ********
 * This Google Apps Script transfers data from a SurveyMonkey survey to the Master Tracker.
 * The script:
 * - Extracts specific data fields from SurveyMonkey that correspond to columns in the Master Tracker. 
 * - Avoids duplicating entries based on the organization name and cohort name.
 * - Copies new entries to the Master Tracker, adding on to the bottom of the list.
 * - Deletes the auto-generated SurveyMonkey survey after the transfer is complete.
 * - Displays an alert to the user after the transfer is complete which indicates the number of rows copied and the number of duplicate rows not copied.
 * 
 * Assumptions:
 * - The script is bound to the Master Tracker.
 * - The first sheet in the Master Tracker is named "Alumni Organizations" and the second sheet is named "TNP Waitlist".
 * - The SurveyMonkey survey is auto-populated in the third sheet in the spreadsheet (index 2).
 * - The SurveyMonkey survey follows the naming convention: "Region Season Year ..." or "Season Year ...". (ex. Victoria Spring 2021...)
 * - The headers in the Master Tracker (lines 69-74) and SurveyMonkey survey (lines 60-66) remain unchanged from their original spelling. 
 * - All application surveys have the same questions and wording as they did as of June 3, 2024. 
 * 
 * Functions:
 * - onOpen: adds custom menu in nav bar in Google Sheets for quick access to automation.
 * - transferRowsFromSurveyMonkey: Main function to perform the data transfer and display an alert.
 * - extractTitleFromSheetName: Extracts the cohort title from the source sheet name.
 * - extractYearFromCohort: Extracts the year from the cohort title.
 * - separateRoleandName: Separates the name and role from single data entry.
 * 
 * Recommended Usage: 
 * - Use after each TNP Application survey is closed. This will ensure all survey responses are collected and limit the number of times script is run.
 * 
 * How to Use:
 * STEP 1: Download SurveyMonkey Add-On to Google Sheets by navigating to Extensions > Add-Ons > Get add-ons. 
 * STEP 2: In the SurveyMonkey Add-On sidebar, navigate to the survey you wish to add to the Master Tracker. Click "Start Importing". 
 *         This will create a third sheet in the spreadsheet that contains the responses from the chosen survey. 
 * STEP 3: Once all responses from SurveyMonkey have been imported, click "Stop Importing". 
 * STEP 4: Go to the sheet called "Alumni Organizations" and click "Transfer Rows" under the tab called "SurveyMonkey Automation" in the top navigation bar. 
 * STEP 5: Once all the rows from the SurveyMonkey sheet are transferred over, you will receive an alert telling you how many rows were successfully copied over. 
 *         The SurveyMonkey sheet will be deleted after the transfer to limit the number of empty sheets. 
 * STEP 6: Repeat steps 1 - 4 periodically, after each application survey is closed. 
 */

// Transfer all rows from auto-generated SurveyMonkey sheet into the Master Tracker.  
function transferRowsFromSurveyMonkey() {
  // Access the Master Tracker.
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  // Get the source and destination sheets.
  var destinationSheet = spreadsheet.getSheetByName('Alumni Organizations');
  var sourceSheet = spreadsheet.getSheets()[2]; // Gets third sheet that was auto-created by SurveyMonkey. Second sheet is TNP Waitlist.

  // Grab cohort name from auto-generated SurveyMonkey sheet.
  var currentCohort = extractTitleFromSheetName(sourceSheet.getName());
  var cohortYear = extractYearFromCohort(currentCohort);

  // Get the range of data in the SurveyMonkey sheet.
  var sourceRange = sourceSheet.getDataRange();
  var sourceValues = sourceRange.getValues();
  
  // Get the range of data in the Master Tracker.
  var destinationRange = destinationSheet.getDataRange();
  var destinationValues = destinationRange.getValues();
  
  // Identify column indices for columns to copy over from SurveyMonkey sheet.
  // NOTE: This must be updated if survey questions are changed or reworded.
  var sourceHeaders = sourceValues[0];
  var sourceOrgNameIndex = sourceHeaders.indexOf("What is your organization's name?");
  var sourceNameIndex = sourceHeaders.indexOf("What is your name and position?");
  var sourceEmailIndex = sourceHeaders.indexOf("What is your email address?");
  var sourceRegionsServedIndex = sourceHeaders.indexOf("What region(s) and/or province(s) do you serve?");
  var sourceBudgetIndex = sourceHeaders.indexOf("What is your current organizational budget?");
  var sourceBankIndex = sourceHeaders.indexOf("Whom do you bank with? This is not a mandatory question, although it does help us identify different partnerships and bursary opportunities to support organizations.");
  
  // Identify column indices for corresponding columns in destination sheet.
  var destinationHeaders = destinationValues[1]; // Note that the column headers are in the second row.
  var destOrgNameIndex = destinationHeaders.indexOf("Org Name");
  var destNameIndex = destinationHeaders.indexOf("Name");
  var destEmailIndex = destinationHeaders.indexOf("Email");
  var destRegionsServedIndex = destinationHeaders.indexOf("Location");
  var destBudgetIndex = destinationHeaders.indexOf("Operating Budget");
  var destBankIndex = destinationHeaders.indexOf("Banking info");
  
  // Convert destinationValues to a set for quick lookup when checking for duplicates.
  var destinationSet = new Set();
  for (var i = 2; i < destinationValues.length; i++) { // Start from index 2 to skip the two header rows.
    var orgName = destinationValues[i][destOrgNameIndex];
    var cohortIndex = destinationHeaders.indexOf("Cohort Name");
    var cohort = destinationValues[i][cohortIndex];
    destinationSet.add(orgName + '|' + cohort);
  }
  
  // Find the last row in the destination sheet.
  var lastRow = destinationSheet.getLastRow();
  
  // Variable to keep track of the number of rows copied. 
  var rowsCopied = 0;
  
  // Loop through SurveyMonkey values and append to Master Tracker if not duplicate.
  for (var i = 1; i < sourceValues.length; i++) { // Start from 1 to skip header row.
    var orgName = sourceValues[i][sourceOrgNameIndex];
    var name = separateRoleandName(sourceValues[i][sourceNameIndex])[0];
    var role = separateRoleandName(sourceValues[i][sourceNameIndex])[1];
    var email = sourceValues[i][sourceEmailIndex];
    var regionsServed = sourceValues[i][sourceRegionsServedIndex];
    var budget = sourceValues[i][sourceBudgetIndex];
    var bank = sourceValues[i][sourceBankIndex];
    
    // If organization hasn't already been copied over, then add it to sheet.
    if (!destinationSet.has(orgName + '|' + currentCohort)) {
      rowsCopied++;
      //make a new row of the SurveyMonkey values to copy over
      var newRow = new Array(destinationHeaders.length).fill('');
      newRow[destOrgNameIndex] = orgName;
      newRow[destinationHeaders.indexOf("Cohort Name")] = currentCohort;
      newRow[destinationHeaders.indexOf("Date")] = cohortYear;
      newRow[destNameIndex] = name;
      newRow[destinationHeaders.indexOf("Roll")] = role;
      newRow[destEmailIndex] = email;
      newRow[destRegionsServedIndex] = regionsServed;
      newRow[destBudgetIndex] = budget;
      newRow[destBankIndex] = bank;
      destinationSheet.getRange(lastRow + rowsCopied, 1, 1, newRow.length).setValues([newRow]);
    }
  }

  // Remove the auto-generated SurveyMonkey sheet after transfer to prevent build-up of empty sheets.
  // spreadsheet.deleteSheet(sourceSheet);

  // Display alert once transfer is complete, telling user the number of rows that were successfully transferred over and checks if there were duplicates. 
  // If there were duplicates that weren't copied over, tell the user how many rows already existed in the Master Tracker. 
  if (sourceValues.length > rowsCopied) {
    var countOfDuplicates = (sourceValues.length - 1) - rowsCopied; // You have to -1 from sourceValues.length to account for header row.
     SpreadsheetApp.getUi().alert('Transfer complete! ' + rowsCopied + ' row(s) have been copied.\nThere were '+ countOfDuplicates + " rows that already exist in the Master Tracker that were not copied over.");
  } else {
    SpreadsheetApp.getUi().alert('Transfer complete! ' + rowsCopied + ' rows have been copied.');
  }
}

// Extract the cohort title from the auto-generated SurveyMonkey survey. 
// SurveyMonkey downloads data in sheet name taken from survey. Survey's should start "[Region] [Season] [Year] ..." or "[Season] [Year]" so the cohort name can be extracted. 
function extractTitleFromSheetName(sheetName) {
  // Regular expression pattern to match "Region Season Year" or "Season Year"
  var pattern = /^(?:(.*?)\s+)?(Winter|Spring|Summer|Fall)\s+(\d{4})/;
  
  // Extract matches from the sheet name.
  var matches = sheetName.match(pattern);
  
  // If matches are found, extract the parts.
  if (matches && matches.length >= 3) {
    var region = matches[1] ? matches[1] : ""; // Region is optional
    var season = matches[2];
    var year = matches[3];
    return (region ? region + " " : "") + season + " " + year;
  } else {
    // Return empty string if no matches found
    return "";
  }
}

// Extract the cohort year from the cohort title. 
// If year is not found in title, return empty string. 
function extractYearFromCohort(cohort) {
  var yearMatch = cohort.match(/(\d{4})/);
  return yearMatch? yearMatch[0] : "";
}

// Checks if the user inputted their name AND role. If so, isolates the name and the role into independent strings. 
// Returns array with two entries: name, role. 
// NOTE: people tend to separate name and role by using a comma or a dash, so this function checks for both symbols.
function separateRoleandName(string) {
  // If there is no comma, default to dash.
  var separaterIndex = string.indexOf(",") !== -1 ? string.indexOf(",") : string.indexOf("-");

  if (separaterIndex !== -1) {
    // If comma is found, split the string into name and role.
    var name = string.substring(0, separaterIndex).trim();
    var role = string.substring(separaterIndex + 1).trim();
    
    return [name, role];
  } else {
    // If no comma is found, return the name and an empty role.
    return [string, ""];
  }
}
