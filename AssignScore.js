// Assigns a score to applicant based on scoring template
function assignScore() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var resultsSheet = ss.getSheetByName('Application Results');
  var ui = SpreadsheetApp.getUi();
  
  // Get the selected range (allows for score to be set for multiple organization's at once)
  var range = resultsSheet.getActiveRange();

  // Check if there are highlighted cells
  if (!range) {
    ui.alert("Please select a range of rows to process.");
    return;
  }

  // Access values from the highlighted range
  var startRow = range.getRow();
  var numRows = range.getNumRows();
  var dataRange = resultsSheet.getRange(startRow, 1, numRows, resultsSheet.getLastColumn());
  var data = dataRange.getValues();

  // Get all data from the sheet, including headers
  var allData = resultsSheet.getDataRange().getValues();
  var headers = allData[0]; // The first row (index 0) is the header row
  
  // Column indices of score breakdown
  var scoreIndex = headers.indexOf('Score');
  var partnerScoreIndex = headers.indexOf('Partner Score');
  var budgetScoreIndex = headers.indexOf('Budget Score');
  var fullTimeEmployeeScoreIndex = headers.indexOf('Full Time Employee Score');
  var partTimeEmployeeScoreIndex = headers.indexOf('Part Time Employee Score');
  var volunteerScoreIndex = headers.indexOf('Volunteer Score');
  var teamScoreIndex = headers.indexOf('Team Score');
  var timeCommitmentScoreIndex = headers.indexOf('Time Commitment Score');
  var formCompletionScoreIndex = headers.indexOf('App Completion Score');

  // Loop through every organization in the selected region
  for (var i = 0; i < data.length; i++) {
    var row = data[i];

    // Extract data from the row using correct indices
    var budget = row[headers.indexOf('Budget')] || '';
    var fullTimeEmployees = parseInt(row[headers.indexOf('Full-time Employees')]) || 0;
    var partTimeEmployees = parseInt(row[headers.indexOf('Part-time Employees')]) || 0;
    var volunteers = parseInt(row[headers.indexOf('Volunteers')]) || 0;
    var participant1Role = row[headers.indexOf('Participant 1 Role')] || '';
    var participant2Role = row[headers.indexOf('Participant 2 Role')] || '';
    var participant3Role = row[headers.indexOf('Participant 3 Role')] || '';
    var participant4Role = row[headers.indexOf('Participant 4 Role')] || '';
    var timeCommitment = row[headers.indexOf('Time commitment')] || '';
    var appsForm = row[headers.indexOf('App Form')] || '';
    var partner = row[headers.indexOf('Partner')] || '';

    // Initialize individual score breakdowns
    var partnerScore = 0;
    var budgetScore = 0;
    var fullTimeEmployeesScore = 0;
    var partTimeEmployeesScore = 0;
    var volunteerScore = 0;
    var teamScore = 0;
    var timeCommitmentScore = 0;
    var formCompletionScore = 0;

    // Calculate scores based on the extracted data

    //PARTNER SCORE
    if (partner.toLowerCase() === "yes") {
      partnerScore = 5;
    } else {
      partnerScore = 1;
    }

    //BUDGET SCORE
    switch (budget) {
      case '$500,001 - 1,000,000': 
        budgetScore = 5;
        break;
      case '$1,000,001 or more':
        budgetScore = 4;
        break;
      case '$250,001 - $500,000':
        budgetScore = 3;
        break;
      case '$100,001 - $250,000':
        budgetScore = 2;
        break;
      case '$0 - $100,000':
        budgetScore = 1;
        break;
      default:
        budgetScore = 0;
    }

    // EMPLOYEE SCORE
    fullTimeEmployeesScore = getEmployeeScore(fullTimeEmployees);
    partTimeEmployeesScore = getEmployeeScore(partTimeEmployees);
    volunteerScore = getEmployeeScore(volunteers);

    // TEAM SCORE
    var combinedRoles = [participant1Role, participant2Role, participant3Role, participant4Role].join(", ");
    teamScore = getTeamScore(combinedRoles);

    // TIME COMMITMENT SCORE
    if (timeCommitment.toLowerCase() === "yes") {
      timeCommitmentScore = 5;
    } else {
      timeCommitmentScore = 1;
    }

    // APP COMPLETION SCORE
    switch (appsForm.toLowerCase()) {
      case 'yes':
        formCompletionScore = 5;
        break;
      case 'half':
        formCompletionScore = 3;
        break;
      case 'no':
        formCompletionScore = 1;
        break;
      default:
        formCompletionScore = 0;
    }

    // Update the scores in the Results sheet
    var resultRow = startRow + i;
    resultsSheet.getRange(resultRow, scoreIndex + 1).setValue(`=SUM(V${resultRow}:AC${resultRow})`);
    resultsSheet.getRange(resultRow, partnerScoreIndex + 1).setValue(partnerScore);
    resultsSheet.getRange(resultRow, budgetScoreIndex + 1).setValue(budgetScore);
    resultsSheet.getRange(resultRow, fullTimeEmployeeScoreIndex + 1).setValue(fullTimeEmployeesScore);
    resultsSheet.getRange(resultRow, partTimeEmployeeScoreIndex + 1).setValue(partTimeEmployeesScore);
    resultsSheet.getRange(resultRow, volunteerScoreIndex + 1).setValue(volunteerScore);
    resultsSheet.getRange(resultRow, teamScoreIndex + 1).setValue(teamScore);
    resultsSheet.getRange(resultRow, timeCommitmentScoreIndex + 1).setValue(timeCommitmentScore);
    resultsSheet.getRange(resultRow, formCompletionScoreIndex + 1).setValue(formCompletionScore);
  }
}

// Helper function to get employee score
function getEmployeeScore(number) {
  if (number > 20) return 5;
  if (number >= 11) return 4;
  if (number >= 6) return 3;
  if (number >= 0) return 2;
  return 1; 
}

// Helper function to get team score
function getTeamScore(roles) {
  if (roles.includes("Executive Director") && roles.includes("Board Member")) return 5;
  if (roles.includes("Board Member")) return 4;
  if (roles.includes("Executive Director")) return 3;
  if (roles.includes("Senior Staff") || roles.includes("Staff")) return 2;
  if (roles.includes("Volunteer")) return 1;
  return 0;
}
