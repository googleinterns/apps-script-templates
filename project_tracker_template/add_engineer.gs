/**
 * Adds new engineer's details in the Team Sheet
 * @param {String} username The username of the engineer
 * @param {String} teamRole The team role of the engineer
 * @param {Number} codingDaysPerWeek The number of coding days the engineer has
 *     per week
 * @param {Date} startDate The date on which the engineer joins the project
 * @param {Date} endDate The date on which the engineer leaves the project
 * @param {Number} countCurrTask Number of tasks the engineer is currently
 *     working on
 * @param {Number} countTotalTask Total number of tasks assigned to the engineer
 * @param {String} initiative The initiatives taken by the engineer
 * @param {String} emailAddress The email address of the engineer
 */
function addEngineer(username, teamRole, codingDaysPerWeek, startDate, endDate,
                     countCurrTask, countTotalTask, initiative, emailAddress) {
  var teamSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Team");
  // Number of milestones
  var countMilestones = sendMilestoneCount();
  // New team size
  var newCountEng = Number(teamSize()) + 1;
  // Engineer Data starts from Row 6 onwards in the Team Sheet
  // Row Number of the new row to be inserted
  var newEngRow = 5 + newCountEng;
  // Row number of the last engineer in the Team sheet
  var lastEngRow = newEngRow - 1;
  // Insert new row below the last engineer's row
  teamSheet.insertRowAfter(lastEngRow);
  var a1Row = newEngRow + ':' + newEngRow;
  var newRowRange = teamSheet.getRange(a1Row);
  // Clear any formatting set in the new engineer's row
  newRowRange.clear();
  // Set values in the newly inserted row
  // Serial number
  var sNoCell = teamSheet.getRange(newEngRow, 1);
  sNoCell.setValue(newCountEng);
  // Role of the engineer
  var teamRoleCell = teamSheet.getRange(newEngRow, 2);
  teamRoleCell.setValue(teamRole);
  // Username of the engineer
  var usernameCell = teamSheet.getRange(newEngRow, 3);
  usernameCell.setValue(username);
  // Number of coding days per week of the engineer
  var codingDaysPerWeekCell = teamSheet.getRange(newEngRow, 4);
  codingDaysPerWeekCell.setValue(codingDaysPerWeek);
  // The date on which the engineer joins the project
  var startDateCell = teamSheet.getRange(newEngRow, 5);
  startDateCell.setValue(startDate);
  // The date on which the engineer leaves the project
  var endDateCell = teamSheet.getRange(newEngRow, 6);
  endDateCell.setValue(endDate);
  // Insert checkboxes under all the milestone columns
  if (countMilestones > 0) {
    // Milestone Data starts from Column G onwards
    var checkBoxRange = teamSheet.getRange(newEngRow, 7, 1, countMilestones);
    var rule = SpreadsheetApp.newDataValidation().requireCheckbox().build();
    checkBoxRange.setDataValidation(rule);
  }
  // Number of tasks the engineer is currently working on
  var countCurrTaskCell = teamSheet.getRange(newEngRow, 7 + countMilestones);
  countCurrTaskCell.setValue(countCurrTask);
  // Number of tasks assigned to the engineer
  var countTotalTaskCell = teamSheet.getRange(newEngRow, 8 + countMilestones);
  countTotalTaskCell.setValue(countTotalTask);
  // Initiatives taken by the engineer
  var initiativeCell = teamSheet.getRange(newEngRow, 9 + countMilestones);
  initiativeCell.setValue(initiative);
  // Email address of the engineer
  var emailAddressCell = teamSheet.getRange(newEngRow, 10 + countMilestones);
  emailAddressCell.setValue(emailAddress);
  // Update spreadsheet values
  updateSpreadsheet();
  return true;
}