/**
 * Returns the usernames of the engineers assigned the given milestone
 * @param {Number} milestone The milestone number
 * @return {Object} assignedEngineers Array object containing usernames of the
 *     assigned engineers
 */
function getAssignedEngineers(milestone) {
  var teamSheet = SpreadsheetApp.getActive().getSheetByName('Team');
  var teamRow = 6;     // Engineer rows start from Row 6
  var teamColumn = 3;  // Column C
  // Milestone Columns start from the 7th column
  var milestoneColumn = 6 + milestone;
  var teamNumRows = teamSize();
  var checkBoxRange = teamSheet.getRange(teamRow, milestoneColumn, teamNumRows);
  var isAssignedEngineer = checkBoxRange.getValues();
  var usernames =
      teamSheet.getRange(teamRow, teamColumn, teamNumRows).getValues();
  var assignedEngineers = [];
  for (var i = 0; i < teamNumRows; i++) {
    if (isAssignedEngineer[i] == 'true') {
      assignedEngineers.push(usernames[i]);
    }
  }
  return assignedEngineers;
}