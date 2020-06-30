/**
 * Ungroups the newly inserted milestone row
 * @param {Sheet} taskSheet The sheet in which the row is to be ungrouped
 * @param {Number} rowNumber The row number of the row to be ungrouped
 */
function ungroupMilestone(taskSheet, rowNumber) {
  var rowRange = (rowNumber) + ':' + (rowNumber);
  taskSheet.getRange(rowRange).activate().shiftRowGroupDepth(-1);
}

/**
 * Adds milestone data in a newly inserted column in the Summary Sheet
 * @param {Date} desiredLaunchDate Date object containing desired launch date of
 *     the milestone
 * @param {Number} currentMilestoneNumber Milestone Number of the latest
 *     milestone
 * @param {Number} newMilestoneRow Row number of the newly inserted row in Task
 *     Sheet
 */
function addMilestoneSummary(desiredLaunchDate, currentMilestoneNumber,
                             newMilestoneRow) {
  var summarySheet =
      SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Summary");
  var lastMilestoneColumn = currentMilestoneNumber + 1;
  // Milestone data starts from Column C onwards in the Summary Sheet
  var newMilestoneColumn = currentMilestoneNumber + 2;
  summarySheet.insertColumnAfter(lastMilestoneColumn);
  // Milestone title
  var titleCell = summarySheet.getRange(5, newMilestoneColumn);
  var title = 'Milestone ' + currentMilestoneNumber;
  titleCell.setValue(title);
  // Desired Launch Date
  var desiredLaunchDateCell = summarySheet.getRange(6, newMilestoneColumn);
  desiredLaunchDateCell.setValue(desiredLaunchDate);
  var desiredLaunchDateA1 = desiredLaunchDateCell.getA1Notation();
  // Estimated Coding Days
  var estimatedDaysCell = summarySheet.getRange(7, newMilestoneColumn);
  estimatedDaysCell.setFormula("Tasks!I" + newMilestoneRow);
  // Completed Coding Days
  var daysCompletedCell = summarySheet.getRange(8, newMilestoneColumn);
  daysCompletedCell.setFormula("Tasks!K" + newMilestoneRow);
  // Remaining Coding Days
  var remainingDaysCell = summarySheet.getRange(9, newMilestoneColumn);
  remainingDaysCell.setFormula("Tasks!J" + newMilestoneRow);
  var remainingDaysA1 = remainingDaysCell.getA1Notation();
  // Start Date
  var startDateCell = summarySheet.getRange(10, newMilestoneColumn);
  startDateCell.setFormula("Tasks!F" + newMilestoneRow);
  // Estimated Launch Date
  var estLaunchDateCell = summarySheet.getRange(11, newMilestoneColumn);
  estLaunchDateCell.setFormula("Tasks!G" + newMilestoneRow);
  var estLaunchDateA1 = estLaunchDateCell.getA1Notation();
  // Tasks with no days
  var tasksWithNoDaysCell = summarySheet.getRange(12, newMilestoneColumn);
  tasksWithNoDaysCell.setFormula(
      '=ArrayFormula(sum(if(FLOOR(Tasks!$A7:$A)=' + currentMilestoneNumber +
      ', if(Tasks!$H7:$H<>"done",if(isblank(Tasks!$J7:$J),1,0),0),0)))');
  var tasksWithNoDaysA1 = tasksWithNoDaysCell.getA1Notation();
  // Days ahead/behind (+/-)
  var daysAheadCell = summarySheet.getRange(13, newMilestoneColumn);
  daysAheadCell.setFormula('=if(OR(' + remainingDaysA1 + '>0,' +
                           tasksWithNoDaysA1 + '>0),' + desiredLaunchDateA1 +
                           '-' + estLaunchDateA1 + ',"DONE")');
  // Remaining Weeks
  var remainingWeeksCell = summarySheet.getRange(14, newMilestoneColumn);
  remainingWeeksCell.setFormula("=max(CEILING((" + desiredLaunchDateA1 +
                                "-today())/7), 0)");
  // Labels and Notes
  var notesCell = summarySheet.getRange(15, newMilestoneColumn);
  notesCell.setFormula("Tasks!N" + newMilestoneRow);
}

/**
 * Adds milestone column in team table
 * @param {Number} currentMilestoneNumber Milestone Number of the latest
 *     milestone
 */
function addMilestoneTeam(currentMilestoneNumber) {
  var teamSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Team");
  // Milestone Data starts from Column G onwards
  var lastMilestoneColumn = currentMilestoneNumber + 5;
  teamSheet.insertColumnAfter(lastMilestoneColumn);
  var newMilestoneColumn = lastMilestoneColumn + 1;
  var teamCount = teamSize();
  // Set Milestone Title
  var titleCell = teamSheet.getRange(5, newMilestoneColumn);
  var title = 'Milestone ' + currentMilestoneNumber;
  titleCell.setValue(title);
  // Add checkboxes in the new milestone column
  var checkBoxRange = teamSheet.getRange(6, newMilestoneColumn, teamCount,
                                         1);  // 6 is row number of engineer 1
  var rule = SpreadsheetApp.newDataValidation().requireCheckbox().build();
  checkBoxRange.setDataValidation(rule);
}

/**
 * Adds milestone data in a newly inserted column in the Task Sheet
 * @param {String} milestoneTitle The title of the milestone
 * @param {String} notes Notes and Labels for the milestone
 * @param {Number} previousRowTaskNumber Serial Number of the task in the
 *     previous row
 */
function addMilestoneTaskSheet(milestoneTitle, notes, previousRowTaskNumber) {
  var taskSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Tasks");
  var lastRow = getLastDataRow(taskSheet) + 1;
  var newMilestoneRow = lastRow + 1;
  var previousMilestoneNumber = taskSheet.getRange(lastRow, 13).getValue();
  var currentMilestoneNumber = previousMilestoneNumber + 1;
  taskSheet.insertRowAfter(lastRow);
  var prevMilestoneRow = lastRow - previousRowTaskNumber;
  var nextRow = newMilestoneRow + 1;
  // Set values in the newly inserted milestone row
  // Milestone Title
  var titleCell = taskSheet.getRange(newMilestoneRow, 2);  // Column B
  titleCell.setValue(milestoneTitle);
  // Start Date of the Milestone
  var startDateCell = taskSheet.getRange(newMilestoneRow, 6);  // Column F
  startDateCell.setFormula('=if(MINIFS(F:F,M:M,"=' + currentMilestoneNumber +
                           '",O:O,"=0") = 0, "", MINIFS(F:F,M:M,"=' +
                           currentMilestoneNumber + '",O:O,"=0"))');
  // Estimated Launch Date
  var estimatedLaunchDateCell =
      taskSheet.getRange(newMilestoneRow, 7);  // Column G
  estimatedLaunchDateCell.setFormula(
      '=if(MAXIFS(F:F,M:M,"=' + currentMilestoneNumber +
      '",O:O,"=0") = 0, "", MAXIFS(F:F,M:M,"=' + currentMilestoneNumber +
      '",O:O,"=0"))');
  // Estimated Coding days
  var estimatedDaysCell = taskSheet.getRange(newMilestoneRow, 9);  // Column I
  estimatedDaysCell.setFormula('=SUMIFS(I:I,M:M,"=' + currentMilestoneNumber +
                               '",O:O,"=0")');
  //  Remaining Coding days
  var remainingDaysCell = taskSheet.getRange(newMilestoneRow, 10);  // Column J
  remainingDaysCell.setFormula("= $I" + newMilestoneRow + "- $K" +
                               newMilestoneRow);
  // Completed Coding Days
  var completedDaysCell = taskSheet.getRange(newMilestoneRow, 11);  // Column K
  completedDaysCell.setFormula('=SUMIFS(K:K,M:M,"=' + currentMilestoneNumber +
                               '",O:O,"=0")');
  //  Milestone number
  var milestoneCell = taskSheet.getRange(newMilestoneRow, 13);  // Column M
  milestoneCell.setValue(currentMilestoneNumber);
  // Labels and notes
  var labelsNotes = taskSheet.getRange(newMilestoneRow, 14);  // Column N
  labelsNotes.setValue(notes);
  // Milestone Row is marked with '1', Task Row with '0' in Row Type Cell
  var typeCell = taskSheet.getRange(newMilestoneRow, 15);  // Column O
  typeCell.setValue('1');
}

/**
 * Adds new milestone data in the spreadsheet.
 * @param {String} milestoneTitle The title of the milestone
 * @param {Date} desiredLaunchDate Date object containing desired launch date of
 *     the milestone
 * @param {String} notes Notes for the milestone
 */
function insertMilestoneMain(milestoneTitle, desiredLaunchDate, notes) {
  var taskSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Tasks");
  var lastRow = getLastDataRow(taskSheet) +
                1;  // after last task of last milestone = done row (ideally)
  var newMilestoneRow = lastRow + 1;
  var previousMilestoneNumber =
      taskSheet.getRange(lastRow, 13)
          .getValue();  // before row insertion, lastRow has the last task
  var currentMilestoneNumber = previousMilestoneNumber + 1;
  var currentTaskNumber =
      getTaskNumber(taskSheet, lastRow, currentMilestoneNumber);
  var previousRowTaskNumber = currentTaskNumber - 1;
  // Insert Milestone in Task Sheet
  addMilestoneTaskSheet(milestoneTitle, notes, previousRowTaskNumber);
  // Insert Milestone in Team Sheet
  addMilestoneTeam(currentMilestoneNumber);
  // Insert Milestone in Summary Sheet
  addMilestoneSummary(desiredLaunchDate, currentMilestoneNumber,
                      newMilestoneRow);
  // The first milestone row and a milestone row without tasks above it don't
  // require ungrouping
  if (currentMilestoneNumber > 1 && previousRowTaskNumber >= 1) {
    ungroupMilestone(taskSheet, newMilestoneRow);
  }
  // Reload sidebar
  showSidebar();
}