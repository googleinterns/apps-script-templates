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
 * Adds milestone data in a new column in the Summary Sheet
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
  // Milestone data starts from Column B onwards in the Summary Sheet
  var newMilestoneColumn = currentMilestoneNumber + 1;
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
      ', if(Tasks!$H7:$H<>"done",if(isblank(Tasks!$I7:$I),1,0),0),0)))');
  var tasksWithNoDaysA1 = tasksWithNoDaysCell.getA1Notation();
  // Days ahead/behind (+/-) 
  var daysAheadCell = summarySheet.getRange(13, newMilestoneColumn);
  daysAheadCell.setFormula(
      '=if(OR(' + remainingDaysA1 + '>0,' + tasksWithNoDaysA1 + '>0), if(AND(' +
      desiredLaunchDateA1 + '<>"",' + estLaunchDateA1 + '<>""),FLOOR(' +
      desiredLaunchDateA1 + '-' + estLaunchDateA1 + '), ""),"DONE")');
  // Remaining Weeks
  var remainingWeeksCell = summarySheet.getRange(14, newMilestoneColumn);
  remainingWeeksCell.setFormula("=max(CEILING((" + desiredLaunchDateA1 +
                                "-today())/7), 0)");
  // Number of engineers assigned
  var engCountCell = summarySheet.getRange(15, newMilestoneColumn);
  var teamMilestoneColumn = currentMilestoneNumber + 6;
  var a1Notation = SpreadsheetApp.getActiveSpreadsheet()
                       .getSheetByName("Team")
                       .getRange(6, teamMilestoneColumn, 50)
                       .getA1Notation();
  engCountCell.setFormula('=COUNTIF(Team!' + a1Notation + ',TRUE)');
  // Labels and Notes
  var notesCell = summarySheet.getRange(16, newMilestoneColumn);
  notesCell.setFormula("Tasks!L" + newMilestoneRow);
  summarySheet.setColumnWidth(newMilestoneColumn, 140);
  summarySheet.getRange(6, newMilestoneColumn, 11)
      .setBorder(true, true, true, true, false, false, "#c65911",
                 SpreadsheetApp.BorderStyle.SOLID);
}

/**
 * Adds new milestone column in the Team Sheet
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
  var firstEngineerRow = 6;
  // Set Milestone Title
  var titleCell = teamSheet.getRange(5, newMilestoneColumn);
  var title = 'Milestone ' + currentMilestoneNumber;
  titleCell.setValue(title);
  // Add checkboxes in the new milestone column
  if (teamCount > 0) {
    var checkBoxRange =
        teamSheet.getRange(firstEngineerRow, newMilestoneColumn, teamCount, 1);
    var rule = SpreadsheetApp.newDataValidation().requireCheckbox().build();
    checkBoxRange.setDataValidation(rule);
  }
}

/**
 * Adds milestone data in a newly inserted column in the Task Sheet
 * @param {String} milestoneTitle The title of the milestone
 * @param {String} notes Notes and Labels for the milestone
 * @param {Number} lastRow Row Number of the last row containing task or
 *     milestone data
 * @param {Number} previousMilestoneNumber Previous number of milestones in the
 *     spreadsheet
 */
function addMilestoneTaskSheet(milestoneTitle, notes, lastRow,
                               previousMilestoneNumber) {
  var taskSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Tasks");
  var newMilestoneRow = lastRow + 1;
  // The formatting and data validation settings of the newly inserted row are
  // copied from the row before/ after which it is inserted depending on the row
  // insertion method used ( insertRowBefore() or insertRowAfter() )
  if (previousMilestoneNumber == 0) {
    taskSheet.insertRowBefore(newMilestoneRow);
  } else {
    taskSheet.insertRowAfter(lastRow);
  }
  var currentMilestoneNumber = previousMilestoneNumber + 1;
  // Set values in the newly inserted milestone row
  // Milestone Title
  var titleCell = taskSheet.getRange(newMilestoneRow, 2);  // Column B
  titleCell.setValue(milestoneTitle);
  // Start Date of the Milestone
  var startDateCell = taskSheet.getRange(newMilestoneRow, 6);  // Column F
  startDateCell.setFormula('=if(MINIFS(F:F,M:M,"=' + currentMilestoneNumber +
                           '",N:N,">0") = 0, "", MINIFS(F:F,M:M,"=' +
                           currentMilestoneNumber + '",N:N,">0"))');
  // Estimated Launch Date
  var estimatedLaunchDateCell =
      taskSheet.getRange(newMilestoneRow, 7);  // Column G
  estimatedLaunchDateCell.setFormula(
      '=if(MAXIFS(F:F,M:M,"=' + currentMilestoneNumber +
      '",N:N,">0") = 0, "", MAXIFS(F:F,M:M,"=' + currentMilestoneNumber +
      '",N:N,">0"))');
  // Estimated Coding days
  var estimatedDaysCell = taskSheet.getRange(newMilestoneRow, 9);  // Column I
  estimatedDaysCell.setFormula('=SUMIFS(I:I,M:M,"=' + currentMilestoneNumber +
                               '",N:N,">0")');
  // Remaining Coding days
  var remainingDaysCell = taskSheet.getRange(newMilestoneRow, 10);  // Column J
  remainingDaysCell.setFormula("= $I" + newMilestoneRow + "- $K" +
                               newMilestoneRow);
  // Completed Coding Days
  var completedDaysCell = taskSheet.getRange(newMilestoneRow, 11);  // Column K
  completedDaysCell.setFormula('=SUMIFS(K:K,M:M,"=' + currentMilestoneNumber +
                               '",N:N,">0")');
  //  Milestone number
  var milestoneCell = taskSheet.getRange(newMilestoneRow, 13);  // Column M
  milestoneCell.setValue(currentMilestoneNumber);
  // Labels and notes
  var labelsNotes = taskSheet.getRange(newMilestoneRow, 12);  // Column L
  labelsNotes.setValue(notes);
  // Milestone Row is marked with '0' in Task Number Cell
  var taskNumberCell = taskSheet.getRange(newMilestoneRow, 14);  // Column N
  taskNumberCell.setValue('0');
}

/**
 * Adds new milestone data in the spreadsheet
 * @param {String} milestoneTitle The title of the milestone
 * @param {Date} desiredLaunchDate Date object containing desired launch date of
 *     the milestone
 * @param {String} notes Notes for the milestone
 */
function insertMilestoneMain(milestoneTitle, desiredLaunchDate, notes) {
  var taskSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Tasks");
  // Get previous milestone count
  var previousMilestoneNumber = sendMilestoneCount();
  // Get last row containing data in the Task Sheet
  var lastRow = (previousMilestoneNumber > 0 ? getLastDataRow(taskSheet) : 6);
  var newMilestoneRow = lastRow + 1;
  var currentMilestoneNumber = previousMilestoneNumber + 1;
  // Number of tasks in the previous milestone
  var prevTaskCount =
      getTaskNumber(taskSheet, lastRow, previousMilestoneNumber) - 1;
  // Insert Milestone in Task Sheet
  addMilestoneTaskSheet(milestoneTitle, notes, lastRow,
                        previousMilestoneNumber);
  // Insert Milestone in Team Sheet
  addMilestoneTeam(currentMilestoneNumber);
  // Insert Milestone in Summary Sheet
  addMilestoneSummary(desiredLaunchDate, currentMilestoneNumber,
                      newMilestoneRow);
  // The first milestone row and a milestone row without tasks above it don't
  // require ungrouping
  if (currentMilestoneNumber > 1 && prevTaskCount >= 1) {
    ungroupMilestone(taskSheet, newMilestoneRow);
  }
  // Update spreadsheet values
  updateSpreadsheet();
  return true;
}