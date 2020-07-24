/**
 * Copyright 2020 Google LLC
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *      https://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */

const TASK_TITLE_COLUMN = 2;
const TASK_START_DATE_COLUMN = 6;
const TASK_ESTIMATED_LAUNCH_DATE_COLUMN = 7;
const TASK_ESTIMATED_DAYS_COLUMN = 9;
const TASK_REMAINING_DAYS_COLUMN = 10;
const TASK_COMPLETED_DAYS_COLUMN = 11;
const TASK_LABELS_COLUMN = 12;
const TASK_MILESTONE_COLUMN = 13;
const TASK_NUMBER_COLUMN = 14;
const SUMMARY_TITLE_ROW = 5;
const SUMMARY_DESIRED_LAUNCH_DATE_ROW = 6;
const SUMMARY_ESTIMATED_DAYS_ROW = 7;
const SUMMARY_DAYS_COMPLETED_ROW = 8;
const SUMMARY_REMAINING_DAYS_ROW = 9;
const SUMMARY_START_DATE_ROW = 10;
const SUMMARY_ESTIMATED_LAUNCH_DATE_ROW = 11;
const SUMMARY_TASKS_WITH_NO_DAYS_ROW = 12;
const SUMMARY_DAYS_AHEAD_ROW = 13;
const SUMMARY_REMAINING_WEEKS_ROW = 14;
const SUMMARY_ENG_COUNT_ROW = 15;
const SUMMARY_NOTES_ROW = 16;
const TEAM_TITLE_ROW = 5;
const TEAM_FIRST_ENGINEER_ROW = 6;

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
  var titleCell = summarySheet.getRange(SUMMARY_TITLE_ROW, newMilestoneColumn);
  var title = 'Milestone ' + currentMilestoneNumber;
  titleCell.setValue(title);
  // Desired Launch Date
  var desiredLaunchDateCell = summarySheet.getRange(
      SUMMARY_DESIRED_LAUNCH_DATE_ROW, newMilestoneColumn);
  desiredLaunchDateCell.setValue(desiredLaunchDate);
  var desiredLaunchDateA1 = desiredLaunchDateCell.getA1Notation();
  // Estimated Coding Days
  var estimatedDaysCell =
      summarySheet.getRange(SUMMARY_ESTIMATED_DAYS_ROW, newMilestoneColumn);
  estimatedDaysCell.setFormula("Tasks!I" + newMilestoneRow);
  // Completed Coding Days
  var daysCompletedCell =
      summarySheet.getRange(SUMMARY_DAYS_COMPLETED_ROW, newMilestoneColumn);
  daysCompletedCell.setFormula("Tasks!K" + newMilestoneRow);
  // Remaining Coding Days
  var remainingDaysCell =
      summarySheet.getRange(SUMMARY_REMAINING_DAYS_ROW, newMilestoneColumn);
  remainingDaysCell.setFormula("Tasks!J" + newMilestoneRow);
  var remainingDaysA1 = remainingDaysCell.getA1Notation();
  // Start Date
  var startDateCell =
      summarySheet.getRange(SUMMARY_START_DATE_ROW, newMilestoneColumn);
  startDateCell.setFormula("Tasks!F" + newMilestoneRow);
  // Estimated Launch Date
  var estLaunchDateCell = summarySheet.getRange(
      SUMMARY_ESTIMATED_LAUNCH_DATE_ROW, newMilestoneColumn);
  estLaunchDateCell.setFormula("Tasks!G" + newMilestoneRow);
  var estLaunchDateA1 = estLaunchDateCell.getA1Notation();
  // Tasks with no days
  var tasksWithNoDaysCell =
      summarySheet.getRange(SUMMARY_TASKS_WITH_NO_DAYS_ROW, newMilestoneColumn);
  tasksWithNoDaysCell.setFormula(
      '=ArrayFormula(sum(if(FLOOR(Tasks!$A7:$A)=' + currentMilestoneNumber +
      ', if(Tasks!$H7:$H<>"done",if(isblank(Tasks!$I7:$I),1,0),0),0)))');
  var tasksWithNoDaysA1 = tasksWithNoDaysCell.getA1Notation();
  // Days ahead/behind (+/-)
  var daysAheadCell =
      summarySheet.getRange(SUMMARY_DAYS_AHEAD_ROW, newMilestoneColumn);
  daysAheadCell.setFormula(
      '=if(OR(' + remainingDaysA1 + '>0,' + tasksWithNoDaysA1 + '>0), if(AND(' +
      desiredLaunchDateA1 + '<>"",' + estLaunchDateA1 + '<>""),FLOOR(' +
      desiredLaunchDateA1 + '-' + estLaunchDateA1 + '), ""),"DONE")');
  // Remaining Weeks
  var remainingWeeksCell =
      summarySheet.getRange(SUMMARY_REMAINING_WEEKS_ROW, newMilestoneColumn);
  remainingWeeksCell.setFormula("=max(CEILING((" + desiredLaunchDateA1 +
                                "-today())/7), 0)");
  // Number of engineers assigned
  var engCountCell =
      summarySheet.getRange(SUMMARY_ENG_COUNT_ROW, newMilestoneColumn);
  var teamMilestoneColumn = currentMilestoneNumber + 6;
  var a1Notation = SpreadsheetApp.getActiveSpreadsheet()
                       .getSheetByName("Team")
                       .getRange(6, teamMilestoneColumn, 50)
                       .getA1Notation();
  engCountCell.setFormula('=COUNTIF(Team!' + a1Notation + ',TRUE)');
  // Labels and Notes
  var notesCell = summarySheet.getRange(SUMMARY_NOTES_ROW, newMilestoneColumn);
  notesCell.setFormula("Tasks!L" + newMilestoneRow);
  // Set column width
  summarySheet.setColumnWidth(newMilestoneColumn, 140);
  // Set border for new column in Summary Table
  summarySheet.getRange(SUMMARY_DESIRED_LAUNCH_DATE_ROW, newMilestoneColumn, 11)
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
  // Set Milestone Title
  var titleCell = teamSheet.getRange(TEAM_TITLE_ROW, newMilestoneColumn);
  var title = 'Milestone ' + currentMilestoneNumber;
  titleCell.setValue(title);
  // Add checkboxes in the new milestone column
  if (teamCount > 0) {
    var checkBoxRange = teamSheet.getRange(TEAM_FIRST_ENGINEER_ROW,
                                           newMilestoneColumn, teamCount, 1);
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
  var titleCell =
      taskSheet.getRange(newMilestoneRow, TASK_TITLE_COLUMN);  // Column B
  titleCell.setValue(milestoneTitle);
  // Start Date of the Milestone
  var startDateCell =
      taskSheet.getRange(newMilestoneRow, TASK_START_DATE_COLUMN);  // Column F
  startDateCell.setFormula('=if(MINIFS(F:F,M:M,"=' + currentMilestoneNumber +
                           '",N:N,">0") = 0, "", MINIFS(F:F,M:M,"=' +
                           currentMilestoneNumber + '",N:N,">0"))');
  // Estimated Launch Date
  var estimatedLaunchDateCell = taskSheet.getRange(
      newMilestoneRow, TASK_ESTIMATED_LAUNCH_DATE_COLUMN);  // Column G
  estimatedLaunchDateCell.setFormula(
      '=if(MAXIFS(F:F,M:M,"=' + currentMilestoneNumber +
      '",N:N,">0") = 0, "", MAXIFS(F:F,M:M,"=' + currentMilestoneNumber +
      '",N:N,">0"))');
  // Estimated Coding days
  var estimatedDaysCell = taskSheet.getRange(
      newMilestoneRow, TASK_ESTIMATED_DAYS_COLUMN);  // Column I
  estimatedDaysCell.setFormula('=SUMIFS(I:I,M:M,"=' + currentMilestoneNumber +
                               '",N:N,">0")');
  // Remaining Coding days
  var remainingDaysCell = taskSheet.getRange(
      newMilestoneRow, TASK_REMAINING_DAYS_COLUMN);  // Column J
  remainingDaysCell.setFormula("= $I" + newMilestoneRow + "- $K" +
                               newMilestoneRow);
  // Completed Coding Days
  var completedDaysCell = taskSheet.getRange(
      newMilestoneRow, TASK_COMPLETED_DAYS_COLUMN);  // Column K
  completedDaysCell.setFormula('=SUMIFS(K:K,M:M,"=' + currentMilestoneNumber +
                               '",N:N,">0")');
  // Labels and notes
  var labelsNotes =
      taskSheet.getRange(newMilestoneRow, TASK_LABELS_COLUMN);  // Column L
  labelsNotes.setValue(notes);
  //  Milestone number
  var milestoneCell =
      taskSheet.getRange(newMilestoneRow, TASK_MILESTONE_COLUMN);  // Column M
  milestoneCell.setValue(currentMilestoneNumber);
  // Milestone Row is marked with '0' in Task Number Cell
  var taskNumberCell =
      taskSheet.getRange(newMilestoneRow, TASK_NUMBER_COLUMN);  // Column N
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
