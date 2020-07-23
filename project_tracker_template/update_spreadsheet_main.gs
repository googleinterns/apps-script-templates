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

/**
 * Updates spreadsheet by resetting values of start dates, launch dates, status,
 * priority of task/ milestone rows in the Task Sheet and the checkboxes in the
 * Team Sheet
 */
function updateSpreadsheet() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var taskSheet = spreadsheet.getSheetByName('Tasks');
  var teamSheet = spreadsheet.getSheetByName('Team');
  var summarySheet = spreadsheet.getSheetByName('Summary');
  var ganttSheet = spreadsheet.getSheetByName('Timeline');
  var projectStart = taskSheet.getRange('H2').getValue();
  // Display alert if the project start date is invalid
  if (!isValidDate(projectStart)) {
    var message = 'Please enter valid Project Start Date!';
    displayAlertMessage(message);
    return;
  }
  // Set Milestone Count
  taskSheet.getRange('E2').setValue('=MAX(M7:M)');
  // Set Estimated Launch Date of the Project
  taskSheet.getRange('H4').setValue('=if(MAX(G7:G)=0,"",MAX(G7:G))');
  // Set Team Size in Team Sheet
  teamSheet.getRange('D2').setValue('=MAX(A6:A)');
  // Set Team Size in Tasks Sheet
  taskSheet.getRange('E4').setValue('=Team!D2');
  // Set Dropdown List for Owner Column in Tasks Sheet
  taskSheet.getRange('$C$7:$C').setDataValidation(
      SpreadsheetApp.newDataValidation()
          .setAllowInvalid(true)
          .requireValueInRange(spreadsheet.getRange('Team!$C$6:$C$30'), true)
          .build());
  // Set Dropdown List for Owner Column in Timeline Sheet
  ganttSheet.getRange('$G$7:$G').setDataValidation(
      SpreadsheetApp.newDataValidation()
          .setAllowInvalid(true)
          .requireValueInRange(spreadsheet.getRange('Team!$C$6:$C$30'), true)
          .build());
  var currDate = new Date();
  var isProjectStartAfterToday = currDate < projectStart ? true : false;
  var firstEngineerRow = 6;
  var usernameColumn = 3;        // Column C in Team Sheet
  var firstMilestoneColumn = 7;  // Column G in Team Sheet
  var firstTaskRow = 7;
  // Total number of milestones
  var milestoneNumber = sendMilestoneCount();
  // Number of rows containing task/ milestone data
  var numRows = getLastDataRow(taskSheet) - firstTaskRow + 1;
  if (isNaN(numRows) || numRows <= 0) {
    Logger.log('numRows is NaN or <=0');
    return;
  }
  // Update values in Gantt Chart
  var ganttTableRange = ganttSheet.getRange('A6:M');
  var ganttTableValues = ganttTableRange.getValues();
  updateGantt(ganttTableValues, numRows);
  ganttTableRange.setValues(ganttTableValues);
  // Number of engineers in the Team
  var countEng = teamSize();
  // Team size should be greater than 0 for further updates in the spreadsheet
  if (countEng == 0) {
    Logger.log('Team Size is 0');
    return;
  }
  // Two dimensional array containing information of the engineers in the Team
  // Sheet
  var engineerInfo =
      teamSheet.getRange(firstEngineerRow, usernameColumn, countEng, 4)
          .getValues();
  // Array containing coding units per day of each engineer
  var codingUnitsPerDay = getCodingUnitsPerDay(engineerInfo, countEng);
  if (!codingUnitsPerDay) {
    var message = 'Coding days per week should be greater than zero';
    displayAlertMessage(message);
    return;
  }
  // Array containing username of each engineer
  var usernames = getUsernames(engineerInfo, countEng);
  // Column A to Column N in Task Sheet
  var taskTableRange = taskSheet.getRange(firstTaskRow, 1, numRows, 14);
  var taskTable = taskTableRange.getValues();
  // Range containing checkboxes in Team Sheet
  var teamTableRange = teamSheet.getRange(
      firstEngineerRow, firstMilestoneColumn, countEng, milestoneNumber);
  var checkBoxValues = initializeArray(countEng, milestoneNumber);
  updateStatus(taskTable, numRows);
  updateDatesPriorityCheckbox(projectStart, currDate, isProjectStartAfterToday,
                              engineerInfo, codingUnitsPerDay, usernames,
                              taskTable, countEng, numRows, checkBoxValues);
  taskTableRange.setValues(taskTable);
  // Clear all the checkbox selections and then set to the updated values
  teamTableRange.setValue("=FALSE");
  teamTableRange.setValues(checkBoxValues);
}
