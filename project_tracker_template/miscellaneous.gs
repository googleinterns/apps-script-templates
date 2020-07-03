/**
 * Returns if given date is valid
 * @param {Date} date Date object
 * @return {Boolean} True if date is valid else False
 */
function isValidDate(date) {
  return date && Object.prototype.toString.call(date) === "[object Date]" &&
         !isNaN(date);
}

/**
 * Returns date in mm/dd/yyyy format.
 * @param {Date} date Date object
 * @return {String} Date in mm/dd/yyyy format.
 */
function getFormattedDate(date) {
  if (!isValidDate(date)) {
    return '';
  }
  var year = date.getFullYear();
  var month = (1 + date.getMonth()).toString();
  month = month.length > 1 ? month : '0' + month;
  var day = date.getDate().toString();
  day = day.length > 1 ? day : '0' + day;
  return month + '/' + day + '/' + year;
}

/**
 * Opens a dialog box in the user's editor with the given message and an "OK"
 * button
 * @param {String} message The message to be displayed
 */
function displayAlertMessage(message) {
  var spreadsheetUi = SpreadsheetApp.getUi();
  spreadsheetUi.alert(message);
}

/**
 * Returns the row number of the last row containing task or milestone data in
 * the Task Sheet
 * @param {Sheet} taskSheet The sheet object for the Task Sheet
 * @return {Number} Row number of the last row containing task or milestone data
 */
function getLastDataRow(taskSheet) {
  var lastSheetRow = taskSheet.getLastRow()
  var a1Notation = 'M7:M' + lastSheetRow;
  var MColumn = taskSheet.getRange(a1Notation);
  var MData = MColumn.getValues();
  var lastRow;
  for (var rowIndex = MData.length - 1; rowIndex >= 0; rowIndex--) {
    if (Number(MData[rowIndex])) {
      lastRow = rowIndex + 6;
      break;
    }
  }
  return lastRow;
}

/**
 * Returns the number of milestones in the project
 * @return {Number} The number of milestones in the project
 */
function sendMilestoneCount() {
  var taskSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Tasks");
  var countMilestones = taskSheet.getRange('E2').getValue();
  return countMilestones;
}
/**
 * Returns the number of engineers working on the project
 * @return {Number} The number of engineers working on the project
 */
function teamSize() {
  var teamSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Team");
  var countEngineer = teamSheet.getRange('D2').getValue();
  return countEngineer;
}

// Update after the 'Row Type' Column is modified
/**
 * Returns task number of the new task in the given milestone
 * @param {Sheet} taskSheet The sheet object for the Task Sheet
 * @param {Number} lastRow The row number of the last task row in a milestone
 * @return {Number} The task number of the new task in the given milestone
 */
function getTaskNumber(taskSheet, lastRow, milestoneNumber) {
  var numberPrevious = taskSheet.getRange(lastRow, 1).getValue();
  var currentTaskNumber;
  var splitNumberPrevious = numberPrevious.split(".");
  if (numberPrevious == '') {
    currentTaskNumber = 1;
  } else {
    currentTaskNumber = Number(splitNumberPrevious[1]) + 1;
  }
  return currentTaskNumber;
}