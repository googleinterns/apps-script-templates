/**
 * Returns the index of the username that matches with the owner selected by the
 * user
 * @param {String} ownerSelection The username selected as owner by the user
 * @param {String[]} usernames String array containing usernames of all the
 *     engineers
 * @return {Number} Index of the matched username in the usernames array
 */
function match(ownerSelection, usernames) {
  for (var i = 0; i < usernames.length; i++) {
    if (usernames[i] == ownerSelection) {
      return i;
    }
  }
  return -1;
}

/**
 * Returns new date after addition of given number of days in the given date
 * @param {Date} date The date in which the given number of days are to be added
 * @param {Number} days The number of days to be added in the given date
 * @return {Date} New date after addition of given number of days in the given
 *     date
 */
function addDays(date, days) {
  var result = new Date(date);
  result.setDate(result.getDate() + days);
  return result;
}
/**
 * Returns array of numbers containing coding units per day of each engineer
 * @param {Object[][]} engineerInfo A two dimensional array containing engineer
 *     information from the Team Sheet
 * @param {Number} countEng Number of engineers in the Team
 * @return {Number[]} Array of numbers containing coding units per day of each
 *     engineer
 */
function getCodingUnitsPerDay(engineerInfo, countEng) {
  var codingUnitsPerDay = [];
  for (var i = 0; i < countEng; i++) {
    var codingUnits = engineerInfo[i][1] / 7;
    codingUnitsPerDay.push(codingUnits);
  }
  return codingUnitsPerDay;
}

/**
 * Returns array of strings containing username of each engineer
 * @param {Object[][]} engineerInfo A two dimensional array containing engineer
 *     information from the Team Sheet
 * @param {Number} countEng Number of engineers in the Team
 * @return {String[]} Array of strings containing username of each engineer
 */
function getUsernames(engineerInfo, countEng) {
  var usernames = [];
  for (var i = 0; i < countEng; i++) {
    usernames.push(engineerInfo[i][0]);
  }
  return usernames;
}

/**
 * Returns array of given length with all values initialized to the same date
 * @param {Date} date The date to which the array values are initialized
 * @param {Number} arrayLength Required length of the array
 * @return {Date[]} Array of given length with all values initialized to the
 *     same date
 */
function initializeDate(date, arrayLength) {
  var taskDate = [];
  for (var i = 0; i < arrayLength; i++) {
    taskDate.push([ date ]);
  }
  return taskDate;
}

/**
 * Returns empty array with given number of rows and columns
 * @param {Number} rowCount The number of rows required in the array
 * @param {Number} columnCount The number of columns required in the array
 * @return {Date[]} Empty array with given number of rows and columns
 */
function initializeArray(rowCount, columnCount) {
  var array = [];
  for (var i = 0; i < rowCount; i++) {
    var row = [];
    for (var j = 0; j < columnCount; j++) {
      row.push([]);
    }
    array.push(row);
  }
  return array;
}

/**
 * Updates status cells' value to 'Not Started','Done' or 'In Progress' in the
 * Task Sheet
 * @param {Object[][]} taskTable A two dimensional array containing values of
 *     the table in the Task Sheet
 * @param {Number} tableLength Length of taskTable array
 * @param {Range} statusColumnRange Range object containing range of the status
 *     column in the Task Sheet
 */
function updateStatus(taskTable, tableLength, statusColumnRange) {
  var statusColumnValues = [];
  for (var i = 0; i < tableLength; i++) {
    statusColumnValues.push([ taskTable[i][5] ]);
  }
  for (var i = 0; i < tableLength; i++) {
    var estimatedDays = taskTable[i][6];
    var remainingDays = taskTable[i][7];
    var completedDays = taskTable[i][8];
    if (estimatedDays == '') {
      statusColumnValues[i] = [ '' ];
    } else if (completedDays == '0' || completedDays == '') {
      statusColumnValues[i] = [ 'Not Started' ];
    } else if (remainingDays == '0') {
      statusColumnValues[i] = [ 'Done' ];
    } else {
      statusColumnValues[i] = [ 'In Progress' ];
    }
  }
  statusColumnRange.setValues(statusColumnValues);
}

/**
 * Updates start date, estimated launch date, priority of a task/ milestone in
 * the Task sheet Updates checkboxes reflecting milestone assignment in the Team
 * Sheet
 * @param {Date} projectStart Date object containing start date of the project
 * @param {Date} currDate Date object containing current date of the project
 * @param {Boolean} isProjectStartAfterToday Boolean value to store whether
 *     project start date is after current date (TRUE) or not (FALSE).
 * @param {Sheet} taskSheet Sheet object for the Task Sheet
 * @param {Range} teamTableRange Range object for the checkbox table in the Team
 *     Sheet
 * @param {Number[]} codingUnitsPerDay Array of numbers containing coding units
 *     per day of each engineer
 * @param {String[]} usernames Array of strings containing username of each
 *     engineer
 * @param {Object[][]} taskTable A two dimensional array containing values of
 *     the table in the Task Sheet
 * @param {Number} tableLength Length of taskTable array
 * @param {Number} countEng Number of engineers in the Team
 * @param {Object[][]} checkBoxValues A two dimensional array to reset checkbox
 *     selections in the Team Sheet
 * @param {Object[][]} priorityColumnValues A two dimensional array containing
 *     values of the priority column in the Task Sheet
 */
function updateDatesPriorityCheckbox(
    projectStart, currDate, isProjectStartAfterToday, taskSheet, teamTableRange,
    codingUnitsPerDay, usernames, taskTable, countEng, tableLength,
    checkBoxValues, priorityColumnValues) {
  // Update to include engineer start date
  var taskEngineerStartDate = initializeDate(projectStart, countEng);
  var taskEngineerEndDate;
  if (!isProjectStartAfterToday) {
    taskEngineerEndDate = initializeDate(currDate, countEng);
  }
  var startDateTaskTable = [];
  var endDateTaskTable = [];
  for (var rowIndex = 0; rowIndex < tableLength; rowIndex++) {
    var currentMilestoneNumber = taskTable[rowIndex][10];
    var ownerSelection = taskTable[rowIndex][0];
    var rowType = taskTable[rowIndex][12];
    var isRowTypeTask = (rowType == '0') ? true : false;
    // If row type is 'task' and an owner is selected
    if (isRowTypeTask && ownerSelection != '') {
      var engineerIndex = match(ownerSelection, usernames);
      var estimateWorkDays = taskTable[rowIndex][6];
      var startDateTask = taskEngineerStartDate[engineerIndex];
      startDateTaskTable.push([ startDateTask ]);
      var days = Math.ceil(estimateWorkDays / codingUnitsPerDay[engineerIndex]);
      if (isProjectStartAfterToday) {
        // Calculate end date using estimated work days
        var endDateTask = addDays(startDateTask, days);
        endDateTaskTable.push([ endDateTask ]);
      } else {
        // Calculate end date using remaining work days
        var remainingWorkDays = taskTable[rowIndex][7];
        var startDateToPrintEndDate = taskEngineerEndDate[engineerIndex];
        var daysPrintEndDate =
            Math.ceil(remainingWorkDays / codingUnitsPerDay[engineerIndex]);
        var endDateTask = addDays(startDateToPrintEndDate, daysPrintEndDate);
        endDateTaskTable.push([ endDateTask ]);
        startDateToPrintEndDate = addDays(endDateTask, 1);
        taskEngineerEndDate[engineerIndex] = startDatePrintEndDate;
      }
      startDateTask = addDays(startDateTask, days + 1);
      taskEngineerStartDate[engineerIndex] = startDateTask;
      checkBoxValues[engineerIndex][currentMilestoneNumber] = "=TRUE";
    }
    // If row type is milestone
    else if (!isRowTypeTask) {
      startDateTaskTable.push(
          [ '=if(MINIFS(F:F,M:M,"=' + currentMilestoneNumber +
            '",O:O,"=0") = 0, "", MINIFS(F:F,M:M,"=' + currentMilestoneNumber +
            '",O:O,"=0"))' ]);
      endDateTaskTable.push([ '=if(MAXIFS(G:G,M:M,"=' + currentMilestoneNumber +
                              '",O:O,"=0") = 0, "", MAXIFS(G:G,M:M,"=' +
                              currentMilestoneNumber + '",O:O,"=0"))' ]);
      // First row containing Milestone/Task data is Row 7
      var nextRowNumber = Number(rowIndex) + 8;
      priorityColumnValues[rowIndex] = [
        '=if(iferror(MATCH("HIGH",ArrayFormula(if(($O' + nextRowNumber +
        ':$O) = 0,if(($M' + nextRowNumber + ':$M)=' + currentMilestoneNumber +
        ',$E' + nextRowNumber +
        ':$E,""),"")),0))>0,"HIGH",if(iferror(MATCH("MEDIUM",ArrayFormula(if(($O' +
        nextRowNumber + ':$O) = 0,if(($M' + nextRowNumber +
        ':$M)=' + currentMilestoneNumber + ', $E' + nextRowNumber +
        ':$E,""),"")),0))>0,"MEDIUM", if(iferror(MATCH("LOW",ArrayFormula(if(($O' +
        nextRowNumber + ':$O) = 0,if(($M' + nextRowNumber +
        ':$M)=' + currentMilestoneNumber + ', $E' + nextRowNumber +
        ':$E,""),"")),0))>0, "LOW","")))'
      ];
    } else {
      startDateTaskTable.push([ '' ]);
      endDateTaskTable.push([ '' ]);
    }
  }
  var startDateRange = taskSheet.getRange(7, 6, tableLength, 1);
  startDateRange.setValues(startDateTaskTable);
  var endDateRange = taskSheet.getRange(7, 7, tableLength, 1);
  endDateRange.setValues(endDateTaskTable);
  teamTableRange.setValue("=FALSE");
  teamTableRange.setValues(checkBoxValues);
  var priorityColumn = taskSheet.getRange(7, 5, tableLength);
  priorityColumn.setValues(priorityColumnValues);
}

/**
 * Updates spreadsheet by resetting values of dates, status, priority in the
 * Task Sheet and the checkboxes in the Team Sheet
 */
function updateSpreadsheet() {
  var projectStart = SpreadsheetApp.getActive()
                         .getSheetByName('Tasks')
                         .getRange('H2')
                         .getValue();
  if (!isValidDate(projectStart)) {
    var message = 'Please enter valid date!';
    displayAlertMessage(message);
    return;
  }
  var currDate = new Date();
  var isProjectStartAfterToday = currDate < projectStart ? true : false;
  var taskSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Tasks');
  var teamSheet = SpreadsheetApp.getActive().getSheetByName('Team');
  var teamRow = 6;
  var teamColumn = 3;
  var countEng = teamSize();
  var engineerInfo =
      teamSheet.getRange(teamRow, teamColumn, countEng, 4).getValues();
  var codingUnitsPerDay = getCodingUnitsPerDay(engineerInfo, countEng);
  var usernames = getUsernames(engineerInfo, countEng);
  var firstTaskRow = 7;
  var ownerColumn = 3;
  var priorityColumn = 5;
  var milestoneNumber = sendMilestoneCount();
  var numRows = getLastDataRow(taskSheet) - firstTaskRow + 1;
  // Column C to Column O
  var taskTable =
      taskSheet.getRange(firstTaskRow, ownerColumn, numRows, 13).getValues();
  var teamTableRange =
      SpreadsheetApp.getActive().getSheetByName('Team').getRange(
          6, 6, countEng, milestoneNumber);
  var checkBoxValues = initializeArray(countEng, milestoneNumber);
  var priorityColumnValues = [];
  for (var i = 0; i < numRows; i++) {
    priorityColumnValues.push([ taskTable[i][2] ]);
  }
  var statusColumnRange = taskSheet.getRange(7, 8, numRows, 1);
  updateStatus(taskTable, numRows, statusColumnRange);
  updateDatesPriorityCheckbox(projectStart, currDate, isProjectStartAfterToday,
                              taskSheet, teamTableRange, codingUnitsPerDay,
                              usernames, taskTable, countEng, numRows,
                              checkBoxValues, priorityColumnValues);
}