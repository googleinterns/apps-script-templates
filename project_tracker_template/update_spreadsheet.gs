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
 * Updates status cells' values to 'Not Started','Done' or 'In Progress' in the
 * Task Sheet
 * @param {Object[][]} taskTable A two dimensional array containing values of
 *     the table in the Task Sheet
 * @param {Number} tableLength Length of taskTable array
 */
function updateStatus(taskTable, tableLength) {
  for (var i = 0; i < tableLength; i++) {
    var estimatedDays = taskTable[i][8];   // Column I
    var remainingDays = taskTable[i][9];   // Column J
    var completedDays = taskTable[i][10];  // Column K
    // Update status cell in Column H in the Task Sheet
    if (estimatedDays == '') {
      taskTable[i][7] = [ '' ];
    } else if (completedDays == '0' || completedDays == '') {
      taskTable[i][7] = [ 'Not Started' ];
    } else if (remainingDays == '0') {
      taskTable[i][7] = [ 'Done' ];
    } else {
      taskTable[i][7] = [ 'In Progress' ];
    }
  }
}

/**
 * Returns an array containing initial task start dates of each engineer
 * Initial Task Start Date for an engineer is the maximum of Project Start Date
 * and the Engineer's Start Date
 * @param {Date} projectStart Date object containing start date of the project
 * @param {Object[][]} engineerInfo A two dimensional array containing
 *     information of the engineers in the Team Sheet
 * @param {Number} countEng Number of engineers in the Team
 * @return {Date[]} Date array containing initial task start dates of each
 *     engineer
 */
function initializeTaskStartDate(projectStart, engineerInfo, countEng) {
  var taskEngineerStartDate = [];
  for (var i = 0; i < countEng; i++) {
    var startDateEng = engineerInfo[i][2];
    if (projectStart < startDateEng) {
      taskEngineerStartDate.push(startDateEng);
    } else {
      taskEngineerStartDate.push(projectStart);
    }
  }
  return taskEngineerStartDate;
}

/**
 * Sets the initial task start dates (to be used for calculating the estimated
 * end date of the task for each engineer) in 'taskEngineerEndDate' 
 * The initial task start dates for these calculations is maximum of the 
 * current date and the displayed task start date for each engineer (where 
 * displayed task start date is maximum of project start date and engineer's 
 * start date) 
 * Sets corresponding value in 'isTaskEngineerStartDateAfterToday' array to 
 * 'true' if the displayed task start date for the engineer is greater than 
 * current date else 'false'
 * @param {Date[]} taskEngineerStartDate Date array containing displayed task
 *     start dates of each engineer
 * @param {Boolean[]} isTaskEngineerStartDateAfterToday Boolean array to store
 *     'true' if the displayed task start date is
 * greater than current date else 'false' for all engineers
 * @param {Date[]} taskEngineerEndDate Date array to store maximum of the
 *     current date and the displayed task start date for each engineer
 * @param {Number} countEng Number of engineers in the Team
 * @param {Date} currDate The current date
 */
function initializeStartForEndDates(taskEngineerStartDate,
                                    isTaskEngineerStartDateAfterToday,
                                    taskEngineerEndDate, countEng, currDate) {
  for (var i = 0; i < countEng; i++) {
    if (taskEngineerStartDate[i] > currDate) {
      taskEngineerEndDate.push(taskEngineerStartDate[i]);
      isTaskEngineerStartDateAfterToday.push(true);
    } else {
      taskEngineerEndDate.push(currDate);
      isTaskEngineerStartDateAfterToday.push(false);
    }
  }
}

/**
 * Updates start date, estimated launch date, priority of a task/ milestone in
 * the Task sheet Updates checkboxes reflecting milestone assignment in the Team
 * Sheet
 * @param {Date} projectStart Date object containing start date of the project
 * @param {Date} currDate Date object containing current date of the project
 * @param {Boolean} isProjectStartAfterToday Boolean value to store whether
 *     project start date is after current date (TRUE) or not (FALSE).
 * @param {Number[]} codingUnitsPerDay Array of numbers containing coding units
 *     per day of each engineer
 * @param {Object[][]} engineerInfo A two dimensional array containing
 *     information of the engineers in the Team Sheet
 * @param {String[]} usernames Array of strings containing username of each
 *     engineer
 * @param {Object[][]} taskTable A two dimensional array containing values of
 *     the table in the Task Sheet
 * @param {Number} tableLength Length of taskTable array
 * @param {Number} countEng Number of engineers in the Team
 * @param {Object[][]} checkBoxValues A two dimensional array to reset checkbox
 *     selections in the Team Sheet
 */
function updateDatesPriorityCheckbox(projectStart, currDate,
                                     isProjectStartAfterToday, engineerInfo,
                                     codingUnitsPerDay, usernames, taskTable,
                                     countEng, tableLength, checkBoxValues) {
  // Array containing task start dates to display in the start date column of
  // Task Sheet
  var taskEngineerStartDate =
      initializeTaskStartDate(projectStart, engineerInfo, countEng);
  // Boolean array storing whether task start date to display is after current
  // date
  var isTaskEngineerStartDateAfterToday = [];
  // Array containing task start dates to calculate estimated launch dates in
  // the launch date column of Task Sheet
  var taskEngineerEndDate = [];
  // Initialize array taskEngineerEndDate with task start dates to calculate
  // estimated launch dates
  initializeStartForEndDates(taskEngineerStartDate,
                             isTaskEngineerStartDateAfterToday,
                             taskEngineerEndDate, countEng, currDate);
  for (var rowIndex = 0; rowIndex < tableLength; rowIndex++) {
    var currentMilestoneNumber = taskTable[rowIndex][12];  // Column M
    var ownerSelection = taskTable[rowIndex][2];           // Column C
    var rowType = taskTable[rowIndex][13];                 // Column N
    var rowNumber = Number(rowIndex) + 7;
    var prevRowNumber = Number(rowNumber) - 1;
    var isRowTypeTask = (rowType > '0') ? true : false;
    // If row type is 'task' and an owner is selected
    if (isRowTypeTask && ownerSelection != '') {
      var engineerIndex = match(ownerSelection, usernames);
      var estimateWorkDays = taskTable[rowIndex][8];
      var startDateTask = taskEngineerStartDate[engineerIndex];
      // Update Start Date
      taskTable[rowIndex][5] = [ startDateTask ];
      // Number of actual days required by an engineer to complete the estimated
      // work days of a task
      var days = Math.ceil(estimateWorkDays / codingUnitsPerDay[engineerIndex]);
      // If the task start date to be displayed is after current date then the
      // estimated launch date is independent of the current date and the
      // remaining work days
      if (isTaskEngineerStartDateAfterToday[engineerIndex]) {
        // Calculate end date using estimated work days
        var endDateTask = addDays(startDateTask, days);
        // Update End Date
        taskTable[rowIndex][6] = [ endDateTask ];
      } else {
        // Calculate end date using remaining work days
        var remainingWorkDays = taskTable[rowIndex][7];
        // Start date to calculate estimated launch date
        var startDateToPrintEndDate = taskEngineerEndDate[engineerIndex];
        // Number of actual days required by an engineer to complete the
        // remaining work days of a task
        var daysToPrintEndDate =
            Math.ceil(remainingWorkDays / codingUnitsPerDay[engineerIndex]);
        var endDateTask = addDays(startDateToPrintEndDate, daysToPrintEndDate);
        // Update End Date
        taskTable[rowIndex][6] = [ endDateTask ];
        // Update the next task start date (for calculating launch date) of the
        // engineer to the next day of current task's estimated launch date
        startDateToPrintEndDate = addDays(endDateTask, 1);
        taskEngineerEndDate[engineerIndex] = startDateToPrintEndDate;
      }
      // Update the next task start date (for display) of the engineer to the
      // next day of current task's estimated launch date
      startDateTask = addDays(startDateTask, days + 1);
      taskEngineerStartDate[engineerIndex] = startDateTask;
      // Check the engineer's cell under the current milestone column in Team
      // Sheet
      checkBoxValues[engineerIndex][currentMilestoneNumber - 1] = "=TRUE";
      // Update task number in Column N
      taskTable[rowIndex][13] = [ '=$N' + prevRowNumber + '+1' ];
      // Update serial number in Column A
      taskTable[rowIndex][0] = [
        '=IF(ISNUMBER(N' + prevRowNumber + '),JOIN(".",M' + prevRowNumber +
        ',N' + prevRowNumber + '+1))'
      ];  //['=JOIN(".",M'+prevRowNumber+',N'+prevRowNumber+' +1)'];
          ////['=IF(ISNUMBER(O'+ prevRowNumber + '),JOIN(".",M'+ prevRowNumber
          //+',N'+prevRowNumber+'+1))'];
    } else if (!isRowTypeTask) {
      // If row type is milestone
      var nextRowNumber = Number(rowNumber) + 1;
      // Update Start Date
      taskTable[rowIndex][5] =
          [ '=if(MINIFS(F:F,M:M,"=' + currentMilestoneNumber +
            '",N:N,">0") = 0, "", MINIFS(F:F,M:M,"=' + currentMilestoneNumber +
            '",N:N,">0"))' ];
      // Update Estimated Launch Date
      taskTable[rowIndex][6] =
          [ '=if(MAXIFS(G:G,M:M,"=' + currentMilestoneNumber +
            '",N:N,">0") = 0, "", MAXIFS(G:G,M:M,"=' + currentMilestoneNumber +
            '",N:N,">0"))' ];
      // Update Priority
      taskTable[rowIndex][4] = [
        '=if(iferror(MATCH("HIGH",ArrayFormula(if(($N' + nextRowNumber +
        ':$N) > 0,if(($M' + nextRowNumber + ':$M)=' + currentMilestoneNumber +
        ',$E' + nextRowNumber +
        ':$E,""),"")),0))>0,"HIGH",if(iferror(MATCH("MEDIUM",ArrayFormula(if(($N' +
        nextRowNumber + ':$N) > 0,if(($M' + nextRowNumber +
        ':$M)=' + currentMilestoneNumber + ', $E' + nextRowNumber +
        ':$E,""),"")),0))>0,"MEDIUM", if(iferror(MATCH("LOW",ArrayFormula(if(($N' +
        nextRowNumber + ':$N) > 0,if(($M' + nextRowNumber +
        ':$M)=' + currentMilestoneNumber + ', $E' + nextRowNumber +
        ':$E,""),"")),0))>0, "LOW","")))'
      ];
      // Update estimated work days
      taskTable[rowIndex][8] =
          [ '=SUMIFS(I:I,M:M,"=' + currentMilestoneNumber + '",N:N,">0")' ];
      // Update completed work days
      taskTable[rowIndex][10] =
          [ '=SUMIFS(K:K,M:M,"=' + currentMilestoneNumber + '",N:N,">0")' ];
      // Update task number in Column N
      taskTable[rowIndex][13] = [ '0' ];
    } else {
      // If row type is task and no owner is selected then
      // clear Start Date and Estimated Launch Date cells
      taskTable[rowIndex][5] = [ '' ];
      taskTable[rowIndex][6] = [ '' ];
      // Update task number in Column N
      taskTable[rowIndex][13] = [ '=$N' + prevRowNumber + '+1' ];
      // Update serial number in Column A
      taskTable[rowIndex][0] =
          [ '=IF(ISNUMBER(N' + prevRowNumber + '),JOIN(".",M' + prevRowNumber +
            ',N' + prevRowNumber + '+1))' ];
    }
    // Update remaining work days for all task and milestone rows
    taskTable[rowIndex][9] = [ '=$I' + rowNumber + '-$K' + rowNumber ];
  }
}

/**
 * Updates spreadsheet by resetting values of start dates, launch dates, status,
 * priority of task/ milestone rows in the Task Sheet and the checkboxes in the
 * Team Sheet
 */
function updateSpreadsheet() {
  var taskSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Tasks');
  var teamSheet = SpreadsheetApp.getActive().getSheetByName('Team');
  var projectStart = taskSheet.getRange('H2').getValue();
  if (!isValidDate(projectStart)) {
    var message = 'Please enter valid date!';
    displayAlertMessage(message);
    return;
  }
  var currDate = new Date();
  var isProjectStartAfterToday = currDate < projectStart ? true : false;
  var firstEngineerRow = 6;
  var usernameColumn = 3;        // Column C in Team Sheet
  var firstMilestoneColumn = 7;  // Column G in Team Sheet
  // Number of engineers in the Team
  var countEng = teamSize();
  // Two dimensional array containing information of the engineers in the Team
  // Sheet
  var engineerInfo =
      teamSheet.getRange(firstEngineerRow, usernameColumn, countEng, 4)
          .getValues();
  // Array containing coding units per day of each engineer
  var codingUnitsPerDay = getCodingUnitsPerDay(engineerInfo, countEng);
  // Array containing username of each engineer
  var usernames = getUsernames(engineerInfo, countEng);
  var firstTaskRow = 7;
  // Total number of milestones
  var milestoneNumber = sendMilestoneCount();
  // Number of rows containing milestone/ task data
  var numRows = getLastDataRow(taskSheet) - firstTaskRow + 1;
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