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
    if (engineerInfo[i][1] <= 0) {
      return;
    }
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
    // Clear previous value from status cell
    taskTable[i][7] = [ '' ];
    if (estimatedDays == '') {
      // Status remains empty if estimate days are not input by the user/ equals
      // 0
      continue;
    } else if (completedDays == '') {
      // Status when completed days are not input by the user/ equals 0
      taskTable[i][7] = [ 'Not Started' ];
    } else if (remainingDays == '0') {
      // Status when remaining days are 0
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
 * end date of the task for each engineer) in 'taskEngineerEndDate'.
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
 * the Task sheet 
 * Updates checkboxes reflecting milestone assignment in the Team Sheet
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
    var isRowTypeTask = (rowType > '0') ? true : false;
    // If row type is 'task' and an owner is selected
    if (isRowTypeTask && ownerSelection != '') {
      var engineerIndex = match(ownerSelection, usernames);
      var estimateWorkDays = taskTable[rowIndex][8];
      // Task start date (to display)
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
        var remainingWorkDays = taskTable[rowIndex][9];
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
      // Update the next 'display task start date' of the engineer to the
      // next day of current task's estimated launch date (calculated using
      // current 'display task start date' and estimated work days)
      startDateTask = addDays(startDateTask, days + 1);
      taskEngineerStartDate[engineerIndex] = startDateTask;
      // Check the engineer's cell under the current milestone column in Team
      // Sheet
      checkBoxValues[engineerIndex][currentMilestoneNumber - 1] = "=TRUE";
      // Update task number in Column N
      taskTable[rowIndex][13] = [
        '=IF(ISNUMBER(INDIRECT("R[-1]C[0]", false)),INDIRECT("R[-1]C[0]", false)+1,1)'
      ];
      // Update serial number in Column A
      taskTable[rowIndex][0] = [
        '=IF(ISNUMBER(INDIRECT("R[-1]C[12]", false)),JOIN(".",INDIRECT("R[-1]C[12]", false),INDIRECT("R[-1]C[13]", false)+1))'
      ];

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
      taskTable[rowIndex][13] = [
        '=IF(ISNUMBER(INDIRECT("R[-1]C[0]", false)),INDIRECT("R[-1]C[0]", false)+1,1)'
      ];
      // Update serial number in Column A
      taskTable[rowIndex][0] = [
        '=IF(ISNUMBER(INDIRECT("R[-1]C[12]", false)),JOIN(".",INDIRECT("R[-1]C[12]", false),INDIRECT("R[-1]C[13]", false)+1))'
      ];
    }
    // Update remaining work days for all task and milestone rows
    taskTable[rowIndex][9] = [ '=$I' + rowNumber + '-$K' + rowNumber ];
  }
}

/**
 * Updates all the custom formulas in the Timeline sheet
 * @param {Object[][]} ganttTableValues A two dimensional array containing
 *     information about tasks and milestones
 * @param {Number} numRows Number of rows containing milestone/ task data
 */
function updateGantt(ganttTableValues, numRows) {
  for (var i = 0; i < ganttTableValues.length; i++) {
    for (var j = 0; j < 13; j++) {
      ganttTableValues[i][j] = [ '' ];
    }
  }
  // Set Project Title
  ganttTableValues[0][1] = [ '=Tasks!B1' ];
  // Set Project Start Date
  ganttTableValues[0][2] = [ '=Tasks!H2' ];
  // Set Estimated Project Launch Date
  ganttTableValues[0][3] = [ '=Tasks!H4' ];
  // Set custom formula for week wise representation of the project timeline
  ganttTableValues[0][4] = [
    '=iferror(SPARKLINE({split(rept("7,",FLOOR((int(D6)-int(C6))/7)),","),mod(int(D6)-int(C6),7)},{"charttype","bar";"color1","white";"color2","lightblue"}),"")'
  ];
  ganttTableValues[0][8] = [ 'Milestone' ];
  // Set Task Serial Number from Task Sheet
  ganttTableValues[1][0] = [ '=ArrayFormula(Tasks!$A$7:$A)' ];
  // Set Task/ Milestone Title from Task Sheet
  ganttTableValues[1][1] = [ '=ArrayFormula(Tasks!$B$7:$B)' ];
  // Set Task/ Milestone Start Date from Task Sheet
  ganttTableValues[1][2] = [ '=ArrayFormula(Tasks!$F$7:$F)' ];
  // Set Task/ Milestone Estimated Launch Date from Task Sheet
  ganttTableValues[1][3] = [ '=ArrayFormula(Tasks!$G$7:$G)' ];
  // Set Task/ Milestone Status from Task Sheet
  ganttTableValues[1][5] = [ '=ArrayFormula(Tasks!$H$7:$H)' ];
  // Set Task/ Milestone Owner from Task Sheet
  ganttTableValues[1][6] = [ '=ArrayFormula(Tasks!$C$7:$C)' ];
  // Set the usernames of team members in the Color Legend Block
  ganttTableValues[1][8] = [ '=ArrayFormula(Team!$C$6:$C$25)' ];
  // Set the Task Number in Column M (hidden) of the Tasks/ Milestones from the
  // Task Sheet
  ganttTableValues[1][12] = [ '=ArrayFormula(Tasks!$N$7:$N)' ];
  // Color Legend for Gantt Chart
  ganttTableValues[0][9] = [ '#9fc5e8' ];
  ganttTableValues[1][9] = [ '#3366CC' ];
  ganttTableValues[2][9] = [ '#DC3912' ];
  ganttTableValues[3][9] = [ '#FF9900' ];
  ganttTableValues[4][9] = [ '#43ff43' ];
  ganttTableValues[5][9] = [ '#990099' ];
  ganttTableValues[6][9] = [ '#ffc66f' ];
  ganttTableValues[7][9] = [ '#0099C6' ];
  ganttTableValues[8][9] = [ '#DD4477' ];
  ganttTableValues[9][9] = [ '#66AA00' ];
  ganttTableValues[10][9] = [ '#B82E2E' ];
  ganttTableValues[11][9] = [ '#316395' ];
  ganttTableValues[12][9] = [ '#994499' ];
  ganttTableValues[13][9] = [ '#22AA99' ];
  ganttTableValues[14][9] = [ '#AAAA11' ];
  ganttTableValues[15][9] = [ '#6633CC' ];
  ganttTableValues[16][9] = [ '#E67300' ];
  ganttTableValues[17][9] = [ '#8B0707' ];
  ganttTableValues[18][9] = [ '#329262' ];
  ganttTableValues[19][9] = [ '#5574A6' ];
  ganttTableValues[20][9] = [ '#3B3EAC' ];
  for (var i = 1; i <= numRows; i++) {
    var rowNumber = Number(i) + 6;
    // Evaluate color code for each task in Column L (hidden) using the color
    // legend and owner column
    ganttTableValues[i][11] = [
      '=iferror(INDIRECT(if(iferror(MATCH($G' + rowNumber +
      ',ArrayFormula($I$7:$I$21),0),-1)>0,JOIN("","$J",TEXT(ADD(6,(MATCH($G' +
      rowNumber + ',ArrayFormula($I$7:$I$21)))),"###")))),"#9fc5e8")'
    ];
    // Create bar charts in Column E using the evaluated color codes in column L
    ganttTableValues[i][4] =
        [ '=iferror(SPARKLINE({int(C' + rowNumber + ')-int($C$6),int(D' +
          rowNumber + ')-int(C' + rowNumber +
          ')},{"charttype","bar";"color1","white";"color2",$L' + rowNumber +
          ';"max",int($D$6)-int($C$6)}),"")' ];
  }
}

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
    var message = 'Please enter valid date!';
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