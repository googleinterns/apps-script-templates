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
  var displayStartDateArray =
      getDisplayStartDate(projectStart, engineerInfo, countEng);
  // Boolean array storing whether task start date (to display) is after current
  // date
  var isStartDateAfterTodayArray = [];
  // Array containing task start dates to calculate estimated launch dates in
  // the launch date column of Task Sheet
  var actualStartDateArray = [];
  // Initialize arrays 'actualStartDateArray' and 'isStartDateAfterTodayArray'
  getActualStartDate(displayStartDateArray, isStartDateAfterTodayArray,
                     actualStartDateArray, countEng, currDate);
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
      var displayStartDate = displayStartDateArray[engineerIndex];
      // Update Start Date (to display)
      taskTable[rowIndex][5] = [ displayStartDate ];
      // Number of actual days required by an engineer to complete the estimated
      // work days of a task
      var actualEstimatedDays =
          Math.ceil(estimateWorkDays / codingUnitsPerDay[engineerIndex]);
      // If the task start date to be displayed is after current date then the
      // estimated launch date is independent of the current date and the
      // remaining work days
      if (isStartDateAfterTodayArray[engineerIndex]) {
        // Calculate end date using estimated work days
        var endDate = addDays(displayStartDate, actualEstimatedDays);
        // Update End Date
        taskTable[rowIndex][6] = [ endDate ];
      } else {
        // Calculate end date using remaining work days
        var remainingWorkDays = taskTable[rowIndex][9];
        // Start date to calculate estimated launch date
        var actualStartDate = actualStartDateArray[engineerIndex];
        // Number of actual days required by an engineer to complete the
        // remaining work days of a task
        var actualRemainingDays =
            Math.ceil(remainingWorkDays / codingUnitsPerDay[engineerIndex]);
        var endDate = addDays(actualStartDate, actualRemainingDays);
        // Update End Date
        taskTable[rowIndex][6] = [ endDate ];
        // Update the next task start date (for calculating launch date) of the
        // engineer to the next day of current task's estimated launch date
        actualStartDate = addDays(endDate, 1);
        actualStartDateArray[engineerIndex] = actualStartDate;
      }
      // Update the next 'display task start date' of the engineer to the
      // next day of current task's estimated launch date (calculated using
      // current 'display task start date' and actual estimated work days)
      displayStartDate = addDays(displayStartDate, actualEstimatedDays + 1);
      displayStartDateArray[engineerIndex] = displayStartDate;
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