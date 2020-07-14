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
      // Status remains empty if estimate days are not input
      // by the user/ equals 0
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