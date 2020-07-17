/**
 * Archives tasks with status 'Done'
 */
function archiveTasks() {
  var taskSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Tasks');
  var statusColumn = taskSheet.getRange('H7:H').getValues();
  var serialNumber = taskSheet.getRange('A7:A').getValues();
  var numRows = getLastDataRow(taskSheet);
  for (var i = 0; i < numRows; i++) {
    // If the row is a Task row and status value is 'Done'
    if (serialNumber[i] != '' && statusColumn[i] == 'Done') {
      var row = taskSheet.getRange(Number(i) + 7, 1);
      taskSheet.hideRow(row);  // Hide the completed task
    }
  }
}
