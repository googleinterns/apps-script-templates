/**
 * Adds grouping for the newly inserted task row
 * @param {Sheet} taskSheet The sheet in which the row is to be grouped
 * @param {Number} rowNumber The row number of the row to be grouped
 */
function groupTask(taskSheet, rowNumber) {
  var rowTask = rowNumber + ':' + rowNumber;
  taskSheet.getRange(rowTask).activate().shiftRowGroupDepth(1);
}

/**
 * Returns the row number of the last row under a milestone (including the
 * milestone row) in the 'Task' Sheet.
 * @param {Sheet} taskSheet The sheet from which row number is to be returned
 * @param {Number} milestoneNumber The serial number of the milestone whose last
 *     row's number is required.
 * @returns {Number} The row number of the last row in a milestone (including
 *     the milestone row)
 */
function getLastMilestoneRow(taskSheet, milestoneNumber) {
  // Get position of the last row that has content
  var lastRow = taskSheet.getLastRow()
  // Column M contains Milestone numbers
  var a1Notation = 'M7:M' + lastRow;
  var milestoneColumn = taskSheet.getRange(a1Notation);
  var milestoneData = milestoneColumn.getValues();
  var lastMilestoneRow;
  for (var i = milestoneData.length - 1; i >= 0; i--) {
    // Milestone and Task data in the Task Sheet is stored Row 7 onwards
    var rowNumber = i + 7;
    if (milestoneData[i] == milestoneNumber) {
      lastMilestoneRow = rowNumber;
      break;
    }
  }
  return lastMilestoneRow;
}

/**
 * Adds task in the spreadsheet with the user input.
 * @param {Number} milestoneNumber Serial number of milestone in which task is
 *     to be added
 * @param {String} taskName Name of the task
 * @param {String} owner Username of engineer to whom task is assigned
 * @param {String} priorityInput Priority of the task
 * @param {Number} estimatedWorkDays Estimated coding days required for the task
 * @param {Number} engDaysCompleted Coding days already completed for the task
 * @param {String} clLink CL Link of the task
 * @param {String} notes Notes for the task
 */
function insertTask(milestoneNumber, taskName, owner, priorityInput,
                    estimatedWorkDays, engDaysCompleted, clLink, notes) {
  var taskSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Tasks");
  // Get row number of last row under the given milestone
  var lastRow = getLastMilestoneRow(taskSheet, milestoneNumber);
  var newTaskRow = lastRow + 1;
  taskSheet.insertRowAfter(lastRow);
  var currentTaskNumber = getTaskNumber(taskSheet, lastRow, milestoneNumber);
  var isFirstTask = (currentTaskNumber == 1 ? true : false);
  var a1Row = newTaskRow + ':' + newTaskRow;
  var newRowRange = taskSheet.getRange(a1Row);
  newRowRange.setBackground(null);
  // Set values in the newly inserted row
  // Task number
  var taskNumber = taskSheet.getRange(newTaskRow, 1);  // Column A
  taskNumber.setValue(
      '=IF(ISNUMBER(INDIRECT("R[-1]C[12]", false)),JOIN(".",INDIRECT("R[-1]C[12]", false),INDIRECT("R[-1]C[13]", false)+1))');
  // Task title
  var titleCell = taskSheet.getRange(newTaskRow, 2);  // Column B
  titleCell.setValue(taskName);
  // Owner
  if (owner != 0) {
    var ownerCell = taskSheet.getRange(newTaskRow, 3);  // Column C
    ownerCell.setValue(owner);
  }
  // CL Link of the task
  var cellCL = taskSheet.getRange(newTaskRow, 4);  // Column D
  cellCL.setValue(clLink);
  // Priority of task
  var priorityCell = taskSheet.getRange(newTaskRow, 5);  // Column E
  priorityCell.setValue(priorityInput);
  // Estimated coding days required for the task
  var estimatedDaysCell = taskSheet.getRange(newTaskRow, 9);  // Column I
  estimatedDaysCell.setValue(estimatedWorkDays);
  // Remaining coding days required for the task
  var remainingDaysCell = taskSheet.getRange(newTaskRow, 10);  // Column J
  remainingDaysCell.setFormula("= $I" + newTaskRow + "- $K" + newTaskRow);
  // Coding days already completed for the task
  var completedDaysCell = taskSheet.getRange(newTaskRow, 11);  // Column K
  completedDaysCell.setValue(engDaysCompleted);
  // Labels and Notes
  var labelsNotes = taskSheet.getRange(newTaskRow, 12);  // Column L
  labelsNotes.setValue(notes);
  // Milestone in which the task is added
  var milestoneCell = taskSheet.getRange(newTaskRow, 13);  // Column M
  milestoneCell.setValue(milestoneNumber);
  // Task Number
  var taskNumberCell = taskSheet.getRange(newTaskRow, 14);  // Column N
  taskNumberCell.setValue(
      '=IF(ISNUMBER(INDIRECT("R[-1]C[0]", false)),INDIRECT("R[-1]C[0]", false)+1,1)');
  // Group the first task row under a milestone
  // New rows inserted below a grouped row automatically get included in the
  // grouping.
  if (isFirstTask) {
    groupTask(taskSheet, newTaskRow);
  }
  // Update spreadsheet values
  updateSpreadsheet();
  return true;
}