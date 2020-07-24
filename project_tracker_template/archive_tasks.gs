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
