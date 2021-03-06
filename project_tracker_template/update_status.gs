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

var STATUS_COLUMN = 7;
var EST_DAYS_COLUMN = 8;
var REMAINING_DAYS_COLUMN = 9;
var COMPLETED_DAYS_COLUMN = 10;

/**
 * Updates status cells' values to 'Not Started','Done' or 'In Progress' in the
 * Task Sheet
 * @param {Object[][]} taskTable A two dimensional array containing values of
 *     the table in the Task Sheet
 * @param {Number} tableLength Length of taskTable array
 */
function updateStatus(taskTable, tableLength) {
  for (var i = 0; i < tableLength; i++) {
    var estimatedDays = taskTable[i][EST_DAYS_COLUMN];   // Column I
    var remainingDays = taskTable[i][REMAINING_DAYS_COLUMN];   // Column J
    var completedDays = taskTable[i][COMPLETED_DAYS_COLUMN];  // Column K
    // Update status cell in Column H in the Task Sheet
    // Clear previous value from status cell
    taskTable[i][STATUS_COLUMN] = [ '' ];
    if (estimatedDays == '') {
      // Status remains empty if estimate days are not input 
      // by the user/ equals 0
      continue;
    } else if (completedDays == '') {
      // Status when completed days are not input by the user/ equals 0
      taskTable[i][STATUS_COLUMN] = [ 'Not Started' ];
    } else if (remainingDays == '0') {
      // Status when remaining days are 0
      taskTable[i][STATUS_COLUMN] = [ 'Done' ];
    } else {
      taskTable[i][STATUS_COLUMN] = [ 'In Progress' ];
    }
  }
}
