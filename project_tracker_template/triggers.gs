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
 * The event handler triggered when the attached spreadsheet is opened.
 * @param {Event} e The onOpen event object.
 */
function onOpen(e) {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Custom Menu')
      .addItem('Backup Spreadsheet', 'backup')
      .addItem('Archive Done Tasks', 'archiveTasks')
      .addItem('Mail Summary', 'sendEmail')
      .addItem('Add', 'showSidebar')
      .addToUi();
}

/**
 * The event handler triggered when an edit is made to the attached spreadsheet.
 * @param {Event} e The onEdit event object.
 */
function onEdit(e) {
  updateSpreadsheet();
}
