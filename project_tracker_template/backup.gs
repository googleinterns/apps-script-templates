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
 * Makes a backup copy of the current spreadsheet.
 * Names it appropriately and provides the backup copy's URL as a hyperlink in a
 * modal dialog box.
 * The backup spreadsheet opens in a new tab when the link in the
 * modal dialog box is clicked.
 */
function backup() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var now = new Date();
  var date = getFormattedDate(now);
  var spreadsheetName = spreadsheet.getName();
  var backupSpreadsheet =
      spreadsheet.copy("Backup - " + spreadsheetName + ' - ' + date);
  var spreadsheetUrl = backupSpreadsheet.getUrl();
  var htmlOutput = HtmlService
                       .createHtmlOutput(
                           '<html><a href="' + spreadsheetUrl +
                           '" target=_blank>Open Backup Spreadsheet</a></html>')
                       .setWidth(350)
                       .setHeight(70);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Backup Copy generated!');
}