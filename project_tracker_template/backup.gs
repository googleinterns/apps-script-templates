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