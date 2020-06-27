/**
 * Makes a backup copy of the current spreadsheet.
 * Names it appropriately and displays the backup copy's URL in an alert box.
 */
function backup() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var now = new Date();
  var spreadsheetName = spreadsheet.getName();
  var backupSpreadsheet =
      spreadsheet.copy("Backup - " + spreadsheetName + ' - ' + now);
  var spreadsheetUrl = backupSpreadsheet.getUrl();
  var message = "Backup copy generated. Access at - " + spreadsheetUrl;
  displayAlertMessage(message);
}