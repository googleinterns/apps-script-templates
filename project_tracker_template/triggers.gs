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
