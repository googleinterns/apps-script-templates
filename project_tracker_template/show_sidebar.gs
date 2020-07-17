/**
 * Displays sidebar in the spreadsheet
 */
function showSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('sidebar_page')
                 .setTitle('Add')
                 .setWidth(300);
  SpreadsheetApp.getUi().showSidebar(html);
}