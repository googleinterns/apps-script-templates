/**
* Custom Sidebar Template with Input TextBox, Drop-Down List, 
* Radio Buttons, CheckBox Buttons functionalities.
*
* The onOpen event handler is triggered when the attached spreadsheet is opened.
* Class SpreadsheetApp is used to access and create Google Sheets files.
*/
function onOpen() {
  // The getUi method returns an instance of the spreadsheet's UI environment. 
  // It allows the script to add features like menus, dialogs, sidebars.
  SpreadsheetApp.getUi()  
      .createMenu('Custom Menu') // Creates a builder that can be used to add a menu to the editor's UI.
      .addItem('Show Sidebar', 'showSidebar') // Adds an item to the menu.
      .addToUi(); // Inserts the menu into the instance of the editor's user interface.
}

/** 
* Displays an HTML service user interface inside Sheets editor using a sidebar.
* Sidebars do not suspend the server-side script while the dialog is open.
*/
function showSidebar() {
  // Creates a new HtmlOutput object from a file in the code editor.
  var html = HtmlService.createHtmlOutputFromFile('SidebarPage') 
      .setTitle('Present Sidebar')
      .setWidth(300);
  SpreadsheetApp.getUi() 
      .showSidebar(html); // Opens sidebar in the user's editor with custom client-side content
}

function textFunction(inputText) {
  // Add functionality to be performed with the input text.
}

function listFunction(selectedListElement) {
  // Add functionality to be performed with the selected list value.
}

function radioButtonFunction(selectedRadioButton) {
  // Add functionality to be performed with the selected radio button value
}

function checkBoxFunction(selectedCheckboxes) {
  // Add functionality to be performed with the array of selected checkboxes 
}


