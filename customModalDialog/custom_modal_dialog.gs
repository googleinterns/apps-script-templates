/**
* Custom modal dialog box template with Input TextBox, Drop-Down List, 
* Radio Buttons, Checkbox Buttons functionalities.
*
* The onOpen event handler is triggered when the attached spreadsheet is opened.
* Class SpreadsheetApp is used to access and create Google Sheets files.
*/
function onOpen() {
  // The getUi method returns an instance of the spreadsheet's UI environment. 
  // It allows the script to add features like menus, dialogs, sidebars.
  SpreadsheetApp.getUi()  
      .createMenu('Custom Menu') // Creates a builder that can be used to add a menu to the editor's UI.
      .addItem('Show Modal Dialog Box', 'showModalDialog') // Adds an item to the menu.
      .addToUi(); // Inserts the menu into the instance of the editor's user interface.
}

/** 
* Opens a modal dialog box in the user's editor with custom client-side content.
* This method does not suspend the server-side script while the dialog is open.
*/
function showModalDialog() {
  // Creates a new HtmlOutput object from a file in the code editor.
  var htmlOutput = HtmlService.createHtmlOutputFromFile('modal_dialog_page')
      .setWidth(350)
      .setHeight(360);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Modal Dialog Box');
}

/**
* Display the input text, selected list element, selected radio button 
* and selected checkboxes using an alert dialog box.
* 
* Method alert(prompt) of Class Ui is used.
* Opens a dialog box in the user's editor with the given message and an "OK" button.
*/
function displayInput(inputText,selectedListElement,selectedRadioButton,selectedCheckboxes) {
  var sheet = SpreadsheetApp.getActiveSheet();  
  var spreadsheetUi = SpreadsheetApp.getUi();
  // Concatenate and display the user input information in a single message
  var messageInputText = "Text entered: " + inputText;
  var messageListElement = "\nList Element selected: " + selectedListElement;
  var messageRadioButton = "\nRadio Button selected: " + selectedRadioButton;
  var messageCheckbox = "\nCheckboxes selected: ";
  for (var i = 0; i < selectedCheckboxes.length; i++) {
    messageCheckbox += "\nCheckbox " + selectedCheckboxes[i];
  }
  var displayMessage = messageInputText + messageListElement + messageRadioButton + messageCheckbox;
  spreadsheetUi.alert(displayMessage);
}
