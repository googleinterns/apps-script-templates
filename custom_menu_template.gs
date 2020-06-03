/**
* A custom menu template for the following - 
* Simple Custom Menu
* Custom Menu with separators
* Custom Menu with sub-menus
* Adding an item to the editor's Add-on menu
*
* The onOpen event handler is triggered when the attached spreadsheet is opened.
* Class SpreadsheetApp is used to access and create Google Sheets files.
*/
function onOpen() {
  // The getUi method returns an instance of the spreadsheet's UI environment.
  // It allows the script to add features like menus, dialogs, sidebars.
  var spreadsheetUi = SpreadsheetApp.getUi(); 
  
  // A simple Custom Menu 
  spreadsheetUi.createMenu('Custom Menu 1')
      .addItem('First item', 'menuItem1') // Adds an item to the menu.
      .addItem('Second item', 'menuItem2')
      .addToUi();  // Inserts the menu into the instance of the editor's user interface.
 
  // Custom Menu with separators
  spreadsheetUi.createMenu('Custom Menu 2')
      .addItem('First item', 'menuItem1')
      .addSeparator() // Adds a visual separator to the menu.
      .addItem('Second item', 'menuItem2')
      .addSeparator() 
      .addItem('Third item', 'menuItem3')
      .addToUi();  
  
  // Custom Menu with a sub-menu
  spreadsheetUi.createMenu('Custom Menu 3')
      .addSubMenu(spreadsheetUi.createMenu('Sub-menu') // Adds a sub-menu to the menu.
          .addItem('First item', 'menuItem1'))
      .addSeparator() 
      .addItem('Second item', 'menuItem2')
      .addSeparator() 
      .addItem('Third item', 'menuItem3')
      .addToUi();  
    
  // Add an item to the editor's Add-on menu, under a sub-menu whose name is set automatically.
  spreadsheetUi.createAddonMenu()
      .addItem('First Item', 'menuItem1')
      .addToUi();
}

function menuItem1() {
  // Add functionality to be performed on clicking the menu item here.
}

function menuItem2(){
  // Add functionality to be performed on clicking the menu item here.
}

function menuItem3() {
  // Add functionality to be performed on clicking the menu item here.
}
