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

/**
* Method alert(prompt) of Class Ui is used.
* Opens a dialog box in the user's editor with the given message and an "OK" button.
*/
function menuItem1() {
  SpreadsheetApp.getUi() 
     .alert('You clicked the first menu item!');
}

/**
* Method alert(prompt, buttons) of Class Ui is used.
* Opens a dialog box in the user's editor with the given message, and 'YES' and 'NO' set of buttons.
* The user can also close the dialog by clicking the close button in its title bar.
*/
function menuItem2(){
  var ui = SpreadsheetApp.getUi();  
  var response = ui.alert('Are you sure you want to continue?', ui.ButtonSet.YES_NO);
  // Process the user's response.
  if (response == ui.Button.YES) {
    ui.alert('The user clicked "Yes."');
  } else {
    ui.alert('The user clicked "No" or the close button in the dialog\'s title bar.');
  }
}

function menuItem3() {
  // Add functionality to be performed on clicking the menu item here.
}
