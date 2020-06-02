/**
 * The event handler triggered when the attached spreadsheet is opened.
 * @param {Event} e The onOpen event object.
 * e.authMode A value from the ScriptApp.AuthMode enum.
 * e.source A Spreadsheet object, representing the Google Sheets file to which the script is bound.
 * e.triggerUid ID of trigger that produced this event (installable triggers only).
 * e.user User object representing the active user, if available.
 * @return {SpreadsheetTriggerBuilder}  Builder for spreadsheet trigger chaining.
 */
function onOpen(e){
}

/**
 * The event handler triggered when an edit is made to the attached spreadsheet.
 * @param {Event} e The onEdit event object.
 * e.authMode A value from the ScriptApp.AuthMode enum.
 * e.oldValue Cell value prior to the edit, if any.
 * e.range A Range object, representing the cell or range of cells that were edited.
 * e.source A Spreadsheet object, representing the Google Sheets file to which the script is bound.
 * e.triggerUid ID of trigger that produced this event (installable triggers only).
 * e.user A User object, representing the active user, if available.
 * e.value New cell value after the edit. Only available if the edited range is a single cell. 
 * @return {SpreadsheetTriggerBuilder} Builder for spreadsheet trigger chaining.
 */
function onEdit(e) {
}

/**
 * The event handler triggered when a response is submitted to the form linked to the attached spreadsheet.
 * @param {Event} e The onFormSubmit event object.
 * e.authMode A value from the ScriptApp.AuthMode enum.
 * e.namedValues An object containing the question names and values from the form submission.
 * e.range A Range object, representing the cell or range of cells that were edited.
 * e.triggerUid ID of trigger that produced this event.
 * e.values Array with values in the same order as they appear in the spreadsheet.
 * @return {SpreadsheetTriggerBuilder} Builder for spreadsheet trigger chaining.
 */
function onFormSubmit(e) {
}

/**
 * The event handler triggered when the attached spreadsheet's content or structure is changed.
 * @param {Event} e The onChange event object.
 * e.authMode A value from the ScriptApp.AuthMode enum.
 * e.changeType The type of change 
 * e.triggerUid ID of trigger that produced this event.
 * e.user A User object, representing the active user, if available 
 * @return {SpreadsheetTriggerBuilder} Builder for spreadsheet trigger chaining.
 */
function onChange(e) {
}

/**
 * The event handler triggered when the selection changes in the spreadsheet. 
 * For example, when a cell/range of cells different from the currently active cell(s) is selected, 
 * or when a sheet different from the currently active sheet is selected.
 * @param {Event} e The onSelectionChange event object.
 * e.authMode A value from the ScriptApp.AuthMode enum.
 * e.source A Spreadsheet object, representing the Google Sheets file to which the script is bound.
 * e.range A Range object, representing the cell or range of cells that were selected.
 * @return {SpreadsheetTriggerBuilder} Builder for spreadsheet trigger chaining.
 */
function onSelectionChange(e) {
}

/**
 * The event handler triggered when an add-on is installed.
 * @param {Event} e The onInstall event object.
 * e.authMode A value from the ScriptApp.AuthMode enum.
 */
function onInstall(e) {
}
