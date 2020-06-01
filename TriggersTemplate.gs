/**
 * The event handler triggered when opening the spreadsheet.
 * @param {Event} e The onOpen event.
 * e.authMode A value from the ScriptApp.AuthMode enum.
 * e.source A Spreadsheet object, representing the Google Sheets file to which the script is bound.
 * e.triggerUid ID of trigger that produced this event (installable triggers only).
 * e.user User object representing the active user, if available.
 * @return {SpreadsheetTriggerBuilder}  Builder for spreadsheet trigger chaining.
 */
function onOpen(e){
}

/**
 * The event handler triggered when editing the spreadsheet.
 * @param {Event} e The onEdit event.
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
 * The event handler triggered when the spreadsheet has a form submitted to it.
 * @param {Event} e The onFormSubmit event.
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
 * The event handler triggered when the spreadsheet's content or structure is changed.
 * @param {Event} e The onChange event.
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
 * @param {Event} e The onSelectionChange event.
 * e.authMode A value from the ScriptApp.AuthMode enum.
 * e.source A Spreadsheet object, representing the Google Sheets file to which the script is bound.
 * e.range A Range object, representing the cell or range of cells that were selected.
 */
function onSelectionChange(e) {
}

/**
 * The event handler triggered when installing the add-on.
 * @param {Event} e The onInstall event.
 * e.authMode A value from the ScriptApp.AuthMode enum.
 */
function onInstall(e) {
}
