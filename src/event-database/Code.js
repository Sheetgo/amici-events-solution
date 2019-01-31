/**
 * Copyright (c) 2018 Sheetgo Europe, S.L.
 *
 * This source code is licensed under the MIT License (MIT) found in the LICENSE file in the
 * root directory of this source tree or on: https://opensource.org/licenses/MIT
 *
 *
 * @link https://github.com/Sheetgo/amici-events-solution
 * @version 1.0.0
 * @licence MIT
 *
 */

/**
 * Creates the Topbar menu in the spreadsheet. This function is fired every time a spreadsheet is open
 * @param {object} e - User and spreadsheet basic parameters
 */
function onOpen(e) {

    // Get the context
    var ui = SpreadsheetApp.getUi()

    // Create the menu
    var menu = ui.createMenu('Amici')

    // Check if user is already authorized
    if (e && e.authMode == ScriptApp.AuthMode.LIMITED) {
        menu.addItem('Activate events solution', 'activateTrigger')
    } else {
        menu.addItem('Update events', 'updateEvents')
    }

    // Add the menu to the spreadsheet
    menu.addToUi()
}


/**
 * Runs once per user, only when the user didn't had authorized the script yet
 * Activates the trigger and callback the onOpen function 
 */
function activateTrigger() {
    toggleTrigger()
    onOpen()
}

/**
 * Switch on the trigger that will run on form submit
 */
function toggleTrigger() {
    try {

        // Get the current spreadsheet
        var activeSpreadsheet = SpreadsheetApp.getActive()

        // Create the script trigger
        ScriptApp
            .newTrigger('updateEvents')
            .forSpreadsheet(activeSpreadsheet)
            .onFormSubmit()
            .create()

    } catch (e) {

        // Show the error
        Logger.log('Something went wrong', e.message)
    }
}

/**
 * Updates the events fields of the forms that uses the events database,
 * also generate an unique ID for events without it
 */
function updateEvents() {
    var events = getEvents()

    var eventsOptions = []
    
    events.forEach(function(event, index) {
        if (event['Event ID'] === "") {
            var id = generateID(event['Event Name'], index)
            setValues('Data Entry', [[id]], index + 2, 10)
            event['Event ID'] = id
        }
        eventsOptions.push(event['Event ID'])
    })

    var forms = getSettings()

    forms.forEach(function(form) {
        updateFormField(
            form['Form ID'],
            form['Field Index'] - 1,
            eventsOptions
        )
    })
}

/**
 * Generates an unique ID based on the event name and it's database row number
 * @returns {string}
 */
function generateID(eventName, row) {
    var name = eventName.replace(/\s+/g, '').toUpperCase().substring(0, 5)
    var number = String(row)
    while (number.length < 5) {
        number = '0' + number
    }
    return name + number
}

/**
 * Get the events
 * @returns {object}
 */
function getEvents() {

    // Get all the sheet data
    var data = SpreadsheetApp
        .getActive()
        .getSheetByName('Data Entry')
        .getDataRange()
        .getValues()

    // Parse the data to a json array
    return parseDataToJsonArray(data)
}

/**
 * Get the solution settings
 * @returns {object}
 */
function getSettings() {

    // Get all the sheet data
    var data = SpreadsheetApp
        .getActive()
        .getSheetByName('Settings')
        .getDataRange()
        .getValues()

    // Parse the data to a json array
    return parseDataToJsonArray(data)
}

/**
 * Get json formatted from data
 * @returns {object}
 */
function parseDataToJsonArray(data) {

    var obj = {}
    var result = []
    var headers = data[0]
    var cols = headers.length
    var row = []

    for (var i = 1, l = data.length; i < l; i++) {

        // Get a row to fill the object
        row = data[i]

        // Clear object
        obj = {}

        for (var col = 0; col < cols; col++) {
            // Fill object with new values
            obj[headers[col]] = row[col]
        }

        // Add object in a final result
        result.push(obj)
    }

    return result
}

/**
 * Updates the form dropdown field with the input data
 * @param {string} formId - The if of the form
 * @param {string} index - Index of the field to be updated
 * @param {Array} eventsOptions - List of events
 */
function updateFormField(formId, index, eventsOptions) {
    var form = FormApp.openById(formId)
    var cb = form.getItems()[index]
    var listItem = cb.asListItem()
    var choices = listItem.getChoices()
    choices.length = 0
    eventsOptions.forEach(function(option) {
        choices.push(listItem.createChoice(option))
    })
    listItem.setChoices(choices)
}

/**
 * Set values into spreadsheet
 * @param {string} sheetName - Name of the sheet that will receive the data
 * @param {array[]} data - The data to save in matrix format
 * @param {int} firstRow - The starting range row cell
 * @param {int} firstColumn - The starting range column cell
 * @param {boolean} [clearData] - If has to clear the old data
 * @param {boolean} [keepData] - If true the new data will be appended below the previous data
 */
function setValues(sheetName, data, firstRow, firstColumn, clearData, keepData) {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName)
    var lastRow = sheet.getLastRow()
    if (clearData && lastRow > firstRow) {
        var range = sheet.getRange(firstRow, firstColumn, lastRow - 1, data[0].length)
        range.clearContent()
    }
    if (keepData) {
        firstRow = lastRow + 1
    }
    var newRange = sheet.getRange(firstRow, firstColumn, data.length, data[0].length)
    newRange.setValues(data)
}
