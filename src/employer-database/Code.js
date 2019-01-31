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

    // Add the option
    menu.addItem('Update forms data', 'updateEmployerForm')

    // Add the menu to the spreadsheet
    menu.addToUi()
}

/**
 * Updates the fields of the forms that uses the employers database
 */
function updateEmployerForm() {

    var employers = getEmployers()

    var forms = getSettings()

    forms.forEach(function(form) {
        updateFormField(
            form['Form ID'],
            form['Field Index'] - 1,
            form['Type'],
            employers
        )
    })
}

/**
 * Get the solution employers
 * @returns {object}
 */
function getEmployers() {

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
 * @param {string} type - Type of employer to register on the field
 * @param {Array} employers - List of employers data
 */
function updateFormField(formId, index, type, employers) {
    var form = FormApp.openById(formId)
    var cb = form.getItems()[index]
    var listItem = cb.asListItem()
    var choices = listItem.getChoices()
    choices.length = 0
    employers.filter(function(employer) {
        return employer[type] === true
    }).forEach(function(employer) {
        choices.push(listItem.createChoice(employer['Code']))
    })
    listItem.setChoices(choices)
}
