/*================================================================================================================*
Invoice Generator
================================================================================================================
Version:      1.0.0
Project Page: https://github.com/Sheetgo/invoice-generator
Copyright:    (c) 2018 by Sheetgo

License:      GNU General Public License, version 3 (GPL-3.0)
http://www.opensource.org/licenses/gpl-3.0.html
----------------------------------------------------------------------------------------------------------------
Changelog:

1.0.0  Initial release
1.1.0  Auto configuration
*================================================================================================================*/

/**
* Project Settings
* @type {JSON}
*/
SETTINGS = {

    // Spreadsheet name
    sheetName: "Data",

    // Document Url
    documentUrl: null,

    // Template Url
    templateUrl: '14oTfL_zUbBdRD4VXY8U0NAJjQ4cKNxHGBax-bfH5NDs',

    // Set name spreadsheet
    spreadsheetName: 'Invoice data',

    //Set name document
    documentName: 'Invoice Template',

    // Sheet Settings
    sheetSettings: "Settings",

    // Column Settings
    // The 'col' object is no longer used in the multi-company architecture.
    // Settings are now read dynamically from the 'Settings' sheet based on company name.
    col: {}
};

/**
* This funcion will run when you open the spreadsheet. It creates a Spreadsheet menu option to run the spript
*/
function onOpen() {

    // Adds a custom menu to the spreadsheet.
    SpreadsheetApp.getUi()
        .createMenu('Invoice Generator')
        .addItem('Generate Invoices', 'sendInvoice')
        .addToUi();
}

/**
 * [DEPRECATED] This function was used to create a single-company system.
 * In the new multi-company architecture, this function is no longer compatible.
 *
 * To add a new company:
 * 1. Manually create a Google Drive folder for the new company's invoices.
 * 2. Manually create a Google Doc template for the new company.
 * 3. Open the 'Settings' sheet.
 * 4. Add a new row with the following information:
 *    - Column A: Company Name (must match the name in the 'Data' sheet)
 *    - Column B: The ID of the Google Doc template.
 *    - Column C: The ID of the Google Drive folder.
 *    - Column D: Initial invoice count (usually 0).
 *    - Column E: Set to 'TRUE'.
 */
function createSystem() {

    try {

        var ss = SpreadsheetApp.getActiveSpreadsheet();


        // Get name tab
        var sheetSettings = ss.getSheetByName(SETTINGS.sheetSettings);

        // Checks function createSystem is run
        var systemCreated = sheetSettings.getRange(SETTINGS.col.systemCreated);
        if (!systemCreated.getValue()) {
            systemCreated.setValue('True');
        } else {
            showUiDialog('Warnning', 'Solution has already been created!');
            return;
        }

        // Checks if cell Count exists
        var count = sheetSettings.getRange(SETTINGS.col.count);
        if (!count.getValue()) {
            count.setValue(0);
        }

        // Create the Solution folder on users Drive
        var invoiceFolder = DriveApp.createFolder('Invoice Folder');
        var folder = invoiceFolder.createFolder('Invoices');

        // Set URL Invoice Folder in tab Instructions
        ss.getSheetByName('Instructions').getRange('C15').setValue(invoiceFolder.getUrl());

        // Move the current Dashboard spreadsheet into the Solution folder
        var file = DriveApp.getFileById(SpreadsheetApp.getActive().getId());
        file.setName(SETTINGS.spreadsheetName);

        // Move the sheet for invoice folder
        moveFile(file, invoiceFolder);

        // Move the current Dashboard template into the Solution folder
        var doc = DriveApp.getFileById(SETTINGS.templateUrl);
        var docCopy = doc.makeCopy(SETTINGS.documentName);

        // Set tab settings document ID
        sheetSettings.getRange(SETTINGS.col.templateId).setValue(docCopy.getId());

        // Move an copy for invoice folder
        moveFile(docCopy, invoiceFolder);

        // Set folder ID 
        sheetSettings.getRange(SETTINGS.col.folderId).setValue(folder.getId());


        // End process
        showUiDialog('Success', 'Your solution is ready');

        return true;
    } catch (e) {

        // Show the error
        showUiDialog('Something went wrong', e.message)

    }
}


/**
* Reads the spreadsheet data and creates the PDF invoice
*/
function sendInvoice() {
    try {
        var ss = SpreadsheetApp.getActiveSpreadsheet();
        var dataSheet = ss.getSheetByName(SETTINGS.sheetName);
        var settingsSheet = ss.getSheetByName(SETTINGS.sheetSettings);

        // --- Simplified Multi-Company Settings Loader ---
        var settingsData = settingsSheet.getDataRange().getValues();
        var settingsMap = {};
        var settingsHeader = settingsData[0];
        var companyNameColIdx = settingsHeader.indexOf("公司名稱");
        var templateIdColIdx = settingsHeader.indexOf("Template ID");
        var folderIdColIdx = settingsHeader.indexOf("Folder ID");
        var systemCreatedColIdx = settingsHeader.indexOf("System created");

        for (var k = 1; k < settingsData.length; k++) {
            var companyName = settingsData[k][companyNameColIdx];
            if (companyName) {
                settingsMap[companyName] = {
                    templateId: settingsData[k][templateIdColIdx],
                    folderId: settingsData[k][folderIdColIdx],
                    systemCreated: settingsData[k][systemCreatedColIdx]
                };
            }
        }
        // --- End Settings Loader ---

        var sheetValues = dataSheet.getDataRange().getValues();
        var dataHeader = sheetValues[0];
        var pdfIndex = dataHeader.indexOf("PDF Url");
        var clientNameIndex = dataHeader.indexOf("client_name");
        var dataCompanyNameIndex = dataHeader.indexOf("公司名稱");
        var dateIndex = dataHeader.indexOf("date");
        var invoiceNumIndex = dataHeader.indexOf("Invoice Number");

        if (dataCompanyNameIndex === -1) throw new Error("The 'Data' sheet must contain a '公司名稱' column.");
        if (dateIndex === -1) throw new Error("The 'Data' sheet must contain a 'date' column.");
        if (invoiceNumIndex === -1) throw new Error("The 'Data' sheet must contain an 'Invoice Number' column.");

        var key, values, pdfName, invoiceNumber, invoiceDateStr;

        for (var i = 1; i < sheetValues.length; i++) {
            var rowData = sheetValues[i];
            var currentCompanyName = rowData[dataCompanyNameIndex];
            var companySettings = settingsMap[currentCompanyName];

            // Skip if PDF URL already exists, or if company settings are invalid
            if (!rowData[pdfIndex] && companySettings && companySettings.systemCreated === true) {

                // --- New Robust Invoice Number Logic ---
                var invoiceDate = new Date(rowData[dateIndex]);
                if (!invoiceDate || isNaN(invoiceDate)) {
                    showUiDialog("Skipping Row " + (i + 1), "Invalid date found.");
                    continue; // Skip row if date is invalid
                }
                var dateFormatted = Utilities.formatDate(invoiceDate, Session.getScriptTimeZone(), "ddMMyyyy");
                
                var dailyCounter = 1;
                // Scan all existing invoice numbers for the same date to find the next sequence
                for (var r = 1; r < sheetValues.length; r++) {
                    var existingInvoiceNum = sheetValues[r][invoiceNumIndex];
                    if (existingInvoiceNum && existingInvoiceNum.startsWith(dateFormatted)) {
                        var existingCounter = parseInt(existingInvoiceNum.substring(8), 10);
                        if (existingCounter >= dailyCounter) {
                            dailyCounter = existingCounter + 1;
                        }
                    }
                }
                
                invoiceNumber = dateFormatted + dailyCounter.padLeft(3, '0');
                // --- End New Logic ---

                var docId = companySettings.templateId;
                var invoiceId = DriveApp.getFileById(docId).makeCopy('Template Copy ' + new Date().getTime()).getId();
                var docBody = DocumentApp.openById(invoiceId).getBody();

                // Update the generated invoice number in the sheet *before* processing
                // This prevents race conditions and ensures the number is reserved
                dataSheet.getRange(i + 1, invoiceNumIndex + 1).setValue(invoiceNumber);
                rowData[invoiceNumIndex] = invoiceNumber; // Update in-memory array as well

                for (var j = 0; j < rowData.length; j++) {
                    key = dataHeader[j].toString();
                    values = rowData[j];

                    if (key.indexOf("date") > -1 && values) {
                        invoiceDateStr = (values.getMonth() + 1) + "/" + values.getDate() + "/" + values.getFullYear();
                        replace(`%${key}%`, invoiceDateStr, docBody);
                    } else if (values) {
                        if (key.indexOf("price") > -1 || key === "discount" || key.indexOf("total") > -1) {
                            replace(`%${key}%`, `$${values.toFixed(2)}`, docBody);
                        } else if (key === "tax_id") {
                            replace(`%${key}%`, `Tax ID: ${values}`, docBody);
                        } else {
                            replace(`%${key}%`, values, docBody);
                        }
                    } else {
                        replace(`%${key}%`, '', docBody);
                    }
                }
                
                replace('%invoice%', invoiceNumber, docBody);

                pdfName = rowData[clientNameIndex] + " " + invoiceNumber;
                DocumentApp.openById(invoiceId).setName(pdfName).saveAndClose();

                var pdfInvoice = convertPDF(invoiceId, companySettings.folderId);
                dataSheet.getRange(i + 1, pdfIndex + 1).setValue(pdfInvoice[0]);

                Drive.Files.remove(invoiceId);
            }
        }
        showUiDialog('Success', 'Invoice generation process completed.');
    } catch (e) {
        showUiDialog('Something went wrong', e.message + " (Script line: " + e.lineNumber + ")");
    }
}

/**
* Move a file from one folder into another
* @param {Object} file A file object in Google Drive
* @param {Object} dest_folder A folder object in Google Drive 
*/
function moveFile(file, dest_folder, isFolder) {

    if (isFolder === true) {
        dest_folder.addFolder(file)
    } else {
        dest_folder.addFile(file);
    }
    var parents = file.getParents();
    while (parents.hasNext()) {
        var folder = parents.next();
        if (folder.getId() != dest_folder.getId()) {
            if (isFolder === true) {
                folder.removeFolder(file)
            } else {
                folder.removeFile(file)
            }

        }
    }
}

/**
* Convert a Google Docs into a PDF file
* @param {string} id - File Id
* @returns {*[]}
*/
function convertPDF(id, folderId) {
    if (!folderId) {
        throw new Error("Folder ID is missing. Cannot convert PDF.");
    }
    var doc = DocumentApp.openById(id);
    var docBlob = doc.getAs('application/pdf');
    docBlob.setName(doc.getName() + ".pdf"); // Add the PDF extension
    var file = DriveApp.getFolderById(folderId).createFile(docBlob);
    var url = file.getUrl();
    var fileId = file.getId();
    return [url, fileId];
}

/**
* Replace the document key/value
* @param {String} key - The document key to be replaced
* @param {String} text - The document text to be inserted
* @param {Body} body - the active document's Body.
* @returns {Element}
*/
function replace(key, text, body) {
    return body.editAsText().replaceText(key, text);
}


/**
* Returns a new string that right-aligns the characters in this instance by padding them with any string on the left,
* for a specified total length.
* @param {Number} n - Number of characters to pad
* @param {String} str - The string to be padded
* @returns {string}
*/
Number.prototype.padLeft = function (n, str) {
    return Array(n - String(this).length + 1).join(str || '0') + this;
};

/**
* Loads the showDialog
*/
function showDialog() {
    var html = HtmlService.createHtmlOutputFromFile('iframe.html')
        .setWidth(200)
        .setHeight(150)
    SpreadsheetApp.getUi().showModalDialog(html, 'Creating Solution..')
}

/**
* Show an UI dialog
* @param {string} title - Dialog title
* @param {string} message - Dialog message
*/
function showUiDialog(title, message) {
    try {
        var ui = SpreadsheetApp.getUi()
        ui.alert(title, message, ui.ButtonSet.OK)
    } catch (e) {
        // pass
    }
}