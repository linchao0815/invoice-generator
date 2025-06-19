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

    // Template Url 已移除，不再需要

    // Set name spreadsheet
    spreadsheetName: 'Invoice data',

    //Set name document
    documentName: 'Invoice Template',

    // Sheet Settings
    sheetSettings: "Settings",

    // Column Settings
    // The 'col' object is no longer used in the multi-company architecture.
    // Settings are now read dynamically from the 'Settings' sheet based on company name.
    // 不再需要 col 物件，所有設定皆從 Settings 工作表欄位動態讀取
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
 *    - "公司名稱": Company Name (must match the name in the 'Data' sheet)
 *    - "Template URL": The url of the Google Doc template.
 *    - "Folder URL": The url of the Google Drive folder.
 */
function createSystem() {
    try {
        var ss = SpreadsheetApp.getActiveSpreadsheet();
        var sheetSettings = ss.getSheetByName(SETTINGS.sheetSettings);
        var instructionsSheet = ss.getSheetByName('Instructions');
        var urlCell = instructionsSheet.getRange('C15');
        var invoiceFolderUrl = urlCell.getValue();

        // 檢查 C15 是否已經有 Invoice Folder 的 URL
        if (invoiceFolderUrl) {
            // 檢查該 URL 是否有效
            try {
                var folderId = invoiceFolderUrl.match(/[-\w]{25,}/);
                if (!folderId) throw new Error();
                var folder = DriveApp.getFolderById(folderId[0]);
                // 若能正確取得資料夾，直接結束
                showUiDialog('Info', 'Invoice Folder 已存在且有效，無需重複建立。');
                return true;
            } catch (err) {
                showUiDialog('錯誤', 'C15 的 Invoice Folder URL 無效，請手動清空後再執行。');
                return false;
            }
        }

        // 若 C15 為空，檢查根目錄是否有名為 'Invoice Folder' 的資料夾
        var folders = DriveApp.getFoldersByName('Invoice Folder');
        var invoiceFolder;
        if (folders.hasNext()) {
            invoiceFolder = folders.next();
        } else {
            invoiceFolder = DriveApp.createFolder('Invoice Folder');
        }

        // 檢查 'Invoices' 子資料夾是否存在
        var invoicesFolders = invoiceFolder.getFoldersByName('Invoices');
        var invoicesRootFolder;
        if (invoicesFolders.hasNext()) {
            invoicesRootFolder = invoicesFolders.next();
        } else {
            invoicesRootFolder = invoiceFolder.createFolder('Invoices');
        }

        // 更新主資料夾網址到 Instructions
        urlCell.setValue(invoiceFolder.getUrl());

        // 將目前 spreadsheet 搬移到主資料夾
        var file = DriveApp.getFileById(SpreadsheetApp.getActive().getId());
        file.setName(SETTINGS.spreadsheetName);
        moveFile(file, invoiceFolder);

        // 不再複製範本文件

        showUiDialog('Success', 'The main folder structure is ready. You can now add companies to the Settings sheet.');
        return true;
    } catch (e) {
        showUiDialog('Something went wrong', e.message);
    }
}

// 遞迴取得所有 PDF 檔案
function getAllPDFFiles(folder, filesArr) {
    var files = folder.getFiles();
    while (files.hasNext()) {
        var file = files.next();
        if (file.getName().match(/\.pdf$/i)) {
            filesArr.push(file);
        }
    }
    var subfolders = folder.getFolders();
    while (subfolders.hasNext()) {
        getAllPDFFiles(subfolders.next(), filesArr);
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

        // --- Multi-Company Settings Loader ---
        var settingsData = settingsSheet.getDataRange().getValues();
        var settingsMap = {};
        var settingsHeader = settingsData[0];
        var companyNameColIdx = settingsHeader.indexOf("公司名稱");
        var templateIdColIdx = settingsHeader.indexOf("Template URL");
        var folderIdColIdx = settingsHeader.indexOf("Folder URL");
        // 不再處理 System created 欄位

        for (var k = 1; k < settingsData.length; k++) {
            var companyName = settingsData[k][companyNameColIdx];
            if (companyName) {
                // 將 Template URL 與 Folder URL 欄位內容視為 URL，需轉換為 ID
                var templateUrl = settingsData[k][templateIdColIdx];
                var folderUrl = settingsData[k][folderIdColIdx];
                var templateId = templateUrl ? (templateUrl.match(/[-\w]{25,}/) || [null])[0] : null;
                var folderId = folderUrl ? (folderUrl.match(/[-\w]{25,}/) || [null])[0] : null;
                settingsMap[companyName] = {
                    templateId: templateId,
                    folderId: folderId,
                    rowIndex: k + 1 // Store row index to write back the new Folder URL
                };
            }
        }
        // --- End Settings Loader ---

        var sheetValues = dataSheet.getDataRange().getValues();
        var dataHeader = sheetValues[0];
        var pdfIndex = dataHeader.indexOf("PDF Url");
        var clientNameIndex = dataHeader.indexOf("client_name");
        var dataCompanyNameIndex = dataHeader.indexOf("公司名稱");
        var dateIndex = dataHeader.indexOf("invoice date");
        // 不再讀取 Invoice Number 欄位

        if (dataCompanyNameIndex === -1) throw new Error("The 'Data' sheet must contain a '公司名稱' column.");
        if (dateIndex === -1) throw new Error("The 'Data' sheet must contain a 'date' column.");
        // 不再檢查 Invoice Number 欄位

        var key, values, pdfName, invoiceNumber, invoiceDateStr;

        // 取得 Instructions C15 的 Invoice Folder
        var instructionsSheet = ss.getSheetByName('Instructions');
        var invoiceFolderUrl = instructionsSheet.getRange('C15').getValue();
        if (!invoiceFolderUrl) {
            throw new Error('Instructions C15 尚未設定 Invoice Folder URL，請先執行 createSystem()。');
        }
        var folderIdMatch = invoiceFolderUrl.match(/[-\w]{25,}/);
        if (!folderIdMatch) {
            throw new Error('Instructions C15 的 Invoice Folder URL 格式錯誤。');
        }
        var invoiceFolder = DriveApp.getFolderById(folderIdMatch[0]);
        // 檢查/建立 Invoices 子目錄
        var invoicesFolders = invoiceFolder.getFoldersByName('Invoices');
        var rootInvoicesFolder;
        if (invoicesFolders.hasNext()) {
            rootInvoicesFolder = invoicesFolders.next();
        } else {
            rootInvoicesFolder = invoiceFolder.createFolder('Invoices');
        }

        // 掃描所有 PDF 檔名，建立最大流水號表與已用號碼表
        var allPDFFiles = [];
        getAllPDFFiles(rootInvoicesFolder, allPDFFiles);
        var maxCounterMap = {};
        for (var idx = 0; idx < allPDFFiles.length; idx++) {
            var file = allPDFFiles[idx];
            var name = file.getName();
            var match = name.match(/(\d{8}\d{3})\.pdf$/);
            if (match) {
                var num = match[1];
                // 轉為 yyyymmdd
                var yyyymmdd = num.substring(0, 8);
                var counter = parseInt(num.substring(8), 10);
                if (!isNaN(counter)) {
                    if (!maxCounterMap[yyyymmdd] || counter > maxCounterMap[yyyymmdd]) {
                        maxCounterMap[yyyymmdd] = counter;
                    }
                }
            }
        }

        for (var i = 1; i < sheetValues.length; i++) {
            var rowData = sheetValues[i];
            var currentCompanyName = rowData[dataCompanyNameIndex];
            var companySettings = settingsMap[currentCompanyName];

            if (!rowData[pdfIndex] && companySettings) {
                // 檢查 Template URL 是否已設定
                if (!companySettings.templateId) {
                    showUiDialog("錯誤", "公司「" + currentCompanyName + "」的 Template URL 尚未設定，請先於 Settings 工作表補齊。");
                    continue;
                }

                // --- Dynamic Folder Creation Logic ---
                var targetFolderId = companySettings.folderId;
                if (!targetFolderId) {
                    // 先檢查是否已存在同名公司資料夾
                    var existingFolders = rootInvoicesFolder.getFoldersByName(currentCompanyName);
                    if (existingFolders.hasNext()) {
                        var existingFolder = existingFolders.next();
                        targetFolderId = existingFolder.getId();
                    } else {
                        var newFolder = rootInvoicesFolder.createFolder(currentCompanyName);
                        targetFolderId = newFolder.getId();
                    }
                    // Write the ID（不論新建或已存在）回 Settings
                    // 以網址格式寫回
                    var folder = DriveApp.getFolderById(targetFolderId);
                    settingsSheet.getRange(companySettings.rowIndex, folderIdColIdx + 1).setValue(folder.getUrl());
                    companySettings.folderId = targetFolderId;
                }
                // --- End Dynamic Folder Creation ---

                var invoiceDate = new Date(rowData[dateIndex]);
                if (!invoiceDate || isNaN(invoiceDate)) {
                    showUiDialog("Skipping Row " + (i + 1), "Invalid date found.");
                    continue;
                }
                // 產生 yyyymmdd key
                var yyyymmddKey = Utilities.formatDate(invoiceDate, Session.getScriptTimeZone(), "yyyyMMdd");
                // 取得本次日期的最大流水號
                var maxCounter = maxCounterMap[yyyymmddKey] || 0;
                maxCounter++;
                maxCounterMap[yyyymmddKey] = maxCounter;
                // 產生 yyyymmdd+三位數流水號
                var invoiceNumber = yyyymmddKey + maxCounter.padLeft(3, '0'); // 產生 yyyymmddNNN 格式
                var dateFormatted = Utilities.formatDate(invoiceDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
                console.log("DEBUG: Row", i + 1, "date:", dateFormatted, "invoiceNumber:", invoiceNumber, "yyyymmddKey:", yyyymmddKey, "maxCounterMap:", JSON.stringify(maxCounterMap));

                var docId = companySettings.templateId;
                var invoiceId = DriveApp.getFileById(docId).makeCopy('Template Copy ' + new Date().getTime()).getId();
                var docBody = DocumentApp.openById(invoiceId).getBody();

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

                var pdfInvoice = convertPDF(invoiceId, targetFolderId); // Use the determined Folder URL
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
        throw new Error("Folder URL is missing. Cannot convert PDF.");
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
        console.log("DEBUG: showUiDialog called with title:", title, "message:", message);
        var ui = SpreadsheetApp.getUi()
        ui.alert(title, message, ui.ButtonSet.OK)
    } catch (e) {
        // pass
    }
}