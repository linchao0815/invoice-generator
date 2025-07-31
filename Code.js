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
1.2.0  linchao: 修改規格可以支援多個樣版,settings修改規格
1.3.0  linchao: 新增寄信功能可以使用google doc做html 樣版
1.4.0  linchao: 修正發票號碼改以"公司名稱"目錄下檔案,而不是Invoice Folder全部檔案
1.4.1  linchao: 修正log:"app"->"App"
1.4.2  linchao: 補上log缺少欄位,新増"invoice_num"欄位
1.4.3  linchao: 檢查開立發票時，工作表名稱 yyyy/mm 格式，例如 2025/07 且目前時間大於工作表名稱指定的年/月,超過"關帳期限"。
*================================================================================================================*/
let logUrl = "https://script.google.com/macros/s/AKfycbykmOscH010Putq3c8dhCYaAxxOCrLIqTfz8K50ZQTROcbWWdNgtX4Ux3aNDTo2FBxU/exec";
function ElkLog(msg) {
    let userName = "",domain="";
    try {
        let acnt = Session.getActiveUser().getEmail();
        if (acnt.indexOf('@') > -1) {
            domain = acnt.split('@')[1];
            userName = acnt.split('@')[0];
        }
    } catch (e) { }
    let payload = Object.assign({ "App": "invoice", "Domain": domain }, msg);
    payload["UserName"]=userName;
    // 遍歷 val 物件的所有屬性
    for (var key in msg) {
        // 檢查屬性名稱是否包含 'date' (不區分大小寫)
        if (key.toLowerCase().includes('date')) {
            // 嘗試將屬性值轉換為 Date 物件並格式化為 ISO 字串
            try {
                // 如果值是有效的日期字串或數字，則轉換
                if (msg[key] && !isNaN(new Date(msg[key]).getTime())) {
                    payload[key] = new Date(msg[key]).toISOString();
                }
            } catch (e) {
                // 如果轉換失敗，則在日誌中記錄錯誤，但不要中斷執行
                Logger.log('Could not convert key "' + key + '" with value "' + msg[key] + '" to ISOString: ' + e.toString());
            }
        }
    }
    var options = {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
    };
    var resp = UrlFetchApp.fetch(logUrl, options);
    Logger.log(`payload: ${JSON.stringify(payload)}, resp: ${resp}`);
}

var msg = {
    app: 'invoice',
    UserName: 'linchao0815@gmail.com',
    client_email: 'linchao@igs.com.tw,linchao.chang@gmail.com',
    client_name: 'Alpha Games',
    '客戶平台': 'Brazino777/ Admiral',
    '公司名稱': 'FaDa',
    client_address: 'Capitao Antonio Rosa Street, No. 409 - Jardim Paulistano Neighborhood - Sao Paulo/Sao Paulo\r\n01.443-010\r',
    DESCRIPTION: '《License Fee》for Jun, 2025',
    '幣別': 'EUR',
    total: 123131.23,
    'PDF Url': 'https://drive.google.com/file/d/14jElMsj7Q36aKvo1iBhEHrO5sf9C9dTb/view?usp=drivesdk',
    'Attachment Url': '',
    'Email Sent Status': '2025-06-23 16:34:32',
    'invoice date': '2025-06-23'
}

function testLog() {
    ElkLog(msg)
}
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
    pdfColName: 'PDF Url',
    pdfFileHead: '客戶平台',
    attFileColName: 'Attachment Url',
    emailRecipientColName: 'client_email',
    emailSubjectColName: 'email_subject',
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
        .addItem('產生發票', 'generateInvoice')
        .addItem('Send Emails', 'sendEmail')
        .addItem('產生並寄送發票', 'generateAndSendInvoice')
        .addToUi();
}

// 合併產生與寄送發票
function generateAndSendInvoice() {
    generateInvoice(false); // 不顯示對話框
    sendEmail();
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
        showUiDialog('Something went wrong'+ e.message+" Stack Trace: " + e.stack);
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

// 遞迴取得所有 PDF 檔案，依公司名稱分組
function getAllPDFFilesByCompany(folder, filesMap) {
    var folderName = folder.getName();
    var files = folder.getFiles();
    while (files.hasNext()) {
        var file = files.next();
        if (file.getName().match(/\.pdf$/i)) {
            if (!filesMap[folderName]) filesMap[folderName] = [];
            filesMap[folderName].push(file);
        }
    }
    var subfolders = folder.getFolders();
    while (subfolders.hasNext()) {
        getAllPDFFilesByCompany(subfolders.next(), filesMap);
    }
}

/**
* Reads the spreadsheet data and creates the PDF invoice
*/
function generateInvoice(bShowDialog = true) {
    try {
        var ss = SpreadsheetApp.getActiveSpreadsheet();
        var dataSheet = ss.getActiveSheet(); // 取得目前 active sheet
        var sheetName = dataSheet.getName();
        // 檢查名稱格式 yyyy/mm
        if (!/^\d{4}\/\d{2}$/.test(sheetName)) {
            throw new Error("目前工作表名稱必須為 yyyy/mm 格式，例如 2025/07。");
        }
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
        var pdfIndex = dataHeader.indexOf(SETTINGS.pdfColName);
        var pdfHeadNameIndex = dataHeader.indexOf(SETTINGS.pdfFileHead);
        var dataCompanyNameIndex = dataHeader.indexOf("公司名稱");
        var dateIndex = dataHeader.indexOf("invoice date");
        var invoice_numIndex = dataHeader.indexOf("invoice_num");
        
        if (dataCompanyNameIndex === -1) throw new Error("The 'Data' sheet must contain a '公司名稱' column.");
        if (dateIndex === -1) throw new Error("The 'Data' sheet must contain a 'date' column.");
        // 檢查 Invoice Number 欄位
        if (invoice_numIndex === -1) throw new Error("The 'Data' sheet must contain a 'invoice_num' column.");

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

        // 掃描所有公司 PDF 檔案，依公司名稱分組
        var allPDFFilesMap = {};
        getAllPDFFilesByCompany(rootInvoicesFolder, allPDFFilesMap);
        // maxCounterMap: { [companyName]: { [yyyymmdd]: maxCounter } }
        var maxCounterMap = {};
        Object.keys(allPDFFilesMap).forEach(function(companyName) {
            var files = allPDFFilesMap[companyName];
            maxCounterMap[companyName] = {};
            for (var idx = 0; idx < files.length; idx++) {
                var file = files[idx];
                var name = file.getName();
                var match = name.match(/(\d{8}\d{3})\.pdf$/);
                if (match) {
                    var num = match[1];
                    var yyyymmdd = num.substring(0, 8);
                    var counter = parseInt(num.substring(8), 10);
                    if (!isNaN(counter)) {
                        if (!maxCounterMap[companyName][yyyymmdd] || counter > maxCounterMap[companyName][yyyymmdd]) {
                            maxCounterMap[companyName][yyyymmdd] = counter;
                        }
                    }
                }
            }
        });

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
                // 檢查 invoiceDate 的年/月是否與 sheetName 相符
                var invoiceYear = invoiceDate.getFullYear();
                var invoiceMonth = invoiceDate.getMonth() + 1; // getMonth() 從 0 開始
                var sheetYearMonth = sheetName.split('/');
                if (sheetYearMonth.length === 2) {
                    var now = new Date();
                    var currentYear = now.getFullYear();
                    var currentMonth = now.getMonth() + 1;                
                    var sheetYear = parseInt(sheetYearMonth[0], 10);
                    var sheetMonth = parseInt(sheetYearMonth[1], 10);
                    if (invoiceYear !== sheetYear || invoiceMonth !== sheetMonth) {
                        throw new Error(`Skipping Row ${i + 1} "invoice date":${rowData[dateIndex]} 年/月 與工作表名稱不符。`);
                    }
                    // 若目前年/月 > sheet 年/月 也報錯
                    if (currentYear > sheetYear || (currentYear === sheetYear && currentMonth > sheetMonth)) {
                        throw new Error(`Skipping Row ${i + 1} "invoice date":${rowData[dateIndex]} 目前時間大於工作表名稱:sheetName 指定的年/月,超過"關帳期限"。`);
                    }
                }                
                // 產生 yyyymmdd key
                var yyyymmddKey = Utilities.formatDate(invoiceDate, Session.getScriptTimeZone(), "yyyyMMdd");
                // 取得本公司本次日期的最大流水號
                if (!maxCounterMap[currentCompanyName]) maxCounterMap[currentCompanyName] = {};
                var maxCounter = maxCounterMap[currentCompanyName][yyyymmddKey] || 0;
                maxCounter++;
                maxCounterMap[currentCompanyName][yyyymmddKey] = maxCounter;
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
                            replace(`%${key}%`, `${values}`, docBody);
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
                rowData[invoice_numIndex] = invoiceNumber; // 更新 rowData 以便寫入 log
                dataSheet.getRange(i + 1, invoice_numIndex + 1).setValue(invoiceNumber);
                pdfName = rowData[pdfHeadNameIndex] + " " + invoiceNumber;
                DocumentApp.openById(invoiceId).setName(pdfName).saveAndClose();

                var pdfInvoice = convertPDF(invoiceId, targetFolderId); // Use the determined Folder URL
                dataSheet.getRange(i + 1, pdfIndex + 1).setValue(pdfInvoice[0]);
                rowData[pdfIndex] = pdfInvoice[0]; // 更新 rowData 以便寫入 log
                Drive.Files.remove(invoiceId);
                // 印出此列 data 的 JSON 格式 log
                var rowJson = {};
                for (var colIdx = 0; colIdx < dataHeader.length; colIdx++) {
                    rowJson[dataHeader[colIdx]] = rowData[colIdx];
                }
                console.log("generateInvoice row data:", JSON.stringify(rowJson));
                // 寫入 log sheet
                writeLogSheet('generateInvoice', rowJson);
            }
        }
        if (bShowDialog) showUiDialog('Success', 'Invoice generation process completed.');
    } catch (e) {
        // 自訂錯誤（你 throw new Error()）都會是 Error 物件
        if (e instanceof Error && e.stack && e.stack.indexOf('generateInvoice') !== -1) {
            showUiDialog('錯誤', e.message);
        } else {
            showUiDialog('Something went wrong', e.message + "\n" + (e.stack || ""));
        }
    }
}
/**
 * Google Apps Script 版本的 HttpPost
 * @param {string} url - 目標網址
 * @param {string} data - POST 的資料（字串，通常為 JSON 或 NDJSON）
 * @param {Array<string>} headers - HTTP 標頭陣列
 * @param {number} timeout - 逾時（毫秒）
 * @return {Object} - {code: 狀態碼, response: 回應內容}
 */
function HttpPost(url, data, headers, timeout) {
    var options = {
        method: "post",
        contentType: "application/x-ndjson;charset=UTF-8",
        payload: data,
        muteHttpExceptions: true,
        headers: {},
        timeoutSeconds: Math.min(Math.ceil(timeout / 1000), 300)
    };

    if (headers && headers.length) {
        headers.forEach(function (h) {
            var idx = h.indexOf(":");
            if (idx > 0) {
                var key = h.substring(0, idx).trim();
                var val = h.substring(idx + 1).trim();
                options.headers[key] = val;
            }
        });
    }

    var resp = UrlFetchApp.fetch(url, options);
    return {
        code: resp.getResponseCode(),
        response: resp.getContentText()
    };
}

// 寫入 log sheet 的共用函式
function writeLogSheet(source, rowJson) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var logSheet = ss.getSheetByName('log');
    if (!logSheet) {
        logSheet = ss.insertSheet('log');
        var protection = logSheet.protect().setDescription('log sheet 保護').setWarningOnly(false);
        protection.removeEditors(protection.getEditors());
    }
    var now = Utilities.formatDate(new Date(), "Asia/Taipei", "yyyy-MM-dd HH:mm:ss");
    var keys = Object.keys(rowJson);
    // 若 log sheet 無標題，寫入標題
    if (logSheet.getLastRow() === 0 || logSheet.getLastColumn() === 0) {
        logSheet.appendRow(['time', 'source'].concat(keys));
    }
    // 取得標題
    var headers = logSheet.getRange(1, 1, 1, logSheet.getLastColumn()).getValues()[0];
    // 檢查是否有缺少的欄位，若有則自動補上
    var missingCols = keys.filter(function(k) { return headers.indexOf(k) === -1; });
    if (missingCols.length > 0) {
        logSheet.insertColumnsAfter(logSheet.getLastColumn(), missingCols.length);
        for (var i = 0; i < missingCols.length; i++) {
            logSheet.getRange(1, headers.length + 1 + i).setValue(missingCols[i]);
        }
        headers = logSheet.getRange(1, 1, 1, logSheet.getLastColumn()).getValues()[0];
    }
    var row = [now, source];
    for (var i = 2; i < headers.length; i++) {
        row.push(rowJson[headers[i]] || '');
    }
    // 解除保護
    var protections = logSheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
    var protection = protections.length > 0 ? protections[0] : null;
    if (protection) protection.remove();
    logSheet.appendRow(row);
    rowJson['kind'] = source;
    ElkLog(rowJson);
    // 再加回保護
    if (!protection) {
        protection = logSheet.protect().setDescription('log sheet 保護').setWarningOnly(false);
    } else {
        protection = logSheet.protect().setDescription('log sheet 保護').setWarningOnly(false);
    }
    protection.removeEditors(protection.getEditors());
}

// 依據規格新增 sendEmail 功能
function sendEmail() {
    try {
        var ss = SpreadsheetApp.getActiveSpreadsheet();
        var dataSheet = ss.getSheetByName(SETTINGS.sheetName);
        var settingsSheet = ss.getSheetByName(SETTINGS.sheetSettings);

        // 載入 Settings
        var settingsData = settingsSheet.getDataRange().getValues();
        var settingsHeader = settingsData[0];
        var companyNameColIdx = settingsHeader.indexOf("公司名稱");
        var emailTemplateUrlIdx = settingsHeader.indexOf("Email Template URL");
        var emailSubjectIdx = settingsHeader.indexOf(SETTINGS.emailSubjectColName);

        // 建立公司設定 map
        var settingsMap = {};
        for (var k = 1; k < settingsData.length; k++) {
            var companyName = settingsData[k][companyNameColIdx];
            if (companyName) {
                settingsMap[companyName] = {
                    emailTemplateUrl: settingsData[k][emailTemplateUrlIdx],
                    emailSubject: settingsData[k][emailSubjectIdx],
                };
            }
        }

        // 讀取 Data
        var sheetValues = dataSheet.getDataRange().getValues();
        var dataHeader = sheetValues[0];
        var pdfIndex = dataHeader.indexOf(SETTINGS.pdfColName);
        var dataCompanyNameIndex = dataHeader.indexOf("公司名稱");
        var emailSentStatusIndex = dataHeader.indexOf("Email Sent Status");
        // recipient 來源改為 data 欄位 SETTINGS.emailRecipientColName
        var recipientColIndex = dataHeader.indexOf(SETTINGS.emailRecipientColName);
        for (var i = 1; i < sheetValues.length; i++) {
            var rowData = sheetValues[i];
            var currentCompanyName = rowData[dataCompanyNameIndex];
            var companySettings = settingsMap[currentCompanyName];
            if (!companySettings) continue;

            var pdfUrl = rowData[pdfIndex];
            var emailSentStatus = rowData[emailSentStatusIndex];
            if (pdfUrl && !emailSentStatus && recipientColIndex !== -1) {
                var recipientRaw = rowData[recipientColIndex];
                // 支援多位收件人，允許逗號或分號分隔
                var recipient = recipientRaw ? recipientRaw.replace(/;/g, ',') : '';
                var emailTemplateUrl = companySettings.emailTemplateUrl;
                // 主旨來源改為 settingsSheet 的 emailSubject 欄位
                var emailSubjectRaw = companySettings.emailSubject || "Invoice";
                if (!emailTemplateUrl || !recipient) continue;

                // 取得 Email Template 的 ID
                var templateIdMatch = emailTemplateUrl.match(/[-\w]{25,}/);
                if (!templateIdMatch) continue;
                var templateId = templateIdMatch[0];

                // 複製範本並填入資料
                var emailDocId = DriveApp.getFileById(templateId).makeCopy('Email Template Copy ' + new Date().getTime()).getId();
                var docBody = DocumentApp.openById(emailDocId).getBody();

                // 取代佔位符
                for (var j = 0; j < rowData.length; j++) {
                    var key = dataHeader[j].toString();
                    var value = rowData[j];
                    if (value) {
                        replace(`%${key}%`, value, docBody);
                    } else {
                        replace(`%${key}%`, '', docBody);
                    }
                }

                DocumentApp.openById(emailDocId).saveAndClose();

                // 取得 HTML 內容
                var htmlBody = getHtmlFromDoc(emailDocId);

                // 內嵌圖片：將 <img src="https://..."> 轉為 cid 並準備 inlineImages
                var inlineImages = {};
                var cidIndex = 1;
                htmlBody = htmlBody.replace(/<img[^>]+src="([^"]+)"[^>]*>/g, function (match, src) {
                    try {
                        var imgResponse = UrlFetchApp.fetch(src, { headers: { "Authorization": "Bearer " + ScriptApp.getOAuthToken() } });
                        if (imgResponse.getResponseCode() === 200) {
                            var contentType = imgResponse.getHeaders()['Content-Type'] || 'image/png';
                            var cid = "img" + (cidIndex++);
                            inlineImages[cid] = imgResponse.getBlob().setName(cid + "." + contentType.split('/')[1]);
                            return match.replace(src, "cid:" + cid);
                        }
                    } catch (e) { }
                    return match;
                });
                // 移除 Google Docs 轉出 HTML 的全域 padding/margin/max-width
                htmlBody = htmlBody
                    .replace(/(padding|margin|background|max-width)\s*:\s*[^;"]+;?/gi, '')
                    .replace(/<body[^>]*>/i, '<body style="padding:0;margin:0;background:none;max-width:none;">')
                    // 修正 img 標籤
                    .replace(/<img([^>]*)>/gi, function (match, attrs) {
                        var newAttrs = attrs
                            .replace(/\s*border\s*=\s*["'][^"']*["']/gi, '')
                            .replace(/\s*style\s*=\s*["'][^"']*["']/gi, '');
                        return '<img' + newAttrs + ' border="0" style="border:none">';
                    })
                    // 修正 span 標籤
                    .replace(/<span([^>]*)>/gi, function (match, attrs) {
                        var newAttrs = attrs
                            .replace(/\s*border\s*:\s*[^;"]+;?/gi, '')
                            .replace(/\s*border\s*=\s*["'][^"']*["']/gi, '')
                            .replace(/\s*style\s*=\s*["']([^"']*)["']/gi, function (m, style) {
                                // 移除 style 內 border 設定
                                var newStyle = style.replace(/border\s*:\s*[^;"]+;?/gi, '');
                                return newStyle.trim() ? ' style="' + newStyle + '"' : '';
                            });
                        // 強制加上 border:0
                        if (!/style\s*=/.test(newAttrs)) {
                            newAttrs += ' style="border:0;"';
                        } else {
                            newAttrs = newAttrs.replace(/style="([^"]*)"/, function (m, style) {
                                return 'style="' + style.replace(/border\s*:\s*[^;"]+;?/gi, '') + 'border:0;"';
                            });
                        }
                        return '<span' + newAttrs + '>';
                    });

                // 刪除臨時文件
                Drive.Files.remove(emailDocId);

                // subject 也支援 %key% 取代
                // 主旨支援 %key% 取代，與 docBody 一致
                var emailSubject = emailSubjectRaw;
                for (var j = 0; j < rowData.length; j++) {
                    var key = dataHeader[j].toString();
                    var value = rowData[j] || '';
                    emailSubject = emailSubject.replace(new RegExp(`%${key}%`, 'g'), value);
                }

                // 處理附件
                var attachments = [];
                var pdfUrl = rowData[pdfIndex];
                var attFileIndex = dataHeader.indexOf(SETTINGS.attFileColName);
                var attUrl = attFileIndex !== -1 ? rowData[attFileIndex] : null;

                function extractFileId(url) {
                    var match = url ? url.match(/[-\w]{25,}/) : null;
                    return match ? match[0] : null;
                }

                // PDF 附件
                var pdfFileId = extractFileId(pdfUrl);
                if (pdfFileId) {
                    try {
                        var pdfFile = DriveApp.getFileById(pdfFileId);
                        attachments.push(pdfFile.getBlob());
                    } catch (e) { }
                }
                // 其他附件
                var attFileId = extractFileId(attUrl);
                if (attFileId) {
                    try {
                        var attFile = DriveApp.getFileById(attFileId);
                        attachments.push(attFile.getBlob());
                    } catch (e) { }
                }

                // 寄送郵件
                GmailApp.sendEmail(recipient, emailSubject, '', {
                    htmlBody: htmlBody,
                    attachments: attachments,
                    inlineImages: inlineImages
                });

                // 寫入寄送時間
                var now = Utilities.formatDate(new Date(), "Asia/Taipei", "yyyy-MM-dd HH:mm:ss");
                dataSheet.getRange(i + 1, emailSentStatusIndex + 1).setValue(now);
                rowData[emailSentStatusIndex] = now; // 更新 rowData 以便寫入 log
                // 印出此列 data 的 JSON 格式 log
                var rowJson = {};
                for (var colIdx = 0; colIdx < dataHeader.length; colIdx++) {
                    rowJson[dataHeader[colIdx]] = rowData[colIdx];
                }
                console.log("sendEmail row data:", JSON.stringify(rowJson));
                // 寫入 log sheet
                writeLogSheet('sendEmail', rowJson);
            }
        }
        showUiDialog('Success', 'Email sending process completed.');
    } catch (e) {
        showUiDialog('Something went wrong', e.message + " (Script line: " + e.lineNumber + ")");
    }
}

// 取得 Google Doc 轉為 HTML 字串
function getHtmlFromDoc(docId) {
    var url = "https://docs.google.com/feeds/download/documents/export/Export?id=" + docId + "&exportFormat=html";
    var token = ScriptApp.getOAuthToken();
    var response = UrlFetchApp.fetch(url, {
        headers: {
            "Authorization": "Bearer " + token
        },
        muteHttpExceptions: true
    });
    return response.getContentText();
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

    try {
        var doc = DocumentApp.openById(id);
        var docBlob = doc.getAs('application/pdf');
        docBlob.setName(doc.getName() + ".pdf");
        
        // 偵錯：檢查目標資料夾是否存在且可存取
        var folder = DriveApp.getFolderById(folderId);
        Logger.log("Successfully accessed folder: " + folder.getName());

        // 建立檔案
        var file = folder.createFile(docBlob);
        var url = file.getUrl();
        var fileId = file.getId();
        
        Logger.log("File created successfully. URL: " + url);
        return [url, fileId];

    } catch (e) {
        // 捕獲並記錄詳細錯誤
        Logger.log("An error occurred: " + e.toString());
        Logger.log("Error Name: " + e.name);
        Logger.log("Error Message: " + e.message);
        Logger.log("Stack Trace: " + e.stack);
        
        // 將錯誤訊息拋出，讓呼叫方知道失敗了
        throw new Error("Failed to create PDF. " + e.toString());
    }
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