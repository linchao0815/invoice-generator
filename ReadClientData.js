/*================================================================================================================*
  ReadClientData - 從 Attachment Url 讀取外部 Google Sheet 資料
  ================================================================================================================
  Version:      1.0.0
  Description:  根據「客戶表」工作表中 TableClient 表格的設定，讀取 Attachment Url 指向的外部檔案
                (支援原生 Google Sheets 及 .xlsx)，搜尋指定欄位的值並回填至目前月份工作表。

  前置需求:
  - 需啟用 Drive Advanced Service (Apps Script 編輯器 → 服務 → Drive API)

  Changelog:
  1.0.0  初始版本：讀取 TableClient 設定，從外部 Sheet/xlsx 查找並回填數值

  TableClient 欄位：
    客戶平台 | client_name | 公司名稱 | client_email | client_address | receive_acnt | 幣別 | WHT/VAT
    sheet_input1 | input1 | sheet_input1|add|1 | input1|add|1 | sheet_input1|add|2 | input1|add|2
    sheet_input1|add|3 | input1|add|3
*================================================================================================================*/

/**
 * 主函式：讀取客戶資料
 * 遍歷當前 yyyy/mm 工作表，對有 Attachment Url 的列，
 * 根據 TableClient 設定從外部 Google Sheet 查找數值並回填。
 */
function readClientData() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var dataSheet = ss.getActiveSheet();
    var sheetName = dataSheet.getName();
    console.log("readClientData START - sheet: '" + sheetName + "'");

    // 檢查名稱格式 yyyy/mm
    if (!/^\d{4}\/\d{2}$/.test(sheetName)) {
      console.log("ERROR: 工作表名稱不符合 yyyy/mm 格式: '" + sheetName + "'");
      showUiDialog("錯誤", "目前工作表名稱必須為 yyyy/mm 格式，例如 2026/02。");
      return;
    }

    // 讀取 TableClient 設定
    var clientMap = getTableClientMap(ss);
    if (!clientMap || Object.keys(clientMap).length === 0) {
      console.log("ERROR: TableClient 表格沒有資料或不存在");
      showUiDialog("錯誤", "TableClient 表格沒有資料或不存在。");
      return;
    }
    console.log("TableClient 載入完成，共 " + Object.keys(clientMap).length + " 筆客戶設定");

    // 讀取當前工作表資料
    var sheetValues = dataSheet.getDataRange().getValues();
    var dataHeader = sheetValues[0];

    var attUrlIndex = dataHeader.indexOf(SETTINGS.attFileColName); // "Attachment Url"
    var clientPlatformIndex = dataHeader.indexOf("客戶平台");
    var input1Index = dataHeader.indexOf("input1");

    if (attUrlIndex === -1) {
      console.log("ERROR: 缺少 'Attachment Url' 欄位 (尋找名稱: '" + SETTINGS.attFileColName + "')");
      showUiDialog("錯誤", "目前工作表缺少 'Attachment Url' 欄位。");
      return;
    }
    if (clientPlatformIndex === -1) {
      console.log("ERROR: 缺少 '客戶平台' 欄位");
      showUiDialog("錯誤", "目前工作表缺少 '客戶平台' 欄位。");
      return;
    }
    if (input1Index === -1) {
      console.log("ERROR: 缺少 'input1' 欄位");
      showUiDialog("錯誤", "目前工作表缺少 'input1' 欄位。");
      return;
    }

    var updatedCount = 0;
    var skippedCount = 0;
    var fileCache = {}; // 快取已開啟的外部檔案，避免重複開檔/轉檔
    console.log("共 " + (sheetValues.length - 1) + " 列資料待處理");

    for (var i = 1; i < sheetValues.length; i++) {
      var rowData = sheetValues[i];
      var attUrl = rowData[attUrlIndex];
      var clientPlatform = rowData[clientPlatformIndex];

      // 跳過沒有 Attachment Url 的列
      if (!attUrl) {
        continue;
      }

      // 從 TableClient 取得此客戶的設定
      var clientConfig = clientMap[clientPlatform];
      if (!clientConfig) {
        console.log("Row " + (i + 1) + " SKIP: 找不到客戶平台 '" + clientPlatform + "'");
        skippedCount++;
        continue;
      }

      // 取得 sheet_input1 和 input1 設定
      var sheetInput1 = clientConfig["sheet_input1"];
      var searchKey1 = clientConfig["input1"];

      if (!sheetInput1 || !searchKey1) {
        console.log("Row " + (i + 1) + " [" + clientPlatform + "] SKIP: 缺少 sheet_input1 或 input1 設定");
        skippedCount++;
        continue;
      }

      // 開啟 Attachment Url 指向的檔案 (支援 Google Sheets 及 .xlsx，帶快取)
      var externalSS;
      var tempFileId = null;
      try {
        var fileId = extractFileIdFromUrl(attUrl);
        if (!fileId) {
          console.log("Row " + (i + 1) + " SKIP: 無法從 Url 取得檔案 ID");
          skippedCount++;
          continue;
        }

        // 查快取，避免重複開檔/轉檔
        if (fileCache[fileId]) {
          externalSS = fileCache[fileId];
        } else {
          var driveFile = DriveApp.getFileById(fileId);
          var mimeType = driveFile.getMimeType();

          if (mimeType === "application/vnd.google-apps.spreadsheet") {
            externalSS = SpreadsheetApp.openById(fileId);
          } else {
            var copiedFile = Drive.Files.copy(
              { title: "temp_readClient_" + fileId, mimeType: "application/vnd.google-apps.spreadsheet" },
              fileId
            );
            tempFileId = copiedFile.id;
            externalSS = SpreadsheetApp.openById(tempFileId);
          }
          fileCache[fileId] = externalSS;
        }
      } catch (e) {
        console.log("Row " + (i + 1) + " ERROR: 無法開啟外部檔案: " + e.message);
        skippedCount++;
        cleanupTempFile_(tempFileId);
        continue;
      }

      // 查找主要值 (sheet_input1 + input1)
      var mainValue = lookupValueInSheet(externalSS, sheetInput1, searchKey1);

      if (mainValue === null) {
        console.log("Row " + (i + 1) + " [" + clientPlatform + "] SKIP: 在 '" + sheetInput1 + "' 找不到 '" + searchKey1 + "'");
        skippedCount++;
        cleanupTempFile_(tempFileId);
        continue;
      }

      var parsed = parseFloat(mainValue);
      var totalValue = isNaN(parsed) ? 0 : parsed;
      var sources = sheetInput1 + "/" + searchKey1 + "=" + totalValue;

      // 處理 add|1 ~ add|3 的額外值
      for (var addIdx = 1; addIdx <= 3; addIdx++) {
        var addSheetKey = "sheet_input1|add|" + addIdx;
        var addInputKey = "input1|add|" + addIdx;

        var addSheetName = clientConfig[addSheetKey];
        var addSearchKey = clientConfig[addInputKey];

        if (!addSheetName || !addSearchKey) {
          continue;
        }

        var addValue = lookupValueInSheet(externalSS, addSheetName, addSearchKey);
        if (addValue !== null) {
          var parsedAdd = parseFloat(addValue);
          var parsedAddValue = isNaN(parsedAdd) ? 0 : parsedAdd;
          totalValue += parsedAddValue;
          sources += " + " + addSheetName + "/" + addSearchKey + "=" + parsedAddValue;
        } else {
          console.log("Row " + (i + 1) + " [" + clientPlatform + "] WARNING: add|" + addIdx + " 在 '" + addSheetName + "' 找不到 '" + addSearchKey + "'");
        }
      }

      // 回填 input1 欄位
      dataSheet.getRange(i + 1, input1Index + 1).setValue(totalValue);
      console.log("Row " + (i + 1) + " [" + clientPlatform + "] OK: " + totalValue + " (" + sources + ")");
      updatedCount++;

      // 清理暫存轉檔
      cleanupTempFile_(tempFileId);
    }

    console.log("readClientData END - 更新: " + updatedCount + ", 跳過: " + skippedCount);
    showUiDialog("完成", "讀取客戶資料完成。\n更新：" + updatedCount + " 列\n跳過：" + skippedCount + " 列");
  } catch (e) {
    console.log("FATAL ERROR: " + e.message);
    console.log("Stack: " + (e.stack || "N/A"));
    showUiDialog("錯誤", e.message + "\n" + (e.stack || ""));
  }
}

/**
 * 讀取 TableClient 表格/範圍，以「客戶平台」為 key 建立設定 map
 * @param {Spreadsheet} ss - 當前試算表
 * @returns {Object} - { "客戶平台名稱": { "sheet_input1": "...", "input1": "...", ... } }
 */
function getTableClientMap(ss) {
  var tableSheet = ss.getSheetByName("客戶表");
  if (!tableSheet) {
    console.log("ERROR: 找不到 '客戶表' 工作表");
    return null;
  }

  var data = tableSheet.getDataRange().getValues();
  if (data.length < 2) {
    console.log("ERROR: TableClient 資料不足");
    return null;
  }

  var headers = data[0];
  var clientPlatformIdx = headers.indexOf("客戶平台");
  if (clientPlatformIdx === -1) {
    console.log("ERROR: TableClient 找不到 '客戶平台' 欄位");
    return null;
  }

  var map = {};
  for (var i = 1; i < data.length; i++) {
    var platform = data[i][clientPlatformIdx];
    if (!platform) continue;

    var config = {};
    for (var j = 0; j < headers.length; j++) {
      config[headers[j]] = data[i][j];
    }
    map[platform] = config;
  }
  return map;
}

/**
 * 在外部 Google Sheet 的指定工作表中，搜尋包含 searchKey 的儲存格，
 * 並取得同列中該儲存格之後的第一個非空值。
 * @param {Spreadsheet} externalSS - 外部 Google Sheet
 * @param {string} sheetName - 工作表名稱
 * @param {string} searchKey - 要搜尋的關鍵字
 * @returns {*} - 找到的值，或 null
 */
function lookupValueInSheet(externalSS, sheetName, searchKey) {
  var sheet = externalSS.getSheetByName(sheetName);
  if (!sheet) {
    console.log("  [lookup] ERROR: 找不到工作表 '" + sheetName + "' (可用: " + externalSS.getSheets().map(function (s) { return s.getName(); }).join(", ") + ")");
    return null;
  }

  var data = sheet.getDataRange().getValues();
  var searchStr = searchKey.toString();

  for (var row = 0; row < data.length; row++) {
    for (var col = 0; col < data[row].length; col++) {
      var cellValue = data[row][col];
      if (cellValue !== null && cellValue !== undefined && cellValue.toString().trim() === searchStr) {
        // 找到 searchKey，取同列後面第一個非空值
        for (var nextCol = col + 1; nextCol < data[row].length; nextCol++) {
          var nextValue = data[row][nextCol];
          if (nextValue !== null && nextValue !== undefined && nextValue.toString().trim() !== "") {
            return nextValue;
          }
        }
        return null;
      }
    }
  }
  return null;
}

/**
 * 從 Google Drive/Sheet URL 中提取檔案 ID
 * @param {string} url - Google Drive 或 Google Sheet 的 URL
 * @returns {string|null} - 檔案 ID，或 null
 */
function extractFileIdFromUrl(url) {
  if (!url) return null;
  var match = url.toString().match(/[-\w]{25,}/);
  return match ? match[0] : null;
}

/**
 * 清理 .xlsx 轉檔產生的暫存 Google Sheets 副本
 * @param {string|null} tempFileId - 暫存檔案 ID，null 則跳過
 */
function cleanupTempFile_(tempFileId) {
  if (!tempFileId) return;
  try {
    DriveApp.getFileById(tempFileId).setTrashed(true);
  } catch (e) {
    console.log("  WARNING: 清理暫存轉檔失敗: " + e.message);
  }
}
