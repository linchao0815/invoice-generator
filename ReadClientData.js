/*================================================================================================================*
  ReadClientData - 從 Attachment Url 讀取外部 Google Sheet 資料
  ================================================================================================================
  Version:      1.1.0
  Description:  根據「客戶表」工作表中 TableClient 表格的設定，讀取 Attachment Url 指向的外部檔案
                (支援原生 Google Sheets 及 .xlsx)，搜尋指定欄位的值並回填至目前月份工作表。

  前置需求:
  - 需啟用 Drive Advanced Service (Apps Script 編輯器 → 服務 → Drive API)

  Changelog:
  1.0.0  初始版本：讀取 TableClient 設定，從外部 Sheet/xlsx 查找並回填數值
  1.1.0  新增 url 指令：從 TableClient 標題以 | 分隔定義，支援從指定欄位的 URL 開啟外部檔案查找數值

  TableClient 欄位：
    客戶平台 | client_name | 公司名稱 | client_email | client_address | receive_acnt | 幣別 | WHT/VAT
    sheet_input1 | input1 | sheet_input1|add|1 | input1|add|1 | sheet_input1|add|2 | input1|add|2
    sheet_input1|add|3 | input1|add|3

  url 指令欄位格式（標題以 | 分隔）：
    sheet_<target>|url|<url_col>  → 值為外部檔案的工作表名稱
    <target>|url|<url_col>        → 值為搜尋關鍵字；<target> 為 yyyy/mm 工作表的目標欄位；<url_col> 為包含 URL 的欄位
  範例：
    sheet_input1|url|Att2 Url | input1|url|Att2 Url
    → 從當前列 "Att2 Url" 欄位取 URL，開啟外部檔案，在指定 sheet 搜尋關鍵字，結果加總至 "input1" 欄位
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

      // 跳過沒有客戶平台的列
      if (!clientPlatform) {
        continue;
      }

      // 從 TableClient 取得此客戶的設定
      var clientConfig = clientMap[clientPlatform];
      if (!clientConfig) {
        console.log("Row " + (i + 1) + " SKIP: 找不到客戶平台 '" + clientPlatform + "'");
        skippedCount++;
        continue;
      }

      // === 處理主要值 (Attachment Url + sheet_input1 + input1 + add) ===
      if (attUrl) {
        // 取得 sheet_input1 和 input1 設定
        var sheetInput1 = clientConfig["sheet_input1"];
        var searchKey1 = clientConfig["input1"];

        if (!sheetInput1 || !searchKey1) {
          console.log("Row " + (i + 1) + " [" + clientPlatform + "] SKIP: 缺少 sheet_input1 或 input1 設定");
          skippedCount++;
        } else {
          // 開啟 Attachment Url 指向的檔案 (支援 Google Sheets 及 .xlsx，帶快取)
          var externalSS;
          var tempFileId = null;
          var attFileOpened = false;
          try {
            var fileId = extractFileIdFromUrl(attUrl);
            if (!fileId) {
              console.log("Row " + (i + 1) + " SKIP: 無法從 Url 取得檔案 ID");
              skippedCount++;
            } else {
              // 查快取，避免重複開檔/轉檔
              if (fileCache[fileId]) {
                externalSS = fileCache[fileId];
                attFileOpened = true;
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
                attFileOpened = true;
              }
            }
          } catch (e) {
            console.log("Row " + (i + 1) + " ERROR: 無法開啟外部檔案: " + e.message);
            skippedCount++;
            cleanupTempFile_(tempFileId);
          }

          if (attFileOpened) {
            // 查找主要值 (sheet_input1 + input1)
            var mainValue = lookupValueInSheet(externalSS, sheetInput1, searchKey1);

            if (mainValue === null) {
              console.log("Row " + (i + 1) + " [" + clientPlatform + "] SKIP: 在 '" + sheetInput1 + "' 找不到 '" + searchKey1 + "'");
              skippedCount++;
            } else {
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
            }

            // 清理暫存轉檔
            cleanupTempFile_(tempFileId);
          }
        }
      } // end if (attUrl)

      // === 處理 url 指令（獨立於 Attachment Url）===
      var urlDirectives = parseUrlDirectives_(clientConfig);
      if (urlDirectives.length > 0) {
        processUrlDirectives_(urlDirectives, clientConfig, rowData, dataHeader, dataSheet, i, clientPlatform, fileCache);
      }
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
  var clientPlatformIdx = -1;
  for (var h = 0; h < headers.length; h++) {
    if (headers[h] !== null && headers[h] !== undefined && headers[h].toString().trim() === "客戶平台") {
      clientPlatformIdx = h;
      break;
    }
  }
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
      var headerStr = headers[j] !== null && headers[j] !== undefined ? headers[j].toString().trim() : "";
      if (headerStr) {
        config[headerStr] = data[i][j];
      }
    }
    map[platform] = config;
  }

  // Debug: 印出含有 url 指令的標題
  var urlHeaders = Object.keys(map[Object.keys(map)[0]] || {}).filter(function(h) { return h.indexOf("|") > -1 && h.indexOf("url") > -1; });
  if (urlHeaders.length > 0) {
    console.log("  [TableClient] url 指令標題: " + JSON.stringify(urlHeaders));
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

/**
 * 從 clientConfig 的 key（即 TableClient 標題）中解析所有 url 指令組。
 * 標題格式：<target>|url|<url_col> 或 sheet_<target>|url|<url_col>
 * 以 url_col 分組，每組包含 sheetName（來自 sheet_ 前綴標題的值）和 searchKey（來自無前綴標題的值）。
 *
 * @param {Object} clientConfig - 該客戶的 TableClient 設定（key=標題, value=儲存格值）
 * @returns {Array} - url 指令組陣列，每個元素為 { targetCol, urlSourceCol, sheetName, searchKey }
 */
function parseUrlDirectives_(clientConfig) {
  var groups = {}; // key = urlSourceCol, value = { targetCol, sheetName, searchKey }

  var keys = Object.keys(clientConfig);
  for (var k = 0; k < keys.length; k++) {
    var header = keys[k];
    var parts = header.split("|");
    if (parts.length !== 3 || parts[1].trim() !== "url") {
      continue;
    }

    var segment1 = parts[0].trim();
    var urlSourceCol = parts[2].trim();

    if (!urlSourceCol) continue;

    // segment1 不應有 sheet_ 前綴（sheet 名稱由獨立的 sheet_<target> 欄位提供）
    if (segment1.indexOf("sheet_") === 0) {
      continue; // 跳過，不應出現此格式
    }

    if (!groups[urlSourceCol]) {
      groups[urlSourceCol] = { targetCol: null, sheetName: null, searchKey: null, urlSourceCol: urlSourceCol };
    }

    // segment1 為目標欄位名稱，值為搜尋關鍵字
    groups[urlSourceCol].targetCol = segment1;
    groups[urlSourceCol].searchKey = clientConfig[header] ? clientConfig[header].toString() : "";

    // sheet 名稱從獨立的 "sheet_<target>" 欄位取得
    var sheetKey = "sheet_" + segment1;
    if (clientConfig[sheetKey] !== undefined && clientConfig[sheetKey] !== null && clientConfig[sheetKey].toString().trim() !== "") {
      groups[urlSourceCol].sheetName = clientConfig[sheetKey].toString().trim();
    }
  }

  // 轉為陣列，過濾掉不完整的組
  var result = [];
  var groupKeys = Object.keys(groups);
  for (var g = 0; g < groupKeys.length; g++) {
    var group = groups[groupKeys[g]];
    if (group.targetCol && group.sheetName && group.searchKey && group.urlSourceCol) {
      result.push(group);
    } else {
      console.log("  [parseUrlDirectives] SKIP incomplete group: urlSourceCol='" + group.urlSourceCol + "' targetCol='" + group.targetCol + "' sheetName='" + group.sheetName + "' searchKey='" + group.searchKey + "'");
    }
  }
  if (result.length > 0) {
    console.log("  [parseUrlDirectives] 解析到 " + result.length + " 組 url 指令");
  }
  return result;
}

/**
 * 處理所有 url 指令組：開啟外部檔案、查找數值、加總並寫入目標欄位。
 *
 * @param {Array} urlDirectives - parseUrlDirectives_ 回傳的指令組陣列
 * @param {Object} clientConfig - 該客戶的 TableClient 設定
 * @param {Array} rowData - 當前列的資料陣列
 * @param {Array} dataHeader - yyyy/mm 工作表的標題列
 * @param {Sheet} dataSheet - yyyy/mm 工作表物件
 * @param {number} rowIndex - 當前列的 0-based index（sheetValues 中的 index）
 * @param {string} clientPlatform - 客戶平台名稱（用於日誌）
 * @param {Object} fileCache - 檔案快取物件
 */
function processUrlDirectives_(urlDirectives, clientConfig, rowData, dataHeader, dataSheet, rowIndex, clientPlatform, fileCache) {
  // 依 targetCol 分組累計值
  var targetAccum = {}; // { targetCol: { total: number, sources: string } }

  for (var d = 0; d < urlDirectives.length; d++) {
    var directive = urlDirectives[d];
    var urlSourceCol = directive.urlSourceCol;
    var targetCol = directive.targetCol;
    var sheetName = directive.sheetName;
    var searchKey = directive.searchKey;

    // 從當前列取得 URL
    var urlColIndex = dataHeader.indexOf(urlSourceCol);
    if (urlColIndex === -1) {
      console.log("Row " + (rowIndex + 1) + " [" + clientPlatform + "] url SKIP: yyyy/mm 工作表找不到欄位 '" + urlSourceCol + "'");
      continue;
    }

    var urlValue = rowData[urlColIndex];
    if (!urlValue) {
      // 靜默跳過：該列此欄位無 URL
      continue;
    }

    // 開啟外部檔案（帶快取）
    var externalSS;
    var tempFileId = null;
    try {
      var fileId = extractFileIdFromUrl(urlValue);
      if (!fileId) {
        console.log("Row " + (rowIndex + 1) + " [" + clientPlatform + "] url SKIP: 無法從 '" + urlSourceCol + "' 的 URL 取得檔案 ID");
        continue;
      }

      if (fileCache[fileId]) {
        externalSS = fileCache[fileId];
      } else {
        var driveFile = DriveApp.getFileById(fileId);
        var mimeType = driveFile.getMimeType();

        if (mimeType === "application/vnd.google-apps.spreadsheet") {
          externalSS = SpreadsheetApp.openById(fileId);
        } else {
          var copiedFile = Drive.Files.copy(
            { title: "temp_readClient_url_" + fileId, mimeType: "application/vnd.google-apps.spreadsheet" },
            fileId
          );
          tempFileId = copiedFile.id;
          externalSS = SpreadsheetApp.openById(tempFileId);
        }
        fileCache[fileId] = externalSS;
      }
    } catch (e) {
      console.log("Row " + (rowIndex + 1) + " [" + clientPlatform + "] url ERROR: 無法開啟 '" + urlSourceCol + "' 的外部檔案: " + e.message);
      cleanupTempFile_(tempFileId);
      continue;
    }

    // 查找數值
    var foundValue = lookupValueInSheet(externalSS, sheetName, searchKey);
    if (foundValue === null) {
      console.log("Row " + (rowIndex + 1) + " [" + clientPlatform + "] url WARNING: 在 '" + sheetName + "' 找不到 '" + searchKey + "' (來源: " + urlSourceCol + ")");
      cleanupTempFile_(tempFileId);
      continue;
    }

    var parsedValue = parseFloat(foundValue);
    var numValue = isNaN(parsedValue) ? 0 : parsedValue;

    // 累計至目標欄位
    if (!targetAccum[targetCol]) {
      targetAccum[targetCol] = { total: 0, sources: "" };
    }
    targetAccum[targetCol].total += numValue;
    var srcLabel = urlSourceCol + ">" + sheetName + "/" + searchKey + "=" + numValue;
    targetAccum[targetCol].sources += (targetAccum[targetCol].sources ? " + " : "") + srcLabel;

    console.log("Row " + (rowIndex + 1) + " [" + clientPlatform + "] url OK: " + targetCol + " += " + numValue + " (" + srcLabel + ")");
    cleanupTempFile_(tempFileId);
  }

  // 寫入各目標欄位
  var targetCols = Object.keys(targetAccum);
  for (var t = 0; t < targetCols.length; t++) {
    var col = targetCols[t];
    var colIndex = dataHeader.indexOf(col);
    if (colIndex === -1) {
      console.log("Row " + (rowIndex + 1) + " [" + clientPlatform + "] url ERROR: yyyy/mm 工作表找不到目標欄位 '" + col + "'");
      continue;
    }

    // 讀取目前欄位的現有值，加總 url 指令的結果
    var currentValue = dataSheet.getRange(rowIndex + 1, colIndex + 1).getValue();
    var currentNum = parseFloat(currentValue);
    if (isNaN(currentNum)) currentNum = 0;

    var newValue = currentNum + targetAccum[col].total;
    dataSheet.getRange(rowIndex + 1, colIndex + 1).setValue(newValue);
    console.log("Row " + (rowIndex + 1) + " [" + clientPlatform + "] url WRITE: " + col + " = " + currentNum + " + " + targetAccum[col].total + " = " + newValue + " (" + targetAccum[col].sources + ")");
  }
}
