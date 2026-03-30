/*================================================================================================================*
  ReadClientData - 從 Attachment Url 讀取外部 Google Sheet 資料
  ================================================================================================================
  Version:      1.0.0
  Description:  根據 TableClient 工作表的設定，讀取 Attachment Url 指向的外部 Google Sheet，
                搜尋指定欄位的值並回填至目前月份工作表。

  Changelog:
  1.0.0  初始版本：讀取 TableClient 設定，自動從外部 Sheet 查找並回填數值

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
    console.log("========== readClientData START ==========");
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var dataSheet = ss.getActiveSheet();
    var sheetName = dataSheet.getName();
    console.log("Active sheet: '" + sheetName + "'");

    // 檢查名稱格式 yyyy/mm
    if (!/^\d{4}\/\d{2}$/.test(sheetName)) {
      console.log("ERROR: 工作表名稱不符合 yyyy/mm 格式: '" + sheetName + "'");
      showUiDialog("錯誤", "目前工作表名稱必須為 yyyy/mm 格式，例如 2026/02。");
      return;
    }

    // 讀取 TableClient 設定
    console.log("--- 讀取 TableClient 設定 ---");
    var clientMap = getTableClientMap(ss);
    if (!clientMap || Object.keys(clientMap).length === 0) {
      console.log("ERROR: TableClient 工作表沒有資料或不存在");
      showUiDialog("錯誤", "TableClient 工作表沒有資料或不存在。");
      return;
    }
    console.log("TableClient 載入完成，共 " + Object.keys(clientMap).length + " 筆客戶設定");
    console.log("TableClient keys: " + JSON.stringify(Object.keys(clientMap)));

    // 讀取當前工作表資料
    var sheetValues = dataSheet.getDataRange().getValues();
    var dataHeader = sheetValues[0];
    console.log("Data header columns: " + JSON.stringify(dataHeader));

    var attUrlIndex = dataHeader.indexOf(SETTINGS.attFileColName); // "Attachment Url"
    var clientPlatformIndex = dataHeader.indexOf("客戶平台");
    var input1Index = dataHeader.indexOf("input1");

    console.log("Column indices - Attachment Url: " + attUrlIndex + ", 客戶平台: " + clientPlatformIndex + ", input1: " + input1Index);

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
    var totalRows = sheetValues.length - 1;
    console.log("共 " + totalRows + " 列資料待處理");

    for (var i = 1; i < sheetValues.length; i++) {
      var rowData = sheetValues[i];
      var attUrl = rowData[attUrlIndex];
      var clientPlatform = rowData[clientPlatformIndex];

      console.log("--- Row " + (i + 1) + " ---");
      console.log("  客戶平台: '" + clientPlatform + "'");
      console.log("  Attachment Url: '" + (attUrl || "(空)") + "'");

      // 跳過沒有 Attachment Url 的列
      if (!attUrl) {
        console.log("  SKIP: Attachment Url 為空");
        continue;
      }

      // 從 TableClient 取得此客戶的設定
      var clientConfig = clientMap[clientPlatform];
      if (!clientConfig) {
        console.log("  SKIP: 在 TableClient 找不到客戶平台 '" + clientPlatform + "' 的設定");
        console.log("  可用的客戶平台: " + JSON.stringify(Object.keys(clientMap)));
        skippedCount++;
        continue;
      }
      console.log("  TableClient 設定: " + JSON.stringify(clientConfig));

      // 取得 sheet_input1 和 input1 設定
      var sheetInput1 = clientConfig["sheet_input1"];
      var searchKey1 = clientConfig["input1"];
      console.log("  sheet_input1: '" + sheetInput1 + "', input1 (searchKey): '" + searchKey1 + "'");

      if (!sheetInput1 || !searchKey1) {
        console.log("  SKIP: 缺少 sheet_input1 或 input1 設定");
        skippedCount++;
        continue;
      }

      // 開啟 Attachment Url 指向的 Google Sheet
      var externalSS;
      try {
        var fileId = extractFileIdFromUrl(attUrl);
        console.log("  Extracted fileId: '" + fileId + "' from url: '" + attUrl + "'");
        if (!fileId) {
          console.log("  SKIP: 無法從 Attachment Url 取得檔案 ID");
          skippedCount++;
          continue;
        }
        externalSS = SpreadsheetApp.openById(fileId);
        console.log("  成功開啟外部 Sheet: '" + externalSS.getName() + "'");
        var externalSheets = externalSS.getSheets().map(function(s) { return s.getName(); });
        console.log("  外部 Sheet 包含的工作表: " + JSON.stringify(externalSheets));
      } catch (e) {
        console.log("  ERROR: 無法開啟外部 Google Sheet: " + e.message);
        console.log("  Stack: " + (e.stack || "N/A"));
        skippedCount++;
        continue;
      }

      // 查找主要值 (sheet_input1 + input1)
      console.log("  === 查找主要值: sheet='" + sheetInput1 + "', key='" + searchKey1 + "' ===");
      var mainValue = lookupValueInSheet(externalSS, sheetInput1, searchKey1);
      console.log("  主要值查找結果: " + (mainValue !== null ? mainValue : "NULL (未找到)"));

      if (mainValue === null) {
        console.log("  SKIP: 在 sheet '" + sheetInput1 + "' 中找不到 '" + searchKey1 + "' 的值");
        skippedCount++;
        continue;
      }

      var totalValue = parseFloat(mainValue) || 0;
      console.log("  主要值 (parsed): " + totalValue);

      // 處理 add|1 ~ add|3 的額外值
      for (var addIdx = 1; addIdx <= 3; addIdx++) {
        var addSheetKey = "sheet_input1|add|" + addIdx;
        var addInputKey = "input1|add|" + addIdx;

        var addSheetName = clientConfig[addSheetKey];
        var addSearchKey = clientConfig[addInputKey];

        console.log("  --- add|" + addIdx + " ---");
        console.log("    " + addSheetKey + ": '" + (addSheetName || "(空)") + "'");
        console.log("    " + addInputKey + ": '" + (addSearchKey || "(空)") + "'");

        if (!addSheetName || !addSearchKey) {
          console.log("    SKIP: add|" + addIdx + " 設定為空");
          continue;
        }

        var addValue = lookupValueInSheet(externalSS, addSheetName, addSearchKey);
        console.log("    add|" + addIdx + " 查找結果: " + (addValue !== null ? addValue : "NULL (未找到)"));

        if (addValue !== null) {
          var parsedAddValue = parseFloat(addValue) || 0;
          totalValue += parsedAddValue;
          console.log("    add|" + addIdx + " parsed: " + parsedAddValue + ", 累計 totalValue: " + totalValue);
        } else {
          console.log("    add|" + addIdx + " 未找到值，不累加");
        }
      }

      // 回填 input1 欄位
      console.log("  >>> 回填 input1: Row " + (i + 1) + ", Col " + (input1Index + 1) + ", value: " + totalValue);
      dataSheet.getRange(i + 1, input1Index + 1).setValue(totalValue);
      updatedCount++;
    }

    console.log("========== readClientData END ==========");
    console.log("結果: 更新 " + updatedCount + " 列, 跳過 " + skippedCount + " 列");
    showUiDialog("完成", "讀取客戶資料完成。\n更新：" + updatedCount + " 列\n跳過：" + skippedCount + " 列");
  } catch (e) {
    console.log("FATAL ERROR: " + e.message);
    console.log("Stack: " + (e.stack || "N/A"));
    showUiDialog("錯誤", e.message + "\n" + (e.stack || ""));
  }
}

/**
 * 讀取 TableClient 工作表，以「客戶平台」為 key 建立設定 map
 * @param {Spreadsheet} ss - 當前試算表
 * @returns {Object} - { "客戶平台名稱": { "sheet_input1": "...", "input1": "...", ... } }
 */
function getTableClientMap(ss) {
  console.log("[getTableClientMap] START");
  var tableSheet = ss.getSheetByName("TableClient");
  if (!tableSheet) {
    console.log("[getTableClientMap] ERROR: 找不到 'TableClient' 工作表");
    return null;
  }

  var data = tableSheet.getDataRange().getValues();
  console.log("[getTableClientMap] TableClient 共 " + data.length + " 列 (含標題)");
  if (data.length < 2) {
    console.log("[getTableClientMap] ERROR: TableClient 資料不足 (少於 2 列)");
    return null;
  }

  var headers = data[0];
  console.log("[getTableClientMap] Headers: " + JSON.stringify(headers));
  var clientPlatformIdx = headers.indexOf("客戶平台");
  if (clientPlatformIdx === -1) {
    console.log("[getTableClientMap] ERROR: 找不到 '客戶平台' 欄位");
    return null;
  }

  var map = {};
  for (var i = 1; i < data.length; i++) {
    var platform = data[i][clientPlatformIdx];
    if (!platform) {
      console.log("[getTableClientMap] Row " + (i + 1) + ": 客戶平台為空，跳過");
      continue;
    }

    var config = {};
    for (var j = 0; j < headers.length; j++) {
      config[headers[j]] = data[i][j];
    }
    map[platform] = config;
    console.log("[getTableClientMap] Row " + (i + 1) + ": 載入客戶平台 '" + platform + "' => " + JSON.stringify(config));
  }
  console.log("[getTableClientMap] END, 共載入 " + Object.keys(map).length + " 筆");
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
  console.log("    [lookupValueInSheet] sheet='" + sheetName + "', searchKey='" + searchKey + "'");
  var sheet = externalSS.getSheetByName(sheetName);
  if (!sheet) {
    console.log("    [lookupValueInSheet] ERROR: 找不到工作表 '" + sheetName + "'");
    console.log("    [lookupValueInSheet] 可用的工作表: " + externalSS.getSheets().map(function(s) { return s.getName(); }).join(", "));
    return null;
  }

  var data = sheet.getDataRange().getValues();
  console.log("    [lookupValueInSheet] 工作表 '" + sheetName + "' 大小: " + data.length + " 列 x " + (data.length > 0 ? data[0].length : 0) + " 欄");
  var searchStr = searchKey.toString();

  for (var row = 0; row < data.length; row++) {
    for (var col = 0; col < data[row].length; col++) {
      var cellValue = data[row][col];
      if (cellValue !== null && cellValue !== undefined && cellValue.toString() === searchStr) {
        console.log("    [lookupValueInSheet] MATCH: 在 Row " + (row + 1) + ", Col " + (col + 1) + " 找到 '" + searchStr + "'");
        // 找到 searchKey，取同列後面第一個非空值
        for (var nextCol = col + 1; nextCol < data[row].length; nextCol++) {
          var nextValue = data[row][nextCol];
          if (nextValue !== null && nextValue !== undefined && nextValue.toString().trim() !== "") {
            console.log("    [lookupValueInSheet] FOUND: Row " + (row + 1) + ", Col " + (nextCol + 1) + " = '" + nextValue + "' (type: " + typeof nextValue + ")");
            return nextValue;
          }
        }
        // 該列後面沒有非空值
        console.log("    [lookupValueInSheet] WARNING: 找到 '" + searchStr + "' 但後面沒有非空值 (Row " + (row + 1) + ")");
        return null;
      }
    }
  }

  // 所有列都沒找到
  console.log("    [lookupValueInSheet] NOT FOUND: 在所有 " + data.length + " 列中都沒找到 '" + searchStr + "'");
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
