/**
 * 接收 POST 請求，並將資料轉發至 ELK (Elasticsearch)。
 * @param {Object} e - Apps Script 的事件物件，包含 POST 請求的詳細資訊。
 */
function doPost(e) {
  try {
    // 2. 解析收到的 POST 請求內容
    var postData = JSON.parse(e.postData.contents);
    Logger.log(`Received POST data: ${JSON.stringify(postData)}`);
    // 3. (重要) 為資料加上 ELK_Log 需要的時間戳欄位 'ts'
    // ELK_Log 函式會檢查此欄位是否存在。我們使用 ISO 格式的時間字串。
    if (!postData.ts) {
      postData.ts = Utilities.formatDate(new Date(), "Asia/Taipei", "yyyy-MM-dd'T'HH:mm:ssXXX");
    }
    
    // 4. 寫入 Google Sheet
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("log");
    if (!sheet) {
      sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet("log");
    }
    var lastCol = sheet.getLastColumn();
    var headers = [];
    if (sheet.getLastRow() === 0 || lastCol === 0) {
      // 無資料，建立 header
      var keys = Object.keys(postData);
      if (keys.indexOf("ts") === -1) keys.unshift("ts");
      sheet.appendRow(keys);
      headers = keys;
    } else {
      headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
      if (headers.length === 0 || headers[0] !== "ts") {
        // header 不存在或第一欄不是 ts，重建 header
        var keys = Object.keys(postData);
        if (keys.indexOf("ts") === -1) keys.unshift("ts");
        sheet.insertRowBefore(1);
        sheet.getRange(1, 1, 1, keys.length).setValues([keys]);
        headers = keys;
      }
    }
    // 檢查有無新欄位，補齊 header
    var dataKeys = Object.keys(postData);
    dataKeys.forEach(function(k) {
      if (headers.indexOf(k) === -1) {
        headers.push(k);
        sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      }
    });
    // 依 header 順序組成 row，嚴格比對 null/undefined
    var row = headers.map(function(k) {
      return (postData[k] !== undefined && postData[k] !== null) ? postData[k] : "";
    });
    sheet.appendRow(row);

    // 5. 呼叫 ELK_Log 函式將資料送出
    var elkResult = ELK_Log(postData);
    Logger.log(`ELK_Log result: ${JSON.stringify(elkResult)}`);
    var response;
    // 5. 根據 ELK_Log 的執行結果，準備回傳的 JSON 訊息
    if (elkResult.code === 0) { // code 0 代表成功
      response = {
        status: "success",
        message: "Log successfully sent to ELK.",
        // 建議將 ELK 的原始回應也包含進來，方便除錯
        elk_response: JSON.parse(elkResult.response) 
      };
    } else {
      // 若 ELK_Log 回傳非 0 的 code，代表失敗
      Logger.log(`Failed to send log to ELK. Code: ${elkResult.code}, Response: ${elkResult.response}`);
      throw new Error(`Failed to send log to ELK. Code: ${elkResult.code}, Response: ${elkResult.response}`);
    }
    Logger.log(`Response to be sent: ${JSON.stringify(response)}`);
    return ContentService.createTextOutput(JSON.stringify(response))
                         .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    // 統一處理所有可能發生的錯誤 (JSON 解析錯誤、ELK 設定錯誤、ELK 傳送失敗等)
    var errorResponse = {
      "status": "error",
      "message": error.toString()
    };
    
    // 將錯誤記錄到 Apps Script 的日誌中，方便開發者查看
    Logger.log(`doPost Error: ${error.toString()}`);

    return ContentService.createTextOutput(JSON.stringify(errorResponse))
                         .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Google Apps Script 版本的 ELK_Log
 * @param {Object|Array} val - 欲寫入的 JSON 物件或陣列
 * @return {Object} - 回傳結果
 */
/**
 * Google Apps Script 版本的 ELK_Log
 * @param {Object|Array} val - 欲寫入的 JSON 物件或陣列
 * @param {Object} options - 設定參數，包含 url、index、apikey
 * @param {string} options.url - ES 端點
 * @param {string} options.index - index 名稱
 * @param {string} options.apikey - API key
 * @return {Object} - 回傳結果
 */
function ELK_Log(val) {
  // Helper function to process date fields in an object
  const processDates = (obj) => {
    for (const key in obj) {
      if (typeof key === 'string' && key.toLowerCase().includes('date')) {
        try {
          if (obj[key] && !isNaN(new Date(obj[key]).getTime())) {
            obj[key] = Utilities.formatDate(new Date(obj[key]), "Asia/Taipei", "yyyy-MM-dd'T'HH:mm:ssXXX");
          }
        } catch (e) {
          Logger.log(`Could not convert key "${key}" with value "${obj[key]}" to ISOString: ${e.toString()}`);
        }
      }
    }
  };

  var start_ts = Utilities.formatDate(new Date(), "Asia/Taipei", "yyyy-MM-dd'T'HH:mm:ssXXX");
  var now = new Date();
  var buffer = Utilities.formatDate(now, Session.getScriptTimeZone(), "'-'yyyy.MM.dd");
  var url = '';
  var urlBody = '';
  var response = '';
  var cTs = 'ts'; // 時間欄位名稱，請依實際欄位調整
  // 1. 從指令碼屬性讀取 ELK 設定
  var scriptProperties = PropertiesService.getScriptProperties();
  var options = {
    url: scriptProperties.getProperty('ELK_URL'),
    index: scriptProperties.getProperty('ELK_INDEX'),
    apikey: scriptProperties.getProperty('ELK_APIKEY')
  };

  // 檢查設定是否齊全
  if (!options.url || !options.index || !options.apikey) {
    throw new Error("ELK 設定不完整，請檢查指令碼屬性 (ELK_URL, ELK_INDEX, ELK_APIKEY)。");
  }
  // 參數檢查
  if (!options || !options.url || !options.index || !options.apikey) {
    Logger.log('缺少必要的 options 參數');
    return { code: -99, message: '缺少必要的 options 參數' };
  }
  var s_url = options.url;
  var s_index = options.index;
  var s_apikey = options.apikey;

  // 判斷 val 是陣列還是物件
  if (Array.isArray(val)) {
    url = s_url + '/' + s_index + buffer + '/_doc/_bulk';
    var body = '';
    for (var i = 0; i < val.length; i++) {
      processDates(val[i]); // Process each object in the array
      if (!val[i].hasOwnProperty(cTs)) {
        val[i][cTs] = start_ts;
        Logger.log('沒有 ' + cTs + ' 欄位, 自動新增目前時間: ' + val[i][cTs]);
      }
      body += JSON.stringify({ index: {} }) + '\n' + JSON.stringify(val[i]) + '\n';
    }
    urlBody = body;
  } else {
    processDates(val); // Process the single object
    if (!val.hasOwnProperty(cTs)) {
      val[cTs] = start_ts;
      Logger.log('沒有 ' + cTs + ' 欄位, 自動新增目前時間: ' + val[cTs]);
    }
    url = s_url + '/' + s_index + buffer + '/_doc';
    urlBody = JSON.stringify(val);
  }

  var vHeader = [
    'Content-Type: application/x-ndjson;charset=UTF-8',
    'Authorization: ApiKey ' + s_apikey
  ];

  var retry = 0;
  var maxRetry = 3;
  var result = {};
  do {
    try {
      var httpResult = HttpPost(url, urlBody, vHeader, 10000);
      response = httpResult.response;
      if (httpResult.code >= 200 && httpResult.code < 300) {
        Logger.log(response);
        result = { code: 0, response: response };
        break;
      } else {
        Logger.log('url: ' + url + ' reTry:' + retry + ' Failed:' + httpResult.code + ' :' + response);
        result = { code: httpResult.code, response: response };
      }
    } catch (e) {
      Logger.log('Exception: ' + e);
      result = { code: -2, response: e.toString() };
    }
    retry++;
  } while (retry < maxRetry);

  if (retry === maxRetry) {
    Logger.log('Post failed after 3 retries');
  }

  return result;
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
    headers.forEach(function(h) {
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