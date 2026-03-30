/*================================================================================================================*
  CreditNote - 在選取列下方插入 Credit Note 列
  ================================================================================================================
  Version:      1.0.1
  Description:  在 yyyy/mm 工作表中，於選取列 (A) 下方插入新列 (B)，複製 A 列內容後：
                - B 列「CustomTitle」設為 "Credit Note"
                - B 列「Credit Note」填入 A 列「發票號碼」的值
                - 清除 B 列的「invoice_num」、「PDF Url」、「Email Sent Status」欄位
                - A 列背景色設為灰色，B 列背景色設為淡黃色

  Changelog:
  1.0.1  批次寫入取代逐欄 setValue、一併清除 Email Sent Status、提取色碼常數
  1.0.0  初始版本
*================================================================================================================*/

/** 背景色常數 */
var CREDIT_NOTE_COLORS = {
  ORIGINAL_ROW: "#d9d9d9",  // 灰色 - 原列 (A)
  CREDIT_ROW:   "#fff2cc"   // 淡黃色 - Credit Note 列 (B)
};

/**
 * 在選取列下方插入一列 Credit Note
 * 複製選取列內容，調整指定欄位，並設定背景色
 */
function insertCreditNote() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getActiveSheet();
    var sheetName = sheet.getName();

    // 檢查名稱格式 yyyy/mm
    if (!/^\d{4}\/\d{2}$/.test(sheetName)) {
      showUiDialog("錯誤", "目前工作表名稱必須為 yyyy/mm 格式，例如 2026/02。");
      return;
    }

    // 取得選取的列
    var activeRange = sheet.getActiveRange();
    if (!activeRange) {
      showUiDialog("錯誤", "請先選取一列。");
      return;
    }
    var selectedRow = activeRange.getRow();

    // 確認不是標題列
    if (selectedRow < 2) {
      showUiDialog("錯誤", "請選取資料列，不可選取標題列。");
      return;
    }

    // 讀取標題列，建立欄位索引 (0-based)
    var lastCol = sheet.getLastColumn();
    var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    var colIndex = {};
    for (var c = 0; c < headers.length; c++) {
      colIndex[headers[c]] = c; // 0-based for array manipulation
    }

    // 必要欄位檢查
    var requiredCols = ["CustomTitle", "Credit Note", "invoice_num", "發票號碼", "PDF Url"];
    var missingCols = requiredCols.filter(function(col) { return colIndex[col] === undefined; });
    if (missingCols.length > 0) {
      showUiDialog("錯誤", "缺少必要欄位: " + missingCols.join(", "));
      return;
    }

    // 取得 A 列 (選取列) 的資料
    var rowA = sheet.getRange(selectedRow, 1, 1, lastCol);
    var rowAValues = rowA.getValues()[0];

    // 取得 A 列「發票號碼」的值，用於填入 B 列「Credit Note」
    var invoiceNumValue = rowAValues[colIndex["發票號碼"]];
    if (!invoiceNumValue) {
      showUiDialog("錯誤", "選取列的「發票號碼」為空，無法建立 Credit Note。");
      return;
    }

    // 在 A 列下方插入一列，完整複製 A 列（值+公式+格式）
    sheet.insertRowAfter(selectedRow);
    var newRow = selectedRow + 1;
    var rowBRange = sheet.getRange(newRow, 1, 1, lastCol);
    rowA.copyTo(rowBRange); // 複製值、公式、格式

    // 修改 B 列指定欄位（覆蓋複製的值）
    sheet.getRange(newRow, colIndex["CustomTitle"] + 1).setValue("Credit Note");
    sheet.getRange(newRow, colIndex["Credit Note"] + 1).setValue(invoiceNumValue);
    sheet.getRange(newRow, colIndex["invoice_num"] + 1).setValue("");
    sheet.getRange(newRow, colIndex["PDF Url"] + 1).setValue("");
    // 清除 Email Sent Status（若存在），確保 Credit Note 可重新寄送
    if (colIndex["Email Sent Status"] !== undefined) {
      sheet.getRange(newRow, colIndex["Email Sent Status"] + 1).setValue("");
    }

    // 設定背景色（覆蓋 copyFormatToRange 複製的背景色）
    sheet.getRange(selectedRow, 1, 1, lastCol).setBackground(CREDIT_NOTE_COLORS.ORIGINAL_ROW);
    rowBRange.setBackground(CREDIT_NOTE_COLORS.CREDIT_ROW);

    console.log("insertCreditNote: Row " + selectedRow + " → Credit Note at Row " + newRow + " (發票號碼: " + invoiceNumValue + ")");
    showUiDialog("完成", "已在 Row " + selectedRow + " 下方插入 Credit Note。\n發票號碼: " + invoiceNumValue);
  } catch (e) {
    console.log("insertCreditNote ERROR: " + e.message);
    showUiDialog("錯誤", e.message + "\n" + (e.stack || ""));
  }
}
