/**
 * 指數型投資紀錄 - 手機專屬 APP 後端腳本
 * 支援：新增 (Create)、讀取 (Read)、修改 (Update)、刪除 (Delete) 全功能 CRUD
 */

function doGet(e) {
  // 外部前端 (GitHub Pages) GET 請求
  if (e && e.parameter && e.parameter.action === 'getData') {
    const data = getInitialData();
    return ContentService.createTextOutput(JSON.stringify(data))
      .setMimeType(ContentService.MimeType.JSON);
  }

  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('指數投資 Portfolio')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no, viewport-fit=cover')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * 接收外部前端 (GitHub Pages) POST 請求
 */
function doPost(e) {
  try {
    let payload = {};
    if (e && e.postData && e.postData.contents) {
      payload = JSON.parse(e.postData.contents);
    } else if (e && e.parameter) {
      payload = e.parameter;
    }

    let result;
    const action = payload.action || 'add';

    if (action === 'delete') {
      result = deleteTransaction(payload.rowIndex);
    } else if (action === 'update') {
      result = updateTransaction(payload);
    } else {
      result = addTransaction(payload);
    }

    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, message: '伺服器處理錯誤: ' + err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * 取得 App 首頁資產儀表板與交易明細 (含列號 rowIndex 供編輯與刪除)
 */
function getInitialData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  let usdRate = 32.4;
  let totalCost = 0;
  let totalMarketValue = 0;
  let returnRate = 0;

  // 1. 讀取資產總覽
  try {
    const overviewSheet = ss.getSheetByName('資產總覽與總報酬');
    if (overviewSheet) {
      const rateVal = overviewSheet.getRange('B3').getValue();
      if (typeof rateVal === 'number' && rateVal > 0) usdRate = rateVal;

      const costVal = overviewSheet.getRange('C18').getValue();
      if (typeof costVal === 'number') totalCost = costVal;

      const valVal = overviewSheet.getRange('C19').getValue();
      if (typeof valVal === 'number') totalMarketValue = valVal;

      const retVal = overviewSheet.getRange('C20').getValue();
      if (typeof retVal === 'number') returnRate = retVal;
    }
  } catch (e) {
    Logger.log('讀取總覽錯誤: ' + e);
  }

  // 2. 讀取歷史交易明細（取得最近 30 筆有效交易）
  const recentRecords = [];
  try {
    const sheet = ss.getSheetByName('歷史交易明細');
    if (sheet) {
      const lastRow = sheet.getLastRow();
      if (lastRow >= 6) {
        const data = sheet.getRange(6, 1, lastRow - 5, 10).getValues();
        // 由新到舊倒序抓取
        for (let i = data.length - 1; i >= 0 && recentRecords.length < 30; i--) {
          const row = data[i];
          const dateVal = row[0];
          const ticker = row[1];
          const type = row[3];
          
          if (ticker && type) {
            let dateStr = '';
            let rawDate = '';
            if (dateVal instanceof Date) {
              dateStr = Utilities.formatDate(dateVal, Session.getScriptTimeZone(), 'yyyy/MM/dd');
              rawDate = Utilities.formatDate(dateVal, Session.getScriptTimeZone(), 'yyyy-MM-dd');
            } else {
              dateStr = String(dateVal || '');
              // 嘗試轉為 yyyy-mm-dd
              const parts = dateStr.split('/');
              if (parts.length === 3) {
                rawDate = `${parts[0]}-${String(parts[1]).padStart(2, '0')}-${String(parts[2]).padStart(2, '0')}`;
              } else {
                rawDate = dateStr;
              }
            }

            recentRecords.push({
              rowIndex: 6 + i, // 記錄真實列號，方便定位修改或刪除
              date: dateStr,
              rawDate: rawDate,
              ticker: ticker,
              market: row[2] || '台股',
              type: type,
              price: Number(row[4]) || 0,
              shares: Number(row[5]) || 0,
              fee: Number(row[6]) || 0,
              exchangeRate: Number(row[8]) || 1,
              twdCost: Number(row[9]) || 0
            });
          }
        }
      }
    }
  } catch (e) {
    Logger.log('讀取明細錯誤: ' + e);
  }

  return {
    usdRate: usdRate,
    totalCost: totalCost,
    totalMarketValue: totalMarketValue,
    returnRate: returnRate,
    recentRecords: recentRecords
  };
}

/**
 * 新增交易紀錄
 */
function addTransaction(record) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('歷史交易明細');
    if (!sheet) {
      return { success: false, message: '找不到 [歷史交易明細] 工作表！' };
    }

    const lastRow = sheet.getLastRow();
    const values = sheet.getRange(6, 1, Math.max(1, lastRow - 5), 4).getValues();
    
    let targetRow = 6;
    for (let i = 0; i < values.length; i++) {
      const ticker = values[i][1];
      const type = values[i][3];
      if (ticker && (type === '買進' || type === '賣出' || type === '股息再投入')) {
        targetRow = 6 + i + 1;
      }
    }

    if (targetRow <= lastRow) {
      sheet.insertRowBefore(targetRow);
    }

    const dateParts = (record.date || '').split('-');
    const formattedDate = dateParts.length === 3 
      ? `${dateParts[0]}/${parseInt(dateParts[1], 10)}/${parseInt(dateParts[2], 10)}`
      : (record.date || '');

    const r = targetRow;
    const formulaH = `=IF(D${r}="買進", E${r}*F${r}+G${r}, IF(D${r}="賣出", -(E${r}*F${r}-G${r}), IF(D${r}="股息再投入", 0, 0)))`;
    const formulaJ = `=IF(D${r}="股息再投入", 0, IF(ISNUMBER(H${r}), H${r}*IF(I${r}="", 1, I${r}), ""))`;

    const rowData = [
      formattedDate,
      record.ticker,
      record.market,
      record.type,
      Number(record.price) || 0,
      Number(record.shares) || 0,
      Number(record.fee) || 0,
      formulaH,
      Number(record.exchangeRate) || 1,
      formulaJ
    ];

    const targetRange = sheet.getRange(targetRow, 1, 1, 10);
    try {
      sheet.getRange(6, 1, 1, 10).copyTo(targetRange, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
    } catch (e) {}

    targetRange.setValues([rowData]);

    return { 
      success: true, 
      message: `記帳成功！已新增 [${record.ticker} ${record.type}]` 
    };
  } catch (err) {
    return { success: false, message: '新增失敗: ' + err.toString() };
  }
}

/**
 * 修改既有交易紀錄
 */
function updateTransaction(record) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('歷史交易明細');
    if (!sheet) {
      return { success: false, message: '找不到 [歷史交易明細] 工作表！' };
    }

    const r = Number(record.rowIndex);
    if (!r || r < 6 || r > sheet.getLastRow()) {
      return { success: false, message: `無效的列號 (第 ${r} 列)` };
    }

    const dateParts = (record.date || '').split('-');
    const formattedDate = dateParts.length === 3 
      ? `${dateParts[0]}/${parseInt(dateParts[1], 10)}/${parseInt(dateParts[2], 10)}`
      : (record.date || '');

    const formulaH = `=IF(D${r}="買進", E${r}*F${r}+G${r}, IF(D${r}="賣出", -(E${r}*F${r}-G${r}), IF(D${r}="股息再投入", 0, 0)))`;
    const formulaJ = `=IF(D${r}="股息再投入", 0, IF(ISNUMBER(H${r}), H${r}*IF(I${r}="", 1, I${r}), ""))`;

    const rowData = [
      formattedDate,
      record.ticker,
      record.market,
      record.type,
      Number(record.price) || 0,
      Number(record.shares) || 0,
      Number(record.fee) || 0,
      formulaH,
      Number(record.exchangeRate) || 1,
      formulaJ
    ];

    sheet.getRange(r, 1, 1, 10).setValues([rowData]);

    return {
      success: true,
      message: `已成功更新第 ${r} 列 [${record.ticker} ${record.type}]！`
    };
  } catch (err) {
    return { success: false, message: '更新失敗: ' + err.toString() };
  }
}

/**
 * 刪除指定交易紀錄
 */
function deleteTransaction(rowIndex) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('歷史交易明細');
    if (!sheet) {
      return { success: false, message: '找不到 [歷史交易明細] 工作表！' };
    }

    const r = Number(rowIndex);
    if (!r || r < 6 || r > sheet.getLastRow()) {
      return { success: false, message: `無法刪除無效的列號 (第 ${r} 列)` };
    }

    // 取得即將被刪除的標的名稱供提示回傳
    const ticker = sheet.getRange(r, 2).getValue();
    const type = sheet.getRange(r, 4).getValue();

    sheet.deleteRow(r);

    return {
      success: true,
      message: `已成功刪除第 ${r} 列 [${ticker} ${type}]！`
    };
  } catch (err) {
    return { success: false, message: '刪除失敗: ' + err.toString() };
  }
}
