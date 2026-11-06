/**
 * 指數型投資紀錄 - 手機專屬 APP 後端腳本
 * 支援：JSON API (給 GitHub Pages / 獨立前端) 與原生 Web App 介面
 */

function doGet(e) {
  // 如果是外部前端 (如 GitHub Pages) 發送 GET 請求取得資產與明細資料
  if (e && e.parameter && e.parameter.action === 'getData') {
    const data = getInitialData();
    return ContentService.createTextOutput(JSON.stringify(data))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // 預設直接在 Apps Script 開啟時呈現的 HTML 頁面
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('指數投資 Portfolio')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no, viewport-fit=cover')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * 接收外部前端 (如 GitHub Pages) 發送的 POST 交易資料
 */
function doPost(e) {
  try {
    let record = {};
    if (e && e.postData && e.postData.contents) {
      record = JSON.parse(e.postData.contents);
    } else if (e && e.parameter) {
      record = e.parameter;
    }
    const result = addTransaction(record);
    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, message: '伺服器處理錯誤: ' + err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * 取得 App 首頁資產儀表板與交易紀錄
 */
function getInitialData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  let usdRate = 32.4;
  let totalCost = 0;
  let totalMarketValue = 0;
  let returnRate = 0;

  // 1. 從「資產總覽與總報酬」讀取即時資產數據
  try {
    const overviewSheet = ss.getSheetByName('資產總覽與總報酬');
    if (overviewSheet) {
      // B3: 今日美元/台幣匯率
      const rateVal = overviewSheet.getRange('B3').getValue();
      if (typeof rateVal === 'number' && rateVal > 0) usdRate = rateVal;

      // C18: 累計總投入成本 (TWD)
      const costVal = overviewSheet.getRange('C18').getValue();
      if (typeof costVal === 'number') totalCost = costVal;

      // C19: 目前總市值 (TWD)
      const valVal = overviewSheet.getRange('C19').getValue();
      if (typeof valVal === 'number') totalMarketValue = valVal;

      // C20: 當前實際總報酬率
      const retVal = overviewSheet.getRange('C20').getValue();
      if (typeof retVal === 'number') returnRate = retVal;
    }
  } catch (e) {
    Logger.log('讀取總覽錯誤: ' + e);
  }

  // 2. 從「歷史交易明細」取得最近 10 筆紀錄
  const recentRecords = [];
  try {
    const sheet = ss.getSheetByName('歷史交易明細');
    if (sheet) {
      const lastRow = sheet.getLastRow();
      if (lastRow >= 6) {
        const data = sheet.getRange(6, 1, lastRow - 5, 10).getValues();
        for (let i = data.length - 1; i >= 0 && recentRecords.length < 10; i--) {
          const row = data[i];
          const dateVal = row[0];
          const ticker = row[1];
          const type = row[3];
          if (ticker && type) {
            let dateStr = '';
            if (dateVal instanceof Date) {
              dateStr = Utilities.formatDate(dateVal, Session.getScriptTimeZone(), 'yyyy/MM/dd');
            } else {
              dateStr = String(dateVal || '');
            }
            recentRecords.push({
              date: dateStr,
              ticker: ticker,
              market: row[2],
              type: type,
              price: Number(row[4]) || 0,
              shares: Number(row[5]) || 0,
              fee: Number(row[6]) || 0,
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
 * 新增一筆交易紀錄
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
    
    // 找出真實交易的最後一列
    let targetRow = 6;
    for (let i = 0; i < values.length; i++) {
      const ticker = values[i][1];
      const type = values[i][3];
      if (ticker && (type === '買進' || type === '賣出' || type === '股息再投入')) {
        targetRow = 6 + i + 1;
      }
    }

    // 插入新列以保留未來的規劃列
    if (targetRow <= lastRow) {
      sheet.insertRowBefore(targetRow);
    }

    // 格式化日期
    const dateParts = (record.date || '').split('-');
    const formattedDate = dateParts.length === 3 
      ? `${dateParts[0]}/${parseInt(dateParts[1], 10)}/${parseInt(dateParts[2], 10)}`
      : (record.date || '');

    const r = targetRow;
    const formulaH = `=IF(D${r}="買進", E${r}*F${r}+G${r}, IF(D${r}="賣出", -(E${r}*F${r}-G${r}), IF(D${r}="股息再投入", 0, 0)))`;
    const formulaJ = `=IF(D${r}="股息再投入", 0, IF(ISNUMBER(H${r}), H${r}*IF(I${r}="", 1, I${r}), ""))`;

    const rowData = [
      formattedDate,                     // A: 交易日期
      record.ticker,                    // B: 標的代號
      record.market,                    // C: 市場
      record.type,                      // D: 交易類型
      Number(record.price) || 0,        // E: 成交單價
      Number(record.shares) || 0,       // F: 成交股數
      Number(record.fee) || 0,          // G: 交易手續費
      formulaH,                         // H: 單筆投入本金 (原幣)
      Number(record.exchangeRate) || 1, // I: 當下成交匯率
      formulaJ                          // J: 折合台幣總成本
    ];

    const targetRange = sheet.getRange(targetRow, 1, 1, 10);
    // 複製第 6 列的格式
    try {
      sheet.getRange(6, 1, 1, 10).copyTo(targetRange, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
    } catch (e) {}

    targetRange.setValues([rowData]);

    return { 
      success: true, 
      message: `記帳成功！已寫入第 ${targetRow} 列 [${record.ticker} ${record.type}]` 
    };
  } catch (err) {
    return { success: false, message: '寫入失敗: ' + err.toString() };
  }
}
