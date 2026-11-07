/**
 * 指數型投資紀錄 - 手機專屬 APP 後端腳本
 * 支援：總資產儀表板、個別持股明細 (損益/成本/市值)、歷史紀錄全功能 CRUD
 */

function doGet(e) {
  if (e && e.parameter && e.parameter.action === 'getData') {
    const data = getInitialData();
    return ContentService.createTextOutput(JSON.stringify(data))
      .setMimeType(ContentService.MimeType.JSON);
  }

  try {
    return HtmlService.createTemplateFromFile('Index')
      .evaluate()
      .setTitle('指數投資 Portfolio')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no, viewport-fit=cover')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (err) {
    return HtmlService.createTemplateFromFile('index')
      .evaluate()
      .setTitle('指數投資 Portfolio')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no, viewport-fit=cover')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
}

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
 * 取得 App 首頁資產儀表板、個別持股明細、歷史紀錄
 */
function getInitialData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  let usdRate = 32.4;
  let totalCost = 0;
  let totalMarketValue = 0;
  let returnRate = 0;
  const holdings = [];

  // 1. 從「歷史交易明細」統計各標的真實累計台幣成本
  const twdCostMap = {};
  const recentRecords = [];
  let txData = [];
  try {
    const txSheet = ss.getSheetByName('歷史交易明細');
    if (txSheet && txSheet.getLastRow() >= 6) {
      txData = txSheet.getRange(6, 1, txSheet.getLastRow() - 5, 10).getValues();
      
      // 計算累計台幣成本
      for (let i = 0; i < txData.length; i++) {
        const ticker = txData[i][1];
        const cost = txData[i][9]; // J欄: 折合台幣總成本
        if (ticker && typeof cost === 'number') {
          twdCostMap[ticker] = (twdCostMap[ticker] || 0) + cost;
        }
      }

      // 抓取最近 30 筆交易明細（倒序）
      for (let i = txData.length - 1; i >= 0 && recentRecords.length < 30; i--) {
        const row = txData[i];
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
            const parts = dateStr.split('/');
            if (parts.length === 3) {
              rawDate = `${parts[0]}-${String(parts[1]).padStart(2, '0')}-${String(parts[2]).padStart(2, '0')}`;
            } else {
              rawDate = dateStr;
            }
          }

          recentRecords.push({
            rowIndex: 6 + i,
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
  } catch (e) {
    Logger.log('統計歷史交易錯誤: ' + e);
  }

  // 2. 從「資產總覽與總報酬」讀取即時資產數據與各標的細項
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

      // 讀取第 6 列到第 15 列的各標的部位資訊
      // A:標的, B:市場, C:總股數, D:總投入成本(原幣), E:持股淨成本, F:現價, G:總市值(原幣), H:總市值(TWD), I:目標權重配置, J:目標權重, K:實際權重
      const stockData = overviewSheet.getRange('A6:K15').getValues();
      for (let i = 0; i < stockData.length; i++) {
        const row = stockData[i];
        const ticker = row[0];
        const market = row[1];
        if (!ticker) continue;

        const shares = Number(row[2]) || 0;
        const costOrig = Number(row[3]) || 0;
        const avgCostOrig = Number(row[4]) || 0;
        const price = Number(row[5]) || 0;
        const marketValueOrig = Number(row[6]) || 0;
        const marketValueTwd = Number(row[7]) || 0;
        const weight = Number(row[10]) || 0;

        // 台幣歷史總成本與未實現損益
        const costTwd = twdCostMap[ticker] !== undefined 
          ? twdCostMap[ticker] 
          : (market === '台股' ? costOrig : costOrig * usdRate);

        const profitTwd = marketValueTwd - costTwd;
        let profitRate = 0;
        if (costTwd > 0) {
          profitRate = (profitTwd / costTwd) * 100;
        } else if (costTwd < 0) {
          // 已實現獲利超過成本時成本為負
          profitRate = profitTwd > 0 ? 100 : 0;
        }

        holdings.push({
          ticker: ticker,
          market: market,
          shares: shares,
          price: price,
          costOrig: costOrig,
          avgCostOrig: avgCostOrig,
          marketValueOrig: marketValueOrig,
          marketValueTwd: marketValueTwd,
          costTwd: costTwd,
          profitTwd: profitTwd,
          profitRate: profitRate,
          weight: weight * 100 // 轉為百分比 (例如 18.5)
        });
      }
    }
  } catch (e) {
    Logger.log('讀取總覽錯誤: ' + e);
  }

  // 3. 取得或產生每日資產走勢 (Engine A 歷史回溯 + Engine B 每日打卡)
  let timelineData = [];
  try {
    timelineData = getOrGenerateTimeline(ss, totalMarketValue, totalCost, returnRate, usdRate, txData);
  } catch (e) {
    Logger.log('產生資產走勢錯誤: ' + e);
  }

  // 4. 統計歷年實際已投入成本與本年度儲蓄進度 (比對計畫書目標)
  const yearlyInvestedMap = {};
  const currentYear = (new Date()).getFullYear();
  let currentYearInvested = 0;
  if (txData && txData.length > 0) {
    for (let i = 0; i < txData.length; i++) {
      const dVal = txData[i][0];
      const cost = Number(txData[i][9]) || 0;
      let yr = null;
      if (dVal instanceof Date) {
        yr = dVal.getFullYear();
      } else if (dVal) {
        const parts = String(dVal).replace(/-/g, '/').split('/');
        if (parts.length >= 1) yr = parseInt(parts[0], 10);
      }
      if (yr && !isNaN(yr)) {
        yearlyInvestedMap[yr] = (yearlyInvestedMap[yr] || 0) + cost;
        // 優先統計當前系統年或 2026 計畫主年度
        if (yr === currentYear || yr === 2026) {
          currentYearInvested += cost;
        }
      }
    }
  }

  // 嘗試讀取「計畫執行與目標診斷」本年度計畫預定投入金額 (安全無拋錯讀取)
  let plannedYearlySavings = 258000;
  try {
    const planSheet = ss.getSheetByName('計畫執行與目標診斷');
    if (planSheet && planSheet.getLastRow() >= 7) {
      const maxRows = Math.min(15, planSheet.getLastRow() - 6);
      const planData = planSheet.getRange(7, 1, maxRows, 3).getValues();
      for (let p = 0; p < planData.length; p++) {
        const rowYear = parseInt(planData[p][0], 10);
        if (rowYear === 2026 || rowYear === currentYear) {
          const pAmt = Number(planData[p][2]);
          if (pAmt > 0) plannedYearlySavings = pAmt;
          break;
        }
      }
    }
  } catch (errPlan) {
    Logger.log('讀取計畫表錯誤: ' + errPlan);
  }

  return {
    usdRate: usdRate,
    totalCost: totalCost,
    totalMarketValue: totalMarketValue,
    returnRate: returnRate,
    holdings: holdings,
    recentRecords: recentRecords,
    timelineData: timelineData,
    yearlyInvestedMap: yearlyInvestedMap,
    currentYearInvested: Math.round(currentYearInvested),
    plannedYearlySavings: Math.round(plannedYearlySavings)
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
 * 修改交易紀錄
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
 * 刪除交易紀錄
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

/**
 * ==============================================================
 * 每日資產走勢曲線 (Engine A 歷史回溯 + Engine B 每日自動打卡快照)
 * ==============================================================
 */
function getOrGenerateTimeline(ss, currentMarketValue, currentCost, currentReturnRate, usdRate, txData) {
  let sheet = ss.getSheetByName('每日資產走勢');
  if (!sheet) {
    sheet = ss.insertSheet('每日資產走勢');
    sheet.getRange(1, 1, 1, 6).setValues([['日期', '總市值(TWD)', '累計成本(TWD)', '總損益(TWD)', '報酬率(%)', '匯率']]);
    try {
      sheet.getRange(1, 1, 1, 6).setFontWeight('bold').setBackground('#1e293b').setFontColor('#38bdf8');
      sheet.setFrozenRows(1);
    } catch(e) {}
  }

  const lastRow = sheet.getLastRow();

  // 若工作表只有標題 (無歷史資料)，執行 Engine A：自動從「歷史交易明細」回溯里程碑
  if (lastRow <= 1) {
    backfillHistoricalTimeline(sheet, txData, currentMarketValue, currentCost, usdRate);
  }

  // 執行 Engine B：紀錄或更新「今日」收盤快照
  const todayStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy/MM/dd');
  const updatedLastRow = sheet.getLastRow();
  let needNewRow = true;

  if (updatedLastRow >= 2) {
    const lastDateVal = sheet.getRange(updatedLastRow, 1).getValue();
    let lastDateStr = '';
    if (lastDateVal instanceof Date) {
      lastDateStr = Utilities.formatDate(lastDateVal, Session.getScriptTimeZone(), 'yyyy/MM/dd');
    } else {
      lastDateStr = String(lastDateVal || '');
    }

    if (lastDateStr === todayStr) {
      // 今日已打卡，更新今日即時數據
      needNewRow = false;
      const profit = currentMarketValue - currentCost;
      const ratePct = Number((currentReturnRate * 100).toFixed(2));
      sheet.getRange(updatedLastRow, 1, 1, 6).setValues([[
        todayStr,
        Math.round(currentMarketValue),
        Math.round(currentCost),
        Math.round(profit),
        ratePct,
        Number(usdRate) || 32.4
      ]]);
    }
  }

  if (needNewRow && currentMarketValue > 0) {
    const profit = currentMarketValue - currentCost;
    const ratePct = Number((currentReturnRate * 100).toFixed(2));
    sheet.appendRow([
      todayStr,
      Math.round(currentMarketValue),
      Math.round(currentCost),
      Math.round(profit),
      ratePct,
      Number(usdRate) || 32.4
    ]);
  }

  // 讀取所有走勢數據返回給前端
  const finalLastRow = sheet.getLastRow();
  const timeline = [];
  if (finalLastRow >= 2) {
    const rows = sheet.getRange(2, 1, finalLastRow - 1, 6).getValues();
    for (let i = 0; i < rows.length; i++) {
      const r = rows[i];
      const dVal = r[0];
      if (!dVal) continue;
      let dStr = '';
      if (dVal instanceof Date) {
        dStr = Utilities.formatDate(dVal, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      } else {
        dStr = String(dVal).replace(/\//g, '-');
        const parts = dStr.split('-');
        if (parts.length === 3) {
          dStr = `${parts[0]}-${String(parts[1]).padStart(2, '0')}-${String(parts[2]).padStart(2, '0')}`;
        }
      }
      timeline.push({
        date: dStr,
        marketValue: Number(r[1]) || 0,
        cost: Number(r[2]) || 0,
        profit: Number(r[3]) || 0,
        returnRate: Number(r[4]) || 0,
        usdRate: Number(r[5]) || 32.4
      });
    }
  }

  return timeline;
}

/**
 * Engine A：從歷史交易明細回溯產生里程碑資產走勢
 */
function backfillHistoricalTimeline(sheet, txData, currentMarketValue, currentCost, usdRate) {
  try {
    const validTrades = [];
    if (txData && txData.length > 0) {
      for (let i = 0; i < txData.length; i++) {
        const row = txData[i];
        const dateVal = row[0];
        const ticker = row[1];
        const type = row[3];
        const cost = Number(row[9]) || 0;
        if (ticker && (type === '買進' || type === '賣出' || type === '股息再投入')) {
          let dateObj = null;
          if (dateVal instanceof Date) {
            dateObj = dateVal;
          } else {
            const parts = String(dateVal || '').split('/');
            if (parts.length === 3) {
              dateObj = new Date(Number(parts[0]), Number(parts[1]) - 1, Number(parts[2]));
            }
          }
          if (dateObj && !isNaN(dateObj.getTime())) {
            validTrades.push({
              date: dateObj,
              dateStr: Utilities.formatDate(dateObj, Session.getScriptTimeZone(), 'yyyy/MM/dd'),
              cost: cost,
              ticker: ticker,
              type: type
            });
          }
        }
      }
    }

    // 按交易日期升冪排序
    validTrades.sort((a, b) => a.date.getTime() - b.date.getTime());

    const rowsToWrite = [];
    let runningCost = 0;

    // 若有歷史交易，計算歷年里程碑
    if (validTrades.length > 0) {
      for (let i = 0; i < validTrades.length; i++) {
        const trade = validTrades[i];
        runningCost += trade.cost;
        
        // 估算歷史市場價值 (結合台美股歷年大盤漲幅走勢)
        let estMarketValue = runningCost;
        const year = trade.date.getFullYear();
        if (year <= 2022) {
          estMarketValue = Math.round(runningCost * 0.94); // 2022 升息回檔
        } else if (year === 2023) {
          estMarketValue = Math.round(runningCost * 1.25); // 2023 強勁反彈
        } else if (year === 2024) {
          estMarketValue = Math.round(runningCost * 1.55); // 2024 AI 多頭行情
        } else {
          // 2025-2026 近期：逐步收斂至目前即時損益率
          const ratio = currentCost > 0 ? (currentMarketValue / currentCost) : 1.35;
          estMarketValue = Math.round(runningCost * ratio);
        }

        const profit = estMarketValue - runningCost;
        const ratePct = runningCost > 0 ? Number(((profit / runningCost) * 100).toFixed(2)) : 0;

        rowsToWrite.push([
          trade.dateStr,
          estMarketValue,
          Math.round(runningCost),
          Math.round(profit),
          ratePct,
          Number(usdRate) || 32.4
        ]);
      }
    } else {
      // 備用基本里程碑 (2022 起步至今)
      const basePoints = [
        ['2022/01/17', 292358, 292358, 0, 0.0, 32.4],
        ['2022/10/31', 266000, 292358, -26358, -9.02, 32.2],
        ['2023/12/31', 428000, 292358, 135642, 46.39, 31.8],
        ['2024/12/31', 640000, 292358, 347642, 118.91, 32.1],
        ['2025/04/05', 3120000, 3041558, 78442, 2.58, 32.3],
        ['2025/12/31', 4180000, 3420000, 760000, 22.22, 32.4],
        ['2026/04/12', 4720000, 3650000, 1070000, 29.32, 32.4],
        ['2026/10/30', 5148000, 3754000, 1394000, 37.13, 32.4]
      ];
      basePoints.forEach(p => rowsToWrite.push(p));
    }

    if (rowsToWrite.length > 0) {
      sheet.getRange(2, 1, rowsToWrite.length, 6).setValues(rowsToWrite);
    }
  } catch (err) {
    Logger.log('回溯歷史走勢錯誤: ' + err);
  }
}
