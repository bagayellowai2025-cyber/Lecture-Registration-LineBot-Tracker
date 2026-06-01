/**
 * 設定「表單送出時」觸發器（只要執行一次）
 */
function setupFormSubmitTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'handleFormSubmit') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger('handleFormSubmit')
    .forSpreadsheet(SpreadsheetApp.getActive())
    .onFormSubmit()
    .create();
}

/**
 * 取得「狀態」與「序號」所在欄位
 */
function getStatusAndOrderColumns(sheet) {
  const HEADER_ROW = 1;
  const lastCol = sheet.getLastColumn();
  const headerRange = sheet.getRange(HEADER_ROW, 1, 1, lastCol);
  const headers = headerRange.getValues()[0];

  let statusCol = null;
  let orderCol  = null;

  headers.forEach((title, index) => {
    if (title === '狀態') statusCol = index + 1; 
    else if (title === '序號') orderCol = index + 1;
  });

  if (statusCol && orderCol) return { STATUS_COL: statusCol, ORDER_COL: orderCol };

  let firstEmptyCol = null;
  for (let i = 0; i < headers.length; i++) {
    if (headers[i] === '' || headers[i] === null) {
      firstEmptyCol = i + 1;
      break;
    }
  }

  if (!firstEmptyCol) firstEmptyCol = lastCol + 1;

  if (!statusCol) {
    statusCol = firstEmptyCol;
    sheet.getRange(HEADER_ROW, statusCol).setValue('狀態');
  }

  if (!orderCol) {
    let col = statusCol + 1;
    let currentLastCol = sheet.getLastColumn();
    if (col <= currentLastCol) {
      let title = sheet.getRange(HEADER_ROW, col).getValue();
      if (title !== '' && title !== null) col = currentLastCol + 1;
    }
    orderCol = col;
    sheet.getRange(HEADER_ROW, orderCol).setValue('序號');
  }

  return { STATUS_COL: statusCol, ORDER_COL: orderCol };
}

/**
 * 實際處理「正取 / 備取」邏輯的函式
 */
function handleFormSubmit(e) {
  if (!e || !e.range) return;

  const sheet = e.range.getSheet();
  const HEADER_ROW = 1; 
  const LIMIT = 3;       

  const lastRow    = sheet.getLastRow();
  const currentRow = e.range.getRow();
  const { STATUS_COL, ORDER_COL } = getStatusAndOrderColumns(sheet);
  const numRows = lastRow - HEADER_ROW;
  if (numRows <= 0) return;

  const statusValues = sheet.getRange(HEADER_ROW + 1, STATUS_COL, numRows, 1).getValues();

  let acceptedCount = 0; 
  let waitlistCount = 0; 

  statusValues.forEach(row => {
    const status = row[0];
    if (status === '正取') acceptedCount++;
    else if (status === '備取') waitlistCount++;
  });

  if (acceptedCount < LIMIT) {
    sheet.getRange(currentRow, STATUS_COL).setValue('正取');
    sheet.getRange(currentRow, ORDER_COL).setValue(acceptedCount + 1); 
  } else {
    sheet.getRange(currentRow, STATUS_COL).setValue('備取');
    sheet.getRange(currentRow, ORDER_COL).setValue(waitlistCount + 1); 
  }
}

/**
 * 對現有所有回覆重新計算 正取 / 備取
 */
function recalcAllStatus() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const HEADER_ROW = 1;
  const LIMIT = 3; 

  const { STATUS_COL, ORDER_COL } = getStatusAndOrderColumns(sheet);
  const lastRow = sheet.getLastRow();
  const numRows = lastRow - HEADER_ROW;
  if (numRows <= 0) return;

  sheet.getRange(HEADER_ROW + 1, STATUS_COL, numRows, 1).clearContent();
  sheet.getRange(HEADER_ROW + 1, ORDER_COL,  numRows, 1).clearContent();

  let acceptedCount = 0;
  let waitlistCount = 0;

  for (let i = HEADER_ROW + 1; i <= lastRow; i++) {
    if (acceptedCount < LIMIT) {
      acceptedCount++;
      sheet.getRange(i, STATUS_COL).setValue('正取');
      sheet.getRange(i, ORDER_COL).setValue(acceptedCount);
    } else {
      waitlistCount++;
      sheet.getRange(i, STATUS_COL).setValue('備取');
      sheet.getRange(i, ORDER_COL).setValue(waitlistCount);
    }
  }
}

// ==========================================
// 全域設定區
// ==========================================
const LINE_ACCESS_TOKEN = 'YOUR_LINE_ACCESS_TOKEN';

// ==========================================
// 1. LINE Bot 核心處理與 Log 日誌區
// ==========================================

function doPost(e) {
  if (!e || !e.postData || !e.postData.contents) return ContentService.createTextOutput("OK");

  try {
    const eventData = JSON.parse(e.postData.contents);
    const events = eventData.events;
    
    if (events && events.length > 0) {
      const event = events[0];
      const replyToken = event.replyToken;
      
      if (event.type === 'message' && event.message.type === 'text') {
        // 為了避免大小寫問題，將使用者輸入全部轉大寫來比對
        const userMessage = event.message.text.toUpperCase(); 

        // 🌟 新增判斷：如果使用者輸入包含 VIP 🌟
        if (userMessage.includes('VIP')) {
          const resultText = getVipStatus();
          replyLineMessage(replyToken, resultText);
        } 
        // 原本判斷：人數、狀態、正取
        else if (userMessage.includes('人數') || userMessage.includes('狀態') || userMessage.includes('正取')) {
          const resultText = getRegistrationStatus();
          replyLineMessage(replyToken, resultText);
        }
      }
    }
  } catch(err) {
    console.error("【嚴重錯誤】doPost 執行期間出錯: " + err.toString());
  }
  
  return ContentService.createTextOutput("OK");
}

/**
 * 🌟 新增函式：讀取試算表並回傳 VIP 資訊 🌟
 */
function getVipStatus() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const HEADER_ROW = 1;
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();

  if (lastRow <= HEADER_ROW) {
    return "目前還沒有任何人報名喔！";
  }

  // 取得整張表的資料（包含標題列）
  const data = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  const headers = data[0];

  // 找出我們需要的欄位所在的索引 (Index 找不到時回傳 -1)
  const timeIdx = headers.indexOf('時間戳記');
  const nameIdx = headers.indexOf('姓名');
  const dogNameIdx = headers.indexOf('狗狗名字');
  const inviteCodeIdx = headers.indexOf('邀請碼');
  const statusIdx = headers.indexOf('狀態');
  const orderIdx = headers.indexOf('序號');

  if (inviteCodeIdx === -1) {
    return "【錯誤】找不到「邀請碼」欄位，無法判斷 VIP 身份。";
  }

  let vipList = [];

  // 從第二列開始往下尋找
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const inviteCode = String(row[inviteCodeIdx]).trim().toUpperCase();

    // 如果這列的邀請碼是 VIP
    if (inviteCode === 'VIP') {
      
      // 處理日期格式
      let dateStr = row[timeIdx];
      if (dateStr instanceof Date) {
        // 如果是 JS Date 物件，轉成好看的格式 (例如: 6/1 16:42)
        dateStr = Utilities.formatDate(dateStr, Session.getScriptTimeZone(), "MM/dd HH:mm");
      }

      // 提取資料 (加上防呆，避免某些表單沒填到資料)
      const name = nameIdx !== -1 ? row[nameIdx] : "未知";
      const dogName = dogNameIdx !== -1 ? row[dogNameIdx] : "無";
      const status = statusIdx !== -1 ? row[statusIdx] : "";
      const order = orderIdx !== -1 ? row[orderIdx] : "";

      // 組合字串格式：報名日期 | 姓名 | 狗狗名字 | 狀態+序號
      vipList.push(`🐾 ${dateStr}｜${name}｜狗狗:${dogName}｜${status}${order}`);
    }
  }

  // 判斷是否有找到 VIP
  if (vipList.length === 0) {
    return "目前還沒有 VIP 報名喔！";
  }

  // 組合最後傳送給 LineBot 的文字
  return "🌟 目前 VIP 報名名單 🌟\n\n" + vipList.join('\n');
}

/**
 * 讀取試算表並計算正取與備取人數
 */
function getRegistrationStatus() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const HEADER_ROW = 1;
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();

  if (lastRow <= HEADER_ROW) {
    return "目前還沒有任何人報名喔！";
  }

  const headers = sheet.getRange(HEADER_ROW, 1, 1, lastCol).getValues()[0];
  const statusColIndex = headers.indexOf('狀態');

  if (statusColIndex === -1) {
    return "找不到「狀態」欄位，請確認試算表格式或等待第一筆報名資料寫入。";
  }

  const statusColNumber = statusColIndex + 1; 
  const statusValues = sheet.getRange(HEADER_ROW + 1, statusColNumber, lastRow - HEADER_ROW, 1).getValues();

  let acceptedCount = 0;
  let waitlistCount = 0;

  statusValues.forEach((row) => {
    const status = row[0];
    if (status === '正取') acceptedCount++;
    if (status === '備取') waitlistCount++;
  });

  let extraMessage = "";
  if (waitlistCount >= 5) {
    extraMessage = "\n\n💡 提示：備取人數偏多，或許可以考慮加開場次囉！";
  }

  return `📊 目前報名狀況\n✅ 正取：${acceptedCount} 人\n⏳ 備取：${waitlistCount} 人` + extraMessage;
}

/**
 * 呼叫 LINE API 回傳訊息
 */
function replyLineMessage(replyToken, text) {
  const url = 'https://api.line.me/v2/bot/message/reply';
  const options = {
    'method': 'post',
    'headers': {
      'Content-Type': 'application/json',
      'Authorization': 'Bearer ' + LINE_ACCESS_TOKEN
    },
    'payload': JSON.stringify({
      'replyToken': replyToken,
      'messages': [{
        'type': 'text',
        'text': text
      }]
    }),
    'muteHttpExceptions': true 
  };
  
  try {
    UrlFetchApp.fetch(url, options);
  } catch(e) {
    console.error("【連線錯誤】無法呼叫 LINE API: " + e.toString());
  }
}

function testAuth() {
  UrlFetchApp.fetch('https://www.google.com');
}

// ==========================================
// 2. 🌟 新增：定時主動推播功能 🌟
// ==========================================

// ⚠️ 請在此處填入你步驟一取得的 LINE User ID (或是群組的 Group ID)
const LINE_USER_ID = 'YOUR_LINE_USER_ID'; 

/**
 * 核心定時推播函式 (整合報名人數與 VIP 清單)
 */
function sendPeriodicUpdate() {
  // 1. 取得正取與備取人數
  const registrationStatus = getRegistrationStatus();
  
  // 2. 取得 VIP 名單
  const vipStatus = getVipStatus();
  
  // 3. 組合文字
  const finalMessage = `🔔【講座報名定時回報】\n\n${registrationStatus}\n\n================\n\n${vipStatus}`;
  
  // 4. 呼叫 LINE Push API 主動推播
  pushLineMessage(LINE_USER_ID, finalMessage);
}

/**
 * 呼叫 LINE API 主動推播訊息 (Push Message)
 */
function pushLineMessage(toId, text) {
  const url = 'https://api.line.me/v2/bot/message/push';
  const options = {
    'method': 'post',
    'headers': {
      'Content-Type': 'application/json',
      'Authorization': 'Bearer ' + LINE_ACCESS_TOKEN
    },
    'payload': JSON.stringify({
      'to': toId,           // 接收者的 ID
      'messages': [{
        'type': 'text',
        'text': text
      }]
    }),
    'muteHttpExceptions': true 
  };
  
  try {
    const response = UrlFetchApp.fetch(url, options);
    console.log("推播結果: " + response.getContentText());
  } catch(e) {
    console.error("【推播連線錯誤】無法呼叫 LINE Push API: " + e.toString());
  }
}

/**
 * 設定「每12小時」執行的觸發器 (只需要手動執行一次此函式即可)
 */
function setupPeriodicTrigger() {
  const functionName = 'sendPeriodicUpdate';
  const triggers = ScriptApp.getProjectTriggers();
  
  // 清除舊的同名觸發器，避免重複設定
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === functionName) {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  // 建立每 12 小時執行一次的觸發器
  ScriptApp.newTrigger(functionName)
    .timeBased()
    .everyHours(12)
    .create();
    
  console.log("✅ 成功建立每 12 小時更新觸發器！");
}
