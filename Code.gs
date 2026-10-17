/**
 * 瓦斯行派工 + LINE 通知系統 — 後端 (Code.gs)
 *
 * 試算表結構（執行選單「🔥 瓦斯派工 > 初始化工作表」會自動建立）：
 *   Customers：電話 | 姓名 | 地址 | 備註 | 建立時間 | 最後訂購
 *   Orders   ：訂單編號 | 建立時間 | 客戶電話 | 客戶姓名 | 送貨地址 | 品項明細 | 總桶數 | 指定時間 | 指派工人 | 備註 | 狀態 | LINE通知
 *   Workers  ：工人姓名 | LINE User ID | 電話 | 啟用(Y/N)
 *
 * 指令碼屬性（專案設定 > 指令碼屬性）：
 *   LINE_CHANNEL_ACCESS_TOKEN  必填，LINE Messaging API 的長期 Channel access token
 *   SPREADSHEET_ID             選填，若程式不是從試算表「擴充功能 > Apps Script」建立的才需要
 */

const CONFIG = {
  TZ: 'Asia/Taipei',
  SHEETS: {
    CUSTOMERS: 'Customers',
    ORDERS: 'Orders',
    WORKERS: 'Workers',
    LINE_USERS: 'LineUsers',
  },
  HEADERS: {
    Customers: ['電話', '姓名', '地址', '備註', '建立時間', '最後訂購'],
    Orders: ['訂單編號', '建立時間', '客戶電話', '客戶姓名', '送貨地址', '品項明細', '總桶數', '指定時間', '指派工人', '備註', '狀態', 'LINE通知'],
    Workers: ['工人姓名', 'LINE User ID', '電話', '啟用(Y/N)'],
  },
  // 表單上的瓦斯品項，可自行增減
  ITEMS: ['20公斤', '16公斤', '10公斤', '4公斤'],
};

// Orders 欄位位置（從 1 開始）
const ORDER_COL = { WORKER: 9, NOTIFY: 12 };

/* ========================= 網頁入口 ========================= */

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('瓦斯行快速接單')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** 前端載入時取得工人清單與品項 */
function getInitData() {
  return {
    workers: getActiveWorkers_().map(w => w.name),
    items: CONFIG.ITEMS,
  };
}

/**
 * 依電話查詢客戶。
 * exact：電話完全相符的客戶；matches：電話包含輸入數字的客戶（最多 8 筆）。
 */
function searchCustomers(query) {
  const digits = normPhone_(query);
  if (digits.length < 3) return { exact: null, matches: [] };

  const list = readCustomers_();
  const toClient = c => ({ phone: c.phone, name: c.name, address: c.address });
  const exact = list.find(c => c.phone === digits);
  const matches = list.filter(c => c.phone.indexOf(digits) !== -1).slice(0, 8);

  return { exact: exact ? toClient(exact) : null, matches: matches.map(toClient) };
}

/** 前端「送出並派工」：寫入 Orders、更新 Customers、LINE 通知工人 */
function submitOrder(form) {
  const phone = normPhone_(form.phone);
  const name = String(form.name || '').trim();
  const address = String(form.address || '').trim();
  const workerName = String(form.worker || '').trim();
  const note = String(form.note || '').trim();
  const items = (form.items || [])
    .map(it => ({ name: String(it.name), qty: parseInt(it.qty, 10) || 0 }))
    .filter(it => it.qty > 0 && CONFIG.ITEMS.indexOf(it.name) !== -1);

  if (!phone) throw new Error('請輸入客戶電話');
  if (!name) throw new Error('請輸入客戶姓名');
  if (!address) throw new Error('請輸入送貨地址');
  if (!items.length) throw new Error('請至少選擇一項瓦斯品項');

  const worker = getActiveWorkers_().find(w => w.name === workerName);
  if (!worker) throw new Error('找不到工人「' + workerName + '」，請重新整理頁面');

  const order = {
    phone: phone,
    name: name,
    address: address,
    itemText: items.map(it => it.name + ' × ' + it.qty).join('、'),
    total: items.reduce((sum, it) => sum + it.qty, 0),
    timeText: formatSchedule_(form.schedule),
    worker: worker.name,
    note: note,
  };

  // 鎖定避免兩人同時送單時訂單編號重複
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  let row;
  try {
    order.orderId = nextOrderId_();
    const sh = getSheet_(CONFIG.SHEETS.ORDERS);
    sh.appendRow([
      order.orderId, new Date(), asText_(phone), cell_(name), cell_(address),
      order.itemText, order.total, order.timeText, worker.name, cell_(note), '已派工', '',
    ]);
    row = sh.getLastRow();
    upsertCustomer_(phone, name, address);
  } finally {
    lock.releaseLock();
  }

  // LINE 通知失敗不影響訂單成立，結果寫回「LINE通知」欄
  const result = notifyWorker_(worker, order);
  getSheet_(CONFIG.SHEETS.ORDERS).getRange(row, ORDER_COL.NOTIFY)
    .setValue(result.ok ? '已通知 ' + nowText_() : '失敗：' + result.error);

  return { orderId: order.orderId, worker: worker.name, notified: result.ok, error: result.error || '' };
}

/* ========================= 試算表選單 ========================= */

function onOpen() {
  SpreadsheetApp.getUi().createMenu('🔥 瓦斯派工')
    .addItem('初始化工作表', 'setupSheets')
    .addItem('測試 LINE 推播（發給第一位工人）', 'testLinePush')
    .addItem('重送選取訂單的 LINE 通知', 'resendSelectedOrder')
    .addToUi();
}

/** 建立三個工作表與標題列（已存在的不會覆蓋） */
function setupSheets() {
  const ss = getSs_();
  Object.keys(CONFIG.HEADERS).forEach(name => {
    const sh = ss.getSheetByName(name) || ss.insertSheet(name);
    if (sh.getLastRow() === 0) {
      const headers = CONFIG.HEADERS[name];
      sh.getRange(1, 1, 1, headers.length).setValues([headers])
        .setFontWeight('bold').setBackground('#fde68a');
      sh.setFrozenRows(1);
    }
  });

  const workers = ss.getSheetByName(CONFIG.SHEETS.WORKERS);
  if (workers.getLastRow() === 1) {
    workers.appendRow(['範例工人(請改名)', '請填入 U 開頭的 LINE User ID', asText_('0912000000'), 'Y']);
  }
  try {
    SpreadsheetApp.getUi().alert('工作表已就緒！請到 Workers 填入工人資料。');
  } catch (err) {
    // 從 Apps Script 編輯器直接執行時沒有試算表畫面可以跳提示，略過即可
  }
}

function testLinePush() {
  const ui = SpreadsheetApp.getUi();
  const worker = getActiveWorkers_()[0];
  if (!worker) return ui.alert('Workers 工作表中沒有啟用的工人');

  const result = notifyWorker_(worker, {
    orderId: 'TEST', name: '測試客戶', phone: '0912345678', address: '台北市信義區市府路1號',
    itemText: '20公斤 × 1', total: 1, timeText: '盡快送達', worker: worker.name, note: '這是測試訊息',
  });
  ui.alert(result.ok ? '✅ 已發送測試訊息給「' + worker.name + '」' : '❌ 發送失敗：' + result.error);
}

/** 在 Orders 選取某一列後執行，重新發送該筆訂單的 LINE 通知 */
function resendSelectedOrder() {
  const ui = SpreadsheetApp.getUi();
  const sh = SpreadsheetApp.getActiveSheet();
  const row = sh.getActiveRange().getRow();
  if (sh.getName() !== CONFIG.SHEETS.ORDERS || row < 2) {
    return ui.alert('請先在 Orders 工作表點選要重送的那一列訂單');
  }

  const v = sh.getRange(row, 1, 1, CONFIG.HEADERS.Orders.length).getDisplayValues()[0];
  const worker = getActiveWorkers_().find(w => w.name === v[ORDER_COL.WORKER - 1]);
  if (!worker) return ui.alert('找不到工人「' + v[ORDER_COL.WORKER - 1] + '」（或已停用）');

  const result = notifyWorker_(worker, {
    orderId: v[0], phone: v[2], name: v[3], address: v[4], itemText: v[5],
    total: v[6], timeText: v[7], worker: worker.name, note: v[9],
  });
  sh.getRange(row, ORDER_COL.NOTIFY).setValue(result.ok ? '已重送 ' + nowText_() : '失敗：' + result.error);
  ui.alert(result.ok ? '✅ 已重新通知「' + worker.name + '」' : '❌ 發送失敗：' + result.error);
}

/* ========================= LINE Messaging API ========================= */

function notifyWorker_(worker, order) {
  if (!/^U[0-9a-f]{32}$/.test(worker.lineId)) {
    return { ok: false, error: '工人「' + worker.name + '」的 LINE User ID 未設定或格式錯誤' };
  }
  return callLineApi_('message/push', {
    to: worker.lineId,
    messages: [{ type: 'text', text: buildMessage_(order) }],
  });
}

function buildMessage_(o) {
  const mapUrl = 'https://www.google.com/maps/search/?api=1&query=' + encodeURIComponent(o.address);
  return [
    '🔥 新派工單 ' + o.orderId,
    '━━━━━━━━━━━━',
    '👤 客戶：' + o.name,
    '📞 電話：' + o.phone,
    '📍 地址：' + o.address,
    '🛢️ 品項：' + o.itemText + '（共 ' + o.total + ' 桶）',
    '⏰ 時間：' + o.timeText,
    '📝 備註：' + (o.note || '無'),
    '━━━━━━━━━━━━',
    '🗺️ 導航：' + mapUrl,
  ].join('\n');
}

/** 呼叫 LINE API，回傳 { ok, error } 而不丟出例外 */
function callLineApi_(path, payload) {
  const token = PropertiesService.getScriptProperties().getProperty('LINE_CHANNEL_ACCESS_TOKEN');
  if (!token) return { ok: false, error: '尚未設定指令碼屬性 LINE_CHANNEL_ACCESS_TOKEN' };

  try {
    const res = UrlFetchApp.fetch('https://api.line.me/v2/bot/' + path, {
      method: 'post',
      contentType: 'application/json',
      headers: { Authorization: 'Bearer ' + token },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true,
    });
    const code = res.getResponseCode();
    if (code === 200) return { ok: true };
    return { ok: false, error: 'LINE API ' + code + '：' + res.getContentText().slice(0, 200) };
  } catch (err) {
    return { ok: false, error: String(err.message || err) };
  }
}

/**
 * LINE Webhook：工人加官方帳號好友並傳任何訊息後，
 * 機器人會回覆他的 User ID，同時記錄到 LineUsers 工作表，方便老闆複製到 Workers。
 * （只在「取得工人 User ID」時需要，詳見 README）
 */
function doPost(e) {
  const body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
  (body.events || []).forEach(ev => {
    const userId = ev.source && ev.source.userId;
    if (!userId || (ev.type !== 'message' && ev.type !== 'follow')) return;

    recordLineUser_(userId);
    if (ev.replyToken) {
      callLineApi_('message/reply', {
        replyToken: ev.replyToken,
        messages: [{ type: 'text', text: '您的 LINE User ID 是：\n' + userId + '\n\n老闆已可在系統中看到，等待設定完成即可收到派工通知。' }],
      });
    }
  });
  return ContentService.createTextOutput('OK');
}

function recordLineUser_(userId) {
  const ss = getSs_();
  let sh = ss.getSheetByName(CONFIG.SHEETS.LINE_USERS);
  if (!sh) {
    sh = ss.insertSheet(CONFIG.SHEETS.LINE_USERS);
    sh.appendRow(['LINE 顯示名稱', 'LINE User ID', '首次記錄時間']);
    sh.getRange(1, 1, 1, 3).setFontWeight('bold').setBackground('#fde68a');
  }
  const ids = sh.getRange(1, 2, sh.getLastRow(), 1).getValues().map(r => r[0]);
  if (ids.indexOf(userId) !== -1) return;

  let displayName = '';
  const token = PropertiesService.getScriptProperties().getProperty('LINE_CHANNEL_ACCESS_TOKEN');
  try {
    const res = UrlFetchApp.fetch('https://api.line.me/v2/bot/profile/' + userId, {
      headers: { Authorization: 'Bearer ' + token },
      muteHttpExceptions: true,
    });
    if (res.getResponseCode() === 200) displayName = JSON.parse(res.getContentText()).displayName;
  } catch (err) {
    // 取不到名稱沒關係，User ID 才是重點
  }
  sh.appendRow([cell_(displayName), userId, new Date()]);
}

/* ========================= 資料存取 ========================= */

function getSs_() {
  const id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  return id ? SpreadsheetApp.openById(id) : SpreadsheetApp.getActiveSpreadsheet();
}

function getSheet_(name) {
  const sh = getSs_().getSheetByName(name);
  if (!sh) throw new Error('找不到工作表「' + name + '」，請先在試算表選單執行「🔥 瓦斯派工 > 初始化工作表」');
  return sh;
}

function getActiveWorkers_() {
  const sh = getSheet_(CONFIG.SHEETS.WORKERS);
  if (sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, 4).getDisplayValues()
    .map(r => ({
      name: r[0].trim(),
      lineId: r[1].trim(),
      phone: r[2].trim(),
      active: r[3].trim().toUpperCase() !== 'N',
    }))
    .filter(w => w.name && w.active);
}

function readCustomers_() {
  const sh = getSheet_(CONFIG.SHEETS.CUSTOMERS);
  if (sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, 3).getDisplayValues()
    .map((r, i) => ({ row: i + 2, phone: normPhone_(r[0]), name: r[1], address: r[2] }))
    .filter(c => c.phone);
}

/** 新客戶新增一列；老客戶更新為最新的姓名、地址與最後訂購時間 */
function upsertCustomer_(phone, name, address) {
  const sh = getSheet_(CONFIG.SHEETS.CUSTOMERS);
  const now = new Date();
  const found = readCustomers_().find(c => c.phone === phone);
  if (found) {
    sh.getRange(found.row, 2, 1, 2).setValues([[cell_(name), cell_(address)]]);
    sh.getRange(found.row, 6).setValue(now);
  } else {
    sh.appendRow([asText_(phone), cell_(name), cell_(address), '', now, now]);
  }
}

/** 訂單編號：yyyyMMdd-001，每天從 001 重新編號 */
function nextOrderId_() {
  const props = PropertiesService.getScriptProperties();
  const today = Utilities.formatDate(new Date(), CONFIG.TZ, 'yyyyMMdd');
  const saved = (props.getProperty('ORDER_SEQ') || '').split(':');
  const seq = saved[0] === today ? Number(saved[1]) + 1 : 1;
  props.setProperty('ORDER_SEQ', today + ':' + seq);
  return today + '-' + ('00' + seq).slice(-3);
}

/* ========================= 小工具 ========================= */

/** 只留數字；試算表把 09 開頭手機存成數字而少了 0 的情況一併修正 */
function normPhone_(s) {
  const digits = String(s || '').replace(/\D/g, '');
  return digits.length === 9 && digits.charAt(0) === '9' ? '0' + digits : digits;
}

/** 前置單引號讓試算表當文字存，保留電話開頭的 0 */
function asText_(s) {
  return "'" + s;
}

/** 防止使用者輸入的 = + - @ 開頭內容被試算表當成公式 */
function cell_(s) {
  s = String(s);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

/** 前端 datetime-local 值（2026-10-17T14:30）→「10/17(六) 14:30」；空白為「盡快送達」 */
function formatSchedule_(value) {
  const m = String(value || '').match(/^(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2})/);
  if (!m) return '盡快送達';
  const weekday = '日一二三四五六'.charAt(new Date(+m[1], +m[2] - 1, +m[3]).getDay());
  return m[2] + '/' + m[3] + '(' + weekday + ') ' + m[4] + ':' + m[5];
}

function nowText_() {
  return Utilities.formatDate(new Date(), CONFIG.TZ, 'MM/dd HH:mm');
}
