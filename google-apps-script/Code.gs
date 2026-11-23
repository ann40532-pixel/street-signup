/**
 * 地推登記 → Google 試算表 同步接收端
 * ------------------------------------------------------------
 * 安裝方式：
 *   1. 打開試算表 → 擴充功能 → Apps Script
 *   2. 把編輯器裡原本的內容全部刪掉，貼上這整個檔案 → 儲存
 *   3. 右上角「部署」→「新增部署作業」→ 齒輪選「網頁應用程式」
 *        執行身分：我
 *        誰可以存取：所有人
 *      → 部署 →（第一次會要求授權，按允許）→ 複製「網頁應用程式網址」
 *   4. 在地推登記 App 的「雲端同步」設定裡貼上這個網址
 *
 * 之後如果修改這份程式碼：部署 → 管理部署作業 → 編輯（鉛筆）
 *   → 版本選「新版本」→ 部署。這樣網址不會變，裝置不用重新設定。
 *
 * 資料格式：每場活動（日期 + 地點）一張工作表，一個 email 一列。
 *   以最右邊的 ID 欄辨識每一筆，重複傳送也不會產生重複列；
 *   App 裡編輯或刪除，試算表也會跟著改。
 */

var HEADERS = ['Email', '登記時間', '最後更新', '登記裝置', 'ID'];
var COL_ID = 5;

// 用瀏覽器打開網址時會看到這個，App 的「測試連線」也是用這個
function doGet() {
  return json_({ ok: true, app: 'street-signup' });
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000); // 多台裝置同時上傳時排隊處理，避免互相覆蓋
  } catch (err) {
    return json_({ ok: false, error: '伺服器忙碌中，稍後自動重試' });
  }
  try {
    var data = JSON.parse(e.postData.contents);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var device = safeText_(String(data.device || '').slice(0, 50));
    (data.events || []).forEach(function (ev) { syncEvent_(ss, ev, device); });
    SpreadsheetApp.flush();
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String((err && err.message) || err) });
  } finally {
    lock.releaseLock();
  }
}

function syncEvent_(ss, ev, device) {
  if (!ev || !ev.date || !ev.place) return;
  var sh = getOrCreateSheet_(ss, sheetName_(ev.date, ev.place));

  // 讀出現有的 ID → 列號
  var last = sh.getLastRow();
  var rowOf = {};
  if (last > 1) {
    sh.getRange(2, COL_ID, last - 1, 1).getValues().forEach(function (r, i) {
      if (r[0] !== '') rowOf[String(r[0])] = i + 2;
    });
  }

  var now = new Date();
  var appends = [];
  (ev.upserts || []).forEach(function (u) {
    if (!u || !u.id || !u.email) return;
    var id = String(u.id);
    var email = safeText_(String(u.email).slice(0, 254));
    if (rowOf[id]) {                       // 已存在 → 更新 email
      sh.getRange(rowOf[id], 1).setValue(email);
      sh.getRange(rowOf[id], 3).setValue(now);
    } else if (!(id in rowOf)) {           // 新的一筆 → 加到最後
      var created = new Date(Number(u.createdAt) || now.getTime());
      appends.push([email, created, now, device, id]);
      rowOf[id] = 0;
    }
  });
  if (appends.length) {
    var need = sh.getLastRow() + appends.length - sh.getMaxRows();
    if (need > 0) sh.insertRowsAfter(sh.getMaxRows(), need);
    sh.getRange(sh.getLastRow() + 1, 1, appends.length, HEADERS.length).setValues(appends);
  }

  // 刪除：由下往上刪，列號才不會跑掉
  var del = [];
  (ev.deletes || []).forEach(function (id) {
    var r = rowOf[String(id)];
    if (r) del.push(r);
  });
  if (del.length) {
    // Google 試算表不允許刪光所有非凍結列，先補空白列
    if (sh.getMaxRows() - del.length <= 1) sh.insertRowsAfter(sh.getMaxRows(), del.length);
    del.sort(function (a, b) { return b - a; }).forEach(function (r) { sh.deleteRow(r); });
  }
}

function getOrCreateSheet_(ss, name) {
  var sh = ss.getSheetByName(name);
  if (sh) return sh;
  sh = ss.insertSheet(name, ss.getNumSheets());
  sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS])
    .setFontWeight('bold').setBackground('#e3f3f1');
  sh.setFrozenRows(1);
  sh.setColumnWidth(1, 260);
  sh.setColumnWidths(2, 2, 150);
  sh.setColumnWidth(4, 140);
  sh.setColumnWidth(5, 120);
  sh.getRange('B:C').setNumberFormat('yyyy/mm/dd hh:mm:ss');
  sh.getRange('E:E').setFontColor('#9e9e9e');
  if (sh.getMaxColumns() > HEADERS.length) {
    sh.deleteColumns(HEADERS.length + 1, sh.getMaxColumns() - HEADERS.length);
  }
  return sh;
}

// 工作表名稱：「2026-11-23 台北車站東三門」
function sheetName_(date, place) {
  var name = (String(date) + ' ' + String(place))
    .replace(/[:\\\/?*\[\]]/g, '-')
    .replace(/\s+/g, ' ')
    .trim();
  return name.slice(0, 99) || '未命名活動';
}

// 避免以 = + - @ 開頭的內容被試算表當成公式
function safeText_(s) {
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
