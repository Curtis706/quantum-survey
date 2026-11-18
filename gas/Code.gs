/**
 * 量子科技跨界交流｜EMBA 參與者背景調查系統
 * Google Apps Script 後端（綁定於 Google 試算表）
 *
 * 公開函式（填表頁可呼叫）：doGet、getFormConfig、submitResponse
 * 主辦人函式（皆有權限檢查，只限試算表擁有者／編輯者）：
 *   setupSystem、openDashboard、openExport、getDashboardData、refreshReportSheets、
 *   exportExcel、exportCsv、applyQuestionSettings、deleteIdentifiableData
 * 名稱以「_」結尾的函式為內部函式，前端無法呼叫。
 */

var APP = {
  TITLE: '量子電腦的未來｜EMBA 參與者背景調查',
  TZ: 'Asia/Taipei',
  SHEET_DATA: '填寫資料',
  SHEET_CONFIG: '題目設定',
  SHEET_STATS: '統計分析',
  SHEET_SUMMARY: '講者摘要',
  MAX_TOTAL: 500,              // 總筆數上限
  MAX_PER_MINUTE: 20,          // 全站每分鐘送出上限
  MIN_FILL_MS: 5000,           // 開啟頁面 5 秒內送出視為機器人
  TOKEN_TTL_MS: 12 * 60 * 60 * 1000, // 頁面憑證 12 小時有效
  MULTI_SEP: '；',
  CONSENT_VALUE: '已同意',
  REPLACED_MARK: '已由較新填寫取代',
  // GitHub Pages 網址（填表頁與統計頁）。改用其他網址時修改這裡。
  PAGES_URL: 'https://curtis706.github.io/quantum-survey/',
  BRAND_NAVY: '#0B1F3A'
};

var HEADERS = [
  '填寫時間', '姓名', '所屬產業', '產業（其他說明）', '公司名稱', '職務類型', '實際職稱',
  '大學主要領域', '畢業學校', '就讀科系', '量子科技熟悉程度', '感興趣的主題',
  '主題（其他說明）', '同意資料使用說明', '送出編號', '備註'
];

var COL = {
  TIME: 0, NAME: 1, IND: 2, IND_OTHER: 3, COMPANY: 4, ROLE: 5, TITLE: 6, EDU: 7,
  SCHOOL: 8, DEPT: 9, FAM: 10, TOPICS: 11, TOPIC_OTHER: 12, CONSENT: 13, SID: 14, NOTE: 15
};

// 題目代碼 → 「填寫資料」欄位
var FIELD_MAP = {
  Q1: { col: COL.NAME },
  Q2: { col: COL.IND, otherCol: COL.IND_OTHER },
  Q2_COMPANY: { col: COL.COMPANY },
  Q3: { col: COL.ROLE },
  Q3_TITLE: { col: COL.TITLE },
  Q4: { col: COL.EDU },
  Q4_SCHOOL: { col: COL.SCHOOL },
  Q4_DEPT: { col: COL.DEPT },
  Q5: { col: COL.FAM },
  Q6: { col: COL.TOPICS, otherCol: COL.TOPIC_OTHER }
};

// 預設題目（「題目設定」工作表可覆寫：題目文字、說明文字、選項、「其他」文字框）
var DEFAULT_QUESTIONS = [
  { code: 'Q1', type: 'text', required: true, max: 50,
    title: '您的姓名', help: '請填寫方便活動主辦人辨識的姓名。', placeholder: '例如：王小明／Ming Wang' },
  { code: 'Q2', type: 'single', required: true, otherText: true,
    title: '您目前主要從事的產業？', help: '',
    options: ['科技／半導體／電子', '傳統製造業', '金融／投資／保險', '醫療／生技', '服務／零售／餐飲',
      '建築／營造／不動產', '資訊／軟體／AI', '教育／學術研究', '其他'] },
  { code: 'Q2_COMPANY', parent: 'Q2', type: 'text', required: false, max: 80,
    title: '目前任職公司名稱', help: '如方便分享，歡迎提供公司名稱，讓講者更了解參與者的產業背景。', placeholder: '選填' },
  { code: 'Q3', type: 'single', required: true, otherText: false,
    title: '您目前的主要職務類型？', help: '',
    options: ['企業負責人／創辦人', '董事長／董事／高階主管', '專業經理人', '工程／技術／研發人員',
      '金融／投資專業人士', '學術／研究人員', '其他'] },
  { code: 'Q3_TITLE', parent: 'Q3', type: 'text', required: false, max: 80,
    title: '目前職稱', help: '例如：總經理、執行長、營運長、研發主管等。', placeholder: '選填' },
  { code: 'Q4', type: 'single', required: true, otherText: false,
    title: '您大學主要就讀的專業領域？',
    help: '張博士希望了解參與者是否具有理工相關的教育背景，以便判斷分享內容可以深入到什麼程度。基於尊重個人隱私，學校名稱與詳細科系不強制填寫。',
    options: ['理工／資訊相關科系', '商學／管理相關科系', '醫學／生命科學相關科系', '人文／社會科學相關科系',
      '設計／藝術相關科系', '其他'] },
  { code: 'Q4_SCHOOL', parent: 'Q4', type: 'text', required: false, max: 80,
    title: '大學畢業學校', help: '', placeholder: '選填' },
  { code: 'Q4_DEPT', parent: 'Q4', type: 'text', required: false, max: 80,
    title: '大學就讀科系', help: '', placeholder: '選填' },
  { code: 'Q5', type: 'single', required: true, otherText: false,
    title: '您目前對量子電腦的了解程度？', help: '',
    options: ['完全不了解，但非常好奇', '曾經聽過，大致知道一些概念', '了解基本原理及可能的應用', '具有相關理工、科技或研究背景'] },
  { code: 'Q6', type: 'multi', required: true, otherText: true,
    title: '您最希望張博士分享哪些內容？', help: '可複選，至少選擇一項。',
    options: ['量子電腦的基本原理，以及與傳統電腦的差異', '量子電腦與 AI 之間的關係', '全球量子科技目前的發展現況',
      '量子電腦未來可能帶來哪些產業革命', '台灣科技產業及製造業的參與機會', '量子科技的商業模式與投資前景',
      '從美國頂尖科學家走向科技創業者的經驗', '其他'] }
];

/* =========================================================
 * 公開：網頁與填表
 * ========================================================= */

function doGet(e) {
  var action = e && e.parameter ? e.parameter.action : '';
  if (action === 'config') return json_(getFormConfig());
  if (action === 'stats') return json_(getPublicStats_(e.parameter.key));
  // 填表頁架在 GitHub Pages；直接開啟 /exec 網址時顯示前往連結
  return HtmlService.createHtmlOutput(
    '<div style="font-family:sans-serif;text-align:center;padding:60px 20px;line-height:1.8">' +
    '<h2>量子電腦的未來｜EMBA 參與者背景調查</h2>' +
    '<p><a href="' + APP.PAGES_URL + '" target="_top" style="font-size:18px">點此前往填寫問卷 →</a></p></div>')
    .setTitle(APP.TITLE)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT);
}

/** GitHub Pages 填表頁以 POST 送出（Content-Type: text/plain，避免跨網域預檢） */
function doPost(e) {
  var p;
  try { p = JSON.parse(e && e.postData ? e.postData.contents : ''); }
  catch (err) { return json_(fail_('INVALID', '資料格式錯誤，請重新整理頁面後再試。')); }
  return json_(submitResponse(p));
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/** 籌備團隊統計頁：只回傳彙整數字，不含姓名、公司、職稱、學校、科系或自由文字 */
function getPublicStats_(key) {
  var k = PropertiesService.getScriptProperties().getProperty('STATS_KEY');
  if (!k || String(key || '') !== k) return { ok: false, message: '統計頁連結無效或已重設，請向主辦人索取最新連結。' };
  try {
    var stats = computeStats_(readRows_());
    return { ok: true, stats: stats, summary: buildSummaryText_(stats), updatedAt: Utilities.formatDate(new Date(), APP.TZ, 'yyyy/MM/dd HH:mm') };
  } catch (err) {
    return { ok: false, message: '系統尚未完成設定。' };
  }
}

function getStatsKey_(rotate) {
  var props = PropertiesService.getScriptProperties();
  var k = props.getProperty('STATS_KEY');
  if (!k || rotate) {
    k = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '').slice(0, 8);
    props.setProperty('STATS_KEY', k);
  }
  return k;
}

function statsUrl_() { return APP.PAGES_URL + 'stats.html#key=' + getStatsKey_(false); }

/** 選單：顯示可分享給籌備同學的統計頁連結 */
function showStatsLink() {
  requireAdmin_();
  var url = statsUrl_();
  var html = HtmlService.createHtmlOutput(
    '<div style="font-family:sans-serif;font-size:14px;line-height:1.7">' +
    '<p>把下面的連結傳給要看統計的同學。統計頁只有彙整數字，<b>不含任何姓名、公司或職稱</b>。</p>' +
    '<textarea id="u" readonly style="width:100%;height:84px;font-size:13px">' + url + '</textarea>' +
    '<p><button onclick="var t=document.getElementById(\'u\');t.select();document.execCommand(\'copy\');this.textContent=\'已複製\'">複製連結</button>　' +
    '<a href="' + url + '" target="_blank">開啟統計頁</a></p>' +
    '<p style="color:#777">連結外流時，可按選單「重設統計頁連結」讓舊連結失效。</p></div>'
  ).setWidth(520).setHeight(300);
  SpreadsheetApp.getUi().showModalDialog(html, '籌備團隊統計頁連結');
}

function resetStatsLink() {
  requireAdmin_();
  var ui = SpreadsheetApp.getUi();
  if (ui.alert('重設統計頁連結', '舊連結將立即失效，需要把新連結重新傳給同學。確定要重設嗎？', ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  getStatsKey_(true);
  showStatsLink();
}

/** 填表頁載入時取得題目與頁面簽章憑證（不含任何填寫資料） */
function getFormConfig() {
  var count = 0;
  try { count = Math.max(0, getDataSheet_().getLastRow() - 1); } catch (e) { /* 尚未初始化 */ }
  return {
    questions: getQuestions_(),
    token: issueToken_(),
    closed: count >= APP.MAX_TOTAL
  };
}

/** 接收一份問卷 */
function submitResponse(payload) {
  try {
    return submitInner_(payload);
  } catch (err) {
    console.error(err && err.stack ? err.stack : err);
    return fail_('SERVER', '系統暫時無法儲存，請稍候再按一次「確認送出」，不會重複登記。');
  }
}

function submitInner_(p) {
  if (!p || typeof p !== 'object') return fail_('INVALID', '資料格式錯誤，請重新整理頁面後再試。');

  // 隱藏陷阱欄位：機器人才會填，假裝成功但不寫入
  if (p.hp) return { ok: true };

  var t = checkToken_(p.token);
  if (t === 'TOO_FAST') return { ok: false, code: 'TOO_FAST', wait: APP.MIN_FILL_MS, message: '安全驗證中，請稍候…' };
  if (t === 'EXPIRED') return fail_('TOKEN', '頁面已開啟太久，正在重新驗證…');
  if (t !== 'OK') return fail_('TOKEN', '頁面驗證失敗，正在重新驗證…');

  var sid = String(p.sid || '');
  if (!/^[A-Za-z0-9_-]{8,64}$/.test(sid)) return fail_('INVALID', '送出編號錯誤，請重新整理頁面後再試。');
  if (p.consent !== true) return fail_('VALIDATION', '請勾選同意資料使用說明後再送出。', 'CONSENT');

  var v = validateAnswers_(p.answers || {});
  if (!v.ok) return v;

  var lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) return fail_('BUSY', '目前填寫人數較多，請稍候幾秒再按一次「確認送出」。');
  try {
    var sh = getDataSheet_();
    var last = sh.getLastRow();

    // 同一份問卷重送（網路不穩重按）只寫入一次
    if (last > 1) {
      var found = sh.getRange(2, COL.SID + 1, last - 1, 1)
        .createTextFinder(sid).matchEntireCell(true).findNext();
      if (found) return { ok: true, duplicate: true };
    }
    if (last - 1 >= APP.MAX_TOTAL) return fail_('CLOSED', '本問卷已達填寫上限，如需協助請聯繫活動主辦人。');

    // 全站每分鐘上限
    var cache = CacheService.getScriptCache();
    var rlKey = 'rl_' + Math.floor(Date.now() / 60000);
    var n = Number(cache.get(rlKey) || 0);
    if (n >= APP.MAX_PER_MINUTE) return fail_('RATE', '目前填寫人數較多，請等候約一分鐘後再按一次「確認送出」。');
    cache.put(rlKey, String(n + 1), 120);

    var row = v.row;
    row[COL.TIME] = Utilities.formatDate(new Date(), APP.TZ, 'yyyy/MM/dd HH:mm:ss');
    row[COL.CONSENT] = APP.CONSENT_VALUE;
    row[COL.SID] = sid;

    // 同名：舊資料標示「已由較新填寫取代」，統計只計最新一筆
    var note = '';
    var key = normName_(row[COL.NAME]);
    if (last > 1 && key) {
      var names = sh.getRange(2, COL.NAME + 1, last - 1, 1).getValues();
      var notesRange = sh.getRange(2, COL.NOTE + 1, last - 1, 1);
      var notes = notesRange.getValues();
      var dup = 0, changed = false;
      for (var i = 0; i < names.length; i++) {
        if (normName_(names[i][0]) !== key) continue;
        dup++;
        var cur = String(notes[i][0] || '');
        if (cur.indexOf(APP.REPLACED_MARK) === -1) {
          notes[i][0] = cur ? cur + '；' + APP.REPLACED_MARK : APP.REPLACED_MARK;
          changed = true;
        }
      }
      if (changed) notesRange.setValues(notes);
      if (dup) note = '同名重複填寫（第 ' + (dup + 1) + ' 次，統計以此筆為準）';
    }
    row[COL.NOTE] = note;

    sh.getRange(last + 1, 1, 1, HEADERS.length).setValues([row.map(safeCell_)]);
    SpreadsheetApp.flush();
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

/** 後端驗證：必填、選項必須在題目中、長度上限 */
function validateAnswers_(a) {
  var qs = getQuestions_();
  var row = [];
  for (var c = 0; c < HEADERS.length; c++) row.push('');

  for (var i = 0; i < qs.length; i++) {
    var q = qs[i];
    var m = FIELD_MAP[q.code];
    if (!m) continue;
    var other = '';

    if (q.type === 'text') {
      var s = clean_(a[q.code]);
      if (q.required && !s) return fail_('VALIDATION', '請填寫「' + q.title + '」。', q.code);
      if (s.length > (q.max || 80)) return fail_('VALIDATION', '「' + q.title + '」字數過長（上限 ' + (q.max || 80) + ' 字）。', q.code);
      row[m.col] = s;
      continue;
    }

    if (q.otherText) {
      other = clean_(a[q.code + '_OTHER']);
      if (other.length > 100) return fail_('VALIDATION', '「其他」說明字數過長（上限 100 字）。', q.code);
    }

    if (q.type === 'single') {
      var val = clean_(a[q.code]);
      if (!val) {
        if (q.required) return fail_('VALIDATION', '請選擇「' + q.title + '」。', q.code);
        continue;
      }
      if (q.options.indexOf(val) === -1) return fail_('VALIDATION', '「' + q.title + '」的選項已更新，請重新整理頁面後再填寫。', q.code);
      if (isOther_(val) && q.otherText && other) {
        if (m.otherCol != null) row[m.otherCol] = other;
        else val = val + '（' + other + '）';
      }
      row[m.col] = val;
      continue;
    }

    if (q.type === 'multi') {
      var arr = Array.isArray(a[q.code]) ? a[q.code] : [];
      var picked = [];
      for (var j = 0; j < arr.length; j++) {
        var x = clean_(arr[j]);
        if (!x) continue;
        if (q.options.indexOf(x) === -1) return fail_('VALIDATION', '「' + q.title + '」的選項已更新，請重新整理頁面後再填寫。', q.code);
        if (picked.indexOf(x) === -1) picked.push(x);
      }
      if (q.required && !picked.length) return fail_('VALIDATION', '「' + q.title + '」請至少選擇一項。', q.code);
      // 依題目順序排列
      picked.sort(function (p1, p2) { return q.options.indexOf(p1) - q.options.indexOf(p2); });
      var hasOther = picked.some(isOther_);
      if (hasOther && q.otherText && other) {
        if (m.otherCol != null) row[m.otherCol] = other;
        else picked = picked.map(function (o) { return isOther_(o) ? o + '（' + other + '）' : o; });
      }
      row[m.col] = picked.join(APP.MULTI_SEP);
    }
  }
  return { ok: true, row: row };
}

/* =========================================================
 * 頁面簽章憑證（防機器人、防偽造）
 * ========================================================= */

function getSecret_() {
  var props = PropertiesService.getScriptProperties();
  var s = props.getProperty('TOKEN_SECRET');
  if (!s) {
    s = Utilities.getUuid() + Utilities.getUuid();
    props.setProperty('TOKEN_SECRET', s);
  }
  return s;
}

function sign_(text) {
  return Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(text, getSecret_()));
}

function issueToken_() {
  var ts = String(Date.now());
  return ts + '.' + sign_(ts);
}

function checkToken_(token) {
  var parts = String(token || '').split('.');
  if (parts.length !== 2 || !/^\d{13}$/.test(parts[0])) return 'INVALID';
  if (sign_(parts[0]) !== parts[1]) return 'INVALID';
  var age = Date.now() - Number(parts[0]);
  if (age < APP.MIN_FILL_MS) return 'TOO_FAST';
  if (age > APP.TOKEN_TTL_MS) return 'EXPIRED';
  return 'OK';
}

/* =========================================================
 * 主辦人：選單與初始化
 * ========================================================= */

function onOpen() {
  SpreadsheetApp.getUi().createMenu('⚛ 量子活動管理')
    .addItem('開啟統計儀表板', 'openDashboard')
    .addItem('更新「統計分析」與「講者摘要」', 'refreshReportSheets')
    .addItem('匯出 Excel／CSV', 'openExport')
    .addSeparator()
    .addItem('取得統計頁分享連結（給籌備同學）', 'showStatsLink')
    .addItem('重設統計頁連結', 'resetStatsLink')
    .addSeparator()
    .addItem('套用題目設定', 'applyQuestionSettings')
    .addSeparator()
    .addItem('活動後：刪除可識別個人資料', 'deleteIdentifiableData')
    .addToUi();
}

/** 第一次使用時，在 Apps Script 編輯器執行一次 */
function setupSystem() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('請從試算表的「擴充功能 › Apps Script」開啟此專案後再執行。');
  PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', ss.getId());
  requireAdmin_();
  getSecret_();
  getStatsKey_(false);
  ss.setSpreadsheetTimeZone(APP.TZ);

  // 填寫資料
  var sh = ss.getSheetByName(APP.SHEET_DATA) || ss.insertSheet(APP.SHEET_DATA, 0);
  sh.getRange(1, 1, sh.getMaxRows(), HEADERS.length).setNumberFormat('@');
  sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS])
    .setFontWeight('bold').setBackground(APP.BRAND_NAVY).setFontColor('#FFFFFF')
    .setVerticalAlignment('middle');
  sh.setRowHeight(1, 32);
  sh.setFrozenRows(1);
  var widths = [150, 110, 160, 140, 160, 170, 130, 170, 140, 140, 220, 380, 160, 110, 120, 220];
  widths.forEach(function (w, i) { sh.setColumnWidth(i + 1, w); });

  // 題目設定（已存在則保留管理者的修改）
  var cs = ss.getSheetByName(APP.SHEET_CONFIG);
  if (!cs) {
    cs = ss.insertSheet(APP.SHEET_CONFIG);
    writeConfigSheet_(cs, getQuestions_());
  }

  if (!ss.getSheetByName(APP.SHEET_STATS)) ss.insertSheet(APP.SHEET_STATS);
  if (!ss.getSheetByName(APP.SHEET_SUMMARY)) ss.insertSheet(APP.SHEET_SUMMARY);

  // 移除預設的空白工作表
  ['工作表1', 'Sheet1', '工作表 1'].forEach(function (name) {
    var d = ss.getSheetByName(name);
    if (d && ss.getSheets().length > 1 && d.getLastRow() === 0) ss.deleteSheet(d);
  });

  refreshReportSheets_();
  ss.setActiveSheet(sh);
  try { onOpen(); } catch (e) { /* 從編輯器執行時沒有 UI，重新整理試算表即會出現選單 */ }
  console.log('初始化完成。請回到試算表並重新整理頁面。');
}

function openDashboard() { showDashboard_('overview'); }
function openExport() { showDashboard_('export'); }

function showDashboard_(tab) {
  requireAdmin_();
  var t = HtmlService.createTemplateFromFile('Dashboard');
  t.initialTab = tab;
  var html = t.evaluate().setWidth(1120).setHeight(780);
  SpreadsheetApp.getUi().showModalDialog(html, '量子科技交流｜主辦人儀表板');
}

/* =========================================================
 * 主辦人：儀表板資料、報表、匯出
 * ========================================================= */

function getDashboardData() {
  requireAdmin_();
  var rows = readRows_();
  var stats = computeStats_(rows);
  return {
    stats: stats,
    summary: buildSummaryText_(stats),
    headers: HEADERS,
    records: rows,
    replacedMark: APP.REPLACED_MARK,
    updatedAt: Utilities.formatDate(new Date(), APP.TZ, 'yyyy/MM/dd HH:mm')
  };
}

function refreshReportSheets() {
  requireAdmin_();
  var stats = refreshReportSheets_();
  SpreadsheetApp.getActiveSpreadsheet().toast('已更新，有效填寫 ' + stats.total + ' 人。', '統計分析與講者摘要', 5);
}

function refreshReportSheets_() {
  var ss = getSS_();
  var stats = computeStats_(readRows_());
  writeStatsSheet_(ss.getSheetByName(APP.SHEET_STATS) || ss.insertSheet(APP.SHEET_STATS), stats);
  writeSummarySheet_(ss.getSheetByName(APP.SHEET_SUMMARY) || ss.insertSheet(APP.SHEET_SUMMARY), stats);
  return stats;
}

function exportCsv() {
  requireAdmin_();
  var values = getDataSheet_().getDataRange().getDisplayValues();
  var csv = values.map(function (r) { return r.map(csvCell_).join(','); }).join('\r\n');
  return { filename: '量子科技交流_填寫資料_' + fileStamp_() + '.csv', content: csv };
}

function exportExcel() {
  requireAdmin_();
  refreshReportSheets_();
  var ss = getSS_();
  var temp = SpreadsheetApp.create('量子科技交流_匯出暫存_' + fileStamp_());
  try {
    var blank = temp.getSheets()[0];
    [APP.SHEET_DATA, APP.SHEET_STATS, APP.SHEET_SUMMARY].forEach(function (name) {
      var src = ss.getSheetByName(name);
      if (src) src.copyTo(temp).setName(name);
    });
    temp.deleteSheet(blank);
    SpreadsheetApp.flush();
    var url = 'https://docs.google.com/spreadsheets/d/' + temp.getId() + '/export?format=xlsx';
    var resp = UrlFetchApp.fetch(url, {
      headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
      muteHttpExceptions: true
    });
    if (resp.getResponseCode() !== 200) throw new Error('匯出失敗（代碼 ' + resp.getResponseCode() + '），請稍後再試。');
    return {
      filename: '量子科技交流_參與者背景_' + fileStamp_() + '.xlsx',
      base64: Utilities.base64Encode(resp.getContent())
    };
  } finally {
    try { DriveApp.getFileById(temp.getId()).setTrashed(true); } catch (e) { /* 忽略 */ }
  }
}

/* =========================================================
 * 主辦人：題目設定
 * ========================================================= */

var CONFIG_HEADERS = ['代碼（請勿修改）', '題目文字', '說明文字', '選項（一行一個）', '選「其他…」時顯示文字框（是／否）', '題型（請勿修改）', '必填（請勿修改）'];
var TYPE_LABEL = { text: '文字', single: '單選', multi: '多選' };

function writeConfigSheet_(cs, qs) {
  cs.clear();
  cs.getRange(1, 1, 1, CONFIG_HEADERS.length).setValues([CONFIG_HEADERS])
    .setFontWeight('bold').setBackground(APP.BRAND_NAVY).setFontColor('#FFFFFF').setWrap(true);
  var rows = qs.map(function (q) {
    return [
      q.code, q.title, q.help || '',
      q.options ? q.options.join('\n') : '',
      q.options ? (q.otherText ? '是' : '否') : '',
      (q.parent ? '補充欄位・' : '') + TYPE_LABEL[q.type],
      q.required ? '必填' : '選填'
    ];
  });
  cs.getRange(2, 1, rows.length, CONFIG_HEADERS.length).setValues(rows)
    .setWrap(true).setVerticalAlignment('top');
  cs.setFrozenRows(1);
  [110, 260, 320, 340, 150, 120, 90].forEach(function (w, i) { cs.setColumnWidth(i + 1, w); });
  cs.getRange(2, 1, rows.length, 1).setBackground('#EEF2F8').setFontColor('#55607A');
  cs.getRange(2, 6, rows.length, 2).setBackground('#EEF2F8').setFontColor('#55607A');
  try {
    cs.getRange(1, 1, rows.length + 1, 1).protect().setDescription('題目代碼').setWarningOnly(true);
    cs.getRange(1, 6, rows.length + 1, 2).protect().setDescription('題型與必填').setWarningOnly(true);
  } catch (e) { /* 忽略 */ }
}

function applyQuestionSettings() {
  requireAdmin_();
  var ui = SpreadsheetApp.getUi();
  var cs = getSS_().getSheetByName(APP.SHEET_CONFIG);
  if (!cs) { ui.alert('找不到「題目設定」工作表，請先執行 setupSystem。'); return; }

  var values = cs.getDataRange().getValues();
  var overrides = {}, errors = [];
  var defaults = {};
  DEFAULT_QUESTIONS.forEach(function (q) { defaults[q.code] = q; });

  for (var r = 1; r < values.length; r++) {
    var code = String(values[r][0] || '').trim();
    if (!code) continue;
    var d = defaults[code];
    if (!d) { errors.push('第 ' + (r + 1) + ' 列：代碼「' + code + '」不存在，請勿修改代碼欄。'); continue; }
    var title = String(values[r][1] || '').trim();
    var help = String(values[r][2] || '').trim();
    if (!title) errors.push(code + '：題目文字不可空白。');
    var o = { title: title, help: help };
    if (d.options) {
      var opts = String(values[r][3] || '').split(/\r?\n/).map(function (x) { return x.trim(); }).filter(String);
      if (opts.length < 2) errors.push(code + '：至少需要兩個選項。');
      var seen = {};
      opts.forEach(function (x) {
        if (seen[x]) errors.push(code + '：選項「' + x + '」重複。');
        if (x.indexOf(APP.MULTI_SEP) !== -1) errors.push(code + '：選項不可包含全形分號「；」。');
        if (x.length > 60) errors.push(code + '：選項「' + x.slice(0, 12) + '…」超過 60 字。');
        seen[x] = true;
      });
      o.options = opts;
      o.otherText = String(values[r][4] || '').trim() === '是';
    }
    overrides[code] = o;
  }

  if (errors.length) {
    ui.alert('題目設定有誤，尚未套用', errors.join('\n'), ui.ButtonSet.OK);
    return;
  }
  PropertiesService.getScriptProperties().setProperty('QUESTION_CONFIG', JSON.stringify(overrides));
  ui.alert('已套用題目設定', '填表頁重新整理後即會顯示新內容。\n\n提醒：已有人填寫後，建議只修改文字，不要刪除或重排選項。', ui.ButtonSet.OK);
}

function getQuestions_() {
  var qs = JSON.parse(JSON.stringify(DEFAULT_QUESTIONS));
  var raw = PropertiesService.getScriptProperties().getProperty('QUESTION_CONFIG');
  if (!raw) return qs;
  var ov = {};
  try { ov = JSON.parse(raw) || {}; } catch (e) { return qs; }
  qs.forEach(function (q) {
    var o = ov[q.code];
    if (!o) return;
    if (o.title) q.title = o.title;
    if (typeof o.help === 'string') q.help = o.help;
    if (q.options && Array.isArray(o.options) && o.options.length >= 2) q.options = o.options;
    if (q.options && typeof o.otherText === 'boolean') q.otherText = o.otherText;
  });
  return qs;
}

/* =========================================================
 * 主辦人：活動後刪除可識別個人資料
 * ========================================================= */

function deleteIdentifiableData() {
  requireAdmin_();
  var ui = SpreadsheetApp.getUi();
  var ans = ui.alert('刪除可識別個人資料',
    '將清除「填寫資料」中所有的姓名、公司、職稱、學校、科系與自由填寫文字，只保留統計用的選項。\n\n' +
    '此動作無法復原。建議先執行「更新統計分析與講者摘要」保留最終統計。\n\n確定要刪除嗎？',
    ui.ButtonSet.YES_NO);
  if (ans !== ui.Button.YES) return;

  var sh = getDataSheet_();
  var last = sh.getLastRow();
  if (last < 2) { ui.alert('目前沒有填寫資料。'); return; }
  var n = last - 1;

  [COL.NAME, COL.IND_OTHER, COL.COMPANY, COL.TITLE, COL.SCHOOL, COL.DEPT, COL.TOPIC_OTHER].forEach(function (c) {
    sh.getRange(2, c + 1, n, 1).clearContent();
  });
  // 若選項後附有「其他（自由文字）」，只保留「其他」
  [COL.IND, COL.ROLE, COL.EDU, COL.FAM, COL.TOPICS].forEach(function (c) {
    var rg = sh.getRange(2, c + 1, n, 1);
    rg.setValues(rg.getValues().map(function (v) {
      return [String(v[0]).replace(/其他（[^；]*）/g, '其他')];
    }));
  });
  PropertiesService.getScriptProperties().setProperty('PII_DELETED_AT', new Date().toISOString());
  SpreadsheetApp.flush();
  refreshReportSheets_();

  ui.alert('已完成',
    '已清除 ' + n + ' 筆資料中的可識別個人資料。\n\n' +
    '還有兩件事請記得：\n' +
    '1. 刪除電腦中已下載的匯出檔（Excel／CSV）。\n' +
    '2. Google 試算表的「版本記錄」仍保有舊內容。若要徹底刪除，請點「檔案 › 建立副本」，確認副本正確後，把這份原始試算表移到垃圾桶並清空。',
    ui.ButtonSet.OK);
}

/* =========================================================
 * 統計
 * ========================================================= */

function readRows_() {
  var sh = getDataSheet_();
  var last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, HEADERS.length).getValues().map(function (r) {
    return r.map(function (v) {
      if (v instanceof Date) return Utilities.formatDate(v, APP.TZ, 'yyyy/MM/dd HH:mm:ss');
      return v == null ? '' : String(v);
    });
  });
}

function computeStats_(rows) {
  var qmap = {};
  getQuestions_().forEach(function (q) { qmap[q.code] = q; });

  var valid = rows.filter(function (r) {
    if (String(r[COL.NOTE] || '').indexOf(APP.REPLACED_MARK) !== -1) return false;
    return !!(r[COL.IND] || r[COL.ROLE] || r[COL.EDU] || r[COL.FAM] || r[COL.TOPICS]);
  });
  var total = valid.length;

  function pct(c) { return total ? Math.round(c / total * 100) : 0; }

  function countSingle(col, options) {
    var counts = {}, order = options.slice();
    options.forEach(function (o) { counts[o] = 0; });
    valid.forEach(function (r) {
      var label = matchOption_(r[col], options);
      if (!label) return;
      if (!(label in counts)) { counts[label] = 0; order.push(label); }
      counts[label]++;
    });
    return order.map(function (o) { return { label: o, count: counts[o], pct: pct(counts[o]) }; });
  }

  var edu = { stem: 0, biz: 0, other: 0 };
  valid.forEach(function (r) {
    var s = String(r[COL.EDU] || '');
    if (s.indexOf('理工') !== -1) edu.stem++;
    else if (s.indexOf('商') !== -1 || s.indexOf('管理') !== -1) edu.biz++;
    else if (s) edu.other++;
  });

  var topicOpts = qmap.Q6.options;
  var tCounts = {}, tOrder = topicOpts.slice();
  topicOpts.forEach(function (o) { tCounts[o] = 0; });
  valid.forEach(function (r) {
    String(r[COL.TOPICS] || '').split(APP.MULTI_SEP).forEach(function (x) {
      var label = matchOption_(x, topicOpts);
      if (!label) return;
      if (!(label in tCounts)) { tCounts[label] = 0; tOrder.push(label); }
      tCounts[label]++;
    });
  });
  var topics = tOrder.map(function (o, i) { return { label: o, count: tCounts[o], pct: pct(tCounts[o]), i: i }; })
    .sort(function (a, b) { return b.count - a.count || a.i - b.i; })
    .map(function (t) { return { label: t.label, count: t.count, pct: t.pct }; });

  return {
    total: total,
    rawCount: rows.length,
    replacedCount: rows.filter(function (r) { return String(r[COL.NOTE] || '').indexOf(APP.REPLACED_MARK) !== -1; }).length,
    eduGroups: [
      { label: '理工背景', count: edu.stem, pct: pct(edu.stem) },
      { label: '商管背景', count: edu.biz, pct: pct(edu.biz) },
      { label: '其他背景', count: edu.other, pct: pct(edu.other) }
    ],
    eduDetail: countSingle(COL.EDU, qmap.Q4.options),
    industry: countSingle(COL.IND, qmap.Q2.options),
    role: countSingle(COL.ROLE, qmap.Q3.options),
    familiarity: countSingle(COL.FAM, qmap.Q5.options),
    topics: topics
  };
}

function matchOption_(value, options) {
  var s = String(value || '').trim();
  if (!s) return '';
  if (options.indexOf(s) !== -1) return s;
  for (var i = 0; i < options.length; i++) {
    if (s.indexOf(options[i] + '（') === 0) return options[i];
  }
  if (isOther_(s)) {
    for (var j = 0; j < options.length; j++) if (isOther_(options[j])) return options[j];
  }
  return s; // 舊選項（題目修改前的資料）
}

function sortedByCount_(list) {
  return list.map(function (x, i) { return { x: x, i: i }; })
    .sort(function (a, b) { return b.x.count - a.x.count || a.i - b.i; })
    .map(function (o) { return o.x; });
}

/** 講者專用摘要：只有彙整數字，不含姓名、公司、職稱或自由填寫文字 */
function buildSummaryText_(s) {
  var L = [];
  L.push('【量子電腦的未來｜參與者背景摘要】');
  L.push('統計時間：' + Utilities.formatDate(new Date(), APP.TZ, 'yyyy/MM/dd HH:mm'));
  L.push('');
  if (!s.total) {
    L.push('目前尚無填寫資料。');
    return L.join('\n');
  }
  L.push('本次參與者共 ' + s.total + ' 人。');
  L.push('');
  L.push('▍教育背景');
  s.eduGroups.forEach(function (g) { L.push('・' + g.label + '：' + g.count + ' 人（' + g.pct + '%）'); });
  L.push('');
  L.push('▍主要參與者產業');
  sortedByCount_(s.industry).filter(function (x) { return x.count; })
    .forEach(function (x) { L.push('・' + x.label + '：' + x.count + ' 人（' + x.pct + '%）'); });
  L.push('');
  L.push('▍職務類型');
  sortedByCount_(s.role).filter(function (x) { return x.count; })
    .forEach(function (x) { L.push('・' + x.label + '：' + x.count + ' 人（' + x.pct + '%）'); });
  L.push('');
  L.push('▍對量子電腦的了解程度');
  s.familiarity.forEach(function (x) { L.push('・' + x.label + '：' + x.count + ' 人（' + x.pct + '%）'); });
  L.push('');
  L.push('▍最感興趣的議題（可複選）');
  s.topics.filter(function (x) { return x.count; })
    .forEach(function (x, i) { L.push((i + 1) + '. ' + x.label + '（' + x.count + ' 人，' + x.pct + '%）'); });
  L.push('');
  L.push('註：本摘要僅含彙整統計，不包含任何姓名、公司、職稱等可識別個人身分的資料。');
  return L.join('\n');
}

function writeStatsSheet_(sh, s) {
  sh.getCharts().forEach(function (c) { sh.removeChart(c); });
  sh.clear();
  sh.getRange('A1').setValue('量子科技交流｜參與者背景統計').setFontSize(15).setFontWeight('bold').setFontColor(APP.BRAND_NAVY);
  sh.getRange('A2').setValue('更新時間：' + Utilities.formatDate(new Date(), APP.TZ, 'yyyy/MM/dd HH:mm') +
    '　有效填寫：' + s.total + ' 人（同名重複填寫只計最新一筆）').setFontColor('#55607A');
  sh.setColumnWidth(1, 300); sh.setColumnWidth(2, 80); sh.setColumnWidth(3, 80); sh.setColumnWidth(4, 30);

  var blocks = [
    ['教育背景分類', s.eduGroups],
    ['大學主要領域', s.eduDetail],
    ['所屬產業', s.industry],
    ['職務類型', s.role],
    ['量子科技熟悉程度', s.familiarity],
    ['感興趣的主題（可複選，依人數排序）', s.topics]
  ];
  var row = 4;
  blocks.forEach(function (b) {
    var list = b[1];
    sh.getRange(row, 1).setValue(b[0]).setFontWeight('bold').setFontColor(APP.BRAND_NAVY);
    var data = [['項目', '人數', '占比']].concat(list.map(function (x) {
      return [x.label, x.count, s.total ? x.count / s.total : 0];
    }));
    var rg = sh.getRange(row + 1, 1, data.length, 3);
    rg.setValues(data);
    sh.getRange(row + 1, 1, 1, 3).setFontWeight('bold').setBackground('#EEF2F8');
    sh.getRange(row + 2, 3, list.length, 1).setNumberFormat('0%');
    sh.getRange(row + 2, 1, list.length, 1).setWrap(true);

    var chart = sh.newChart()
      .setChartType(Charts.ChartType.BAR)
      .addRange(sh.getRange(row + 1, 1, list.length + 1, 2))
      .setNumHeaders(1)
      .setPosition(row, 5, 0, 0)
      .setOption('title', b[0])
      .setOption('legend', { position: 'none' })
      .setOption('colors', ['#2F6BFF'])
      .setOption('width', 560)
      .setOption('height', Math.max(220, 60 + list.length * 34))
      .build();
    sh.insertChart(chart);
    var chartRows = Math.ceil(Math.max(220, 60 + list.length * 34) / 21) + 1;
    row += Math.max(data.length + 3, chartRows);
  });
  return sh;
}

function writeSummarySheet_(sh, s) {
  sh.clear();
  var lines = buildSummaryText_(s).split('\n');
  sh.getRange(1, 1, lines.length, 1).setValues(lines.map(function (l) { return [l]; }))
    .setWrap(true).setFontSize(12).setVerticalAlignment('middle');
  sh.getRange(1, 1).setFontSize(15).setFontWeight('bold').setFontColor(APP.BRAND_NAVY);
  lines.forEach(function (l, i) {
    if (l.indexOf('▍') === 0) sh.getRange(i + 1, 1).setFontWeight('bold').setFontColor(APP.BRAND_NAVY);
  });
  sh.setColumnWidth(1, 720);
  return sh;
}

/* =========================================================
 * 權限與工具
 * ========================================================= */

function isAdmin_() {
  var email = String(Session.getActiveUser().getEmail() || '').toLowerCase();
  if (!email) return false; // 匿名填表者一律拒絕
  var ss = getSS_();
  var list = ss.getEditors().map(function (u) { return String(u.getEmail() || '').toLowerCase(); });
  try { var o = ss.getOwner(); if (o) list.push(String(o.getEmail() || '').toLowerCase()); } catch (e) { /* 忽略 */ }
  return list.indexOf(email) !== -1;
}

function requireAdmin_() {
  if (!isAdmin_()) throw new Error('權限不足：只有試算表擁有者或編輯者可以使用此功能。');
}

function getSS_() {
  var ss = null;
  try { ss = SpreadsheetApp.getActiveSpreadsheet(); } catch (e) { ss = null; }
  if (ss) return ss;
  var id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (!id) throw new Error('系統尚未初始化，請先執行 setupSystem。');
  return SpreadsheetApp.openById(id);
}

function getDataSheet_() {
  var sh = getSS_().getSheetByName(APP.SHEET_DATA);
  if (!sh) throw new Error('找不到「填寫資料」工作表，請先執行 setupSystem。');
  return sh;
}

function fail_(code, message, field) {
  return { ok: false, code: code, message: message, field: field || '' };
}

function clean_(v) {
  return String(v == null ? '' : v)
    .replace(/[\u0000-\u001f\u007f\u200b-\u200f\u2028\u2029\ufeff]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function isOther_(s) { return String(s || '').indexOf('其他') === 0; }

function normName_(s) {
  return String(s || '').replace(/^'/, '').replace(/\s+/g, '').toLowerCase();
}

/** 防止以 = + - @ 開頭的文字被試算表當成公式 */
function safeCell_(v) {
  var s = String(v == null ? '' : v);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function csvCell_(v) {
  var s = String(v == null ? '' : v);
  if (/^[=+\-@]/.test(s)) s = "'" + s;
  if (/[",\r\n]/.test(s)) s = '"' + s.replace(/"/g, '""') + '"';
  return s;
}

function fileStamp_() {
  return Utilities.formatDate(new Date(), APP.TZ, 'yyyyMMdd_HHmm');
}
