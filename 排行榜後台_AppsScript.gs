/**
 * 「消失的學生：校園防疫迷蹤（內新國小版）」排行榜後台
 * 貼到 Google 試算表的「擴充功能 → Apps Script」，再部署成「網頁應用程式」。
 * 成績會存在這份試算表的「排行榜」工作表；要清空排行榜，刪掉第 2 列以下的資料即可。
 */
const SHEET_NAME = '排行榜';

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(['時間', '班級', '暱稱', '分數', '用時(秒)', '答錯次數']);
    sh.setFrozenRows(1);
  }
  return sh;
}

function clean_(v, max) {
  return String(v == null ? '' : v).replace(/[<>=+\-@\t\r\n]/g, '').trim().slice(0, max);
}

// 學生破案時送出成績
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const d = JSON.parse(e.postData.contents || '{}');
    const name = clean_(d.name, 12);
    const cls = clean_(d.cls, 10);
    const score = Math.max(0, Math.min(1000, Number(d.score) || 0));
    const sec = Math.max(0, Math.min(86400, Number(d.sec) || 0));
    const wrong = Math.max(0, Math.min(999, Number(d.wrong) || 0));
    if (!name || !cls) return out_({ ok: false });
    getSheet_().appendRow([new Date(), cls, name, score, sec, wrong]);
    return out_({ ok: true });
  } finally {
    lock.releaseLock();
  }
}

// 排行榜讀取成績
function doGet(e) {
  const values = getSheet_().getDataRange().getValues().slice(1);
  const list = values
    .filter(r => r[2] !== '')
    .map(r => ({ cls: String(r[1]), name: String(r[2]), score: Number(r[3]) || 0, sec: Number(r[4]) || 0, wrong: Number(r[5]) || 0 }))
    .sort((a, b) => (b.score - a.score) || (a.sec - b.sec))
    .slice(0, 300);
  return out_(list);
}

function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
