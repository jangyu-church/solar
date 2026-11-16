/**
 * 사랑관 태양광 점검 앱 — 기록 저장 (구글 시트에 묶인 Apps Script)
 *
 * 처음 한 번: setup() 실행 → 권한 허용. 그다음 배포(웹 앱) → 나온 /exec 주소를 config.js의 GAS_URL에 붙여 넣기.
 * 자세한 순서는 GUIDE.md 참고.
 */

// ===== 여기만 고치면 됩니다 =====
var PIN = '4321';            // ← 꼭 바꾸세요. 앱에서 처음 한 번 입력하는 번호입니다. (6자리 이상 권장)
var TZ = 'Asia/Seoul';
var DAY_SHEET = '기록';
var CFG_SHEET = '설정';
// ================================

var KEYS = ['date', 'full', 'k1', 'k2', 'p1', 'p2', 't1', 't2', 'alarm', 'regH', 'bigH', 'bigCap', 'bigKwh', 'note', 'updated'];
var HEADERS = ['날짜', '하루치(Y/N)', '인버터1 발전량(kWh)', '인버터2 발전량(kWh)', '인버터1 출력(kW)', '인버터2 출력(kW)',
  '인버터1 누적(kWh)', '인버터2 누적(kWh)', '경고 문구', '지역평균 발전시간(h)', '큰 발전소 발전시간(h)', '큰 발전소 용량(kW)',
  '큰 발전소 발전량(kWh)', '메모', '수정 시각(ms)'];
var NUM_KEYS = ['k1', 'k2', 'p1', 'p2', 't1', 't2', 'regH', 'bigH', 'bigCap', 'bigKwh', 'updated'];

// 설정 기본값 (처음 setup() 때 시트에 들어갑니다. 앱의 '설정' 화면에서 고치면 시트에도 반영됩니다)
var CFG_DEFAULT = {
  cap1: 2.635, cap2: 2.635,
  janKwh: 11, novKwh: 10.7, augKwh: 21,
  ratioJan: 73, ratioNov: 77, ratioAug: 98,
  baselineFrom: ''
};
var CFG_NUM = ['cap1', 'cap2', 'janKwh', 'novKwh', 'augKwh', 'ratioJan', 'ratioNov', 'ratioAug'];

function setup() {
  ensure_();
  Logger.log('준비 완료. 이제 배포 → 새 배포 → 웹 앱 으로 주소를 만들어 주세요.');
}

function ensure_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var d = ss.getSheetByName(DAY_SHEET) || ss.insertSheet(DAY_SHEET);
  if (d.getLastRow() === 0) {
    d.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold').setBackground('#FDEBD0');
    d.setFrozenRows(1);
  }
  d.getRange('A:A').setNumberFormat('@');
  var c = ss.getSheetByName(CFG_SHEET) || ss.insertSheet(CFG_SHEET);
  if (c.getLastRow() === 0) {
    c.getRange(1, 1, 1, 2).setValues([['항목', '값']]).setFontWeight('bold').setBackground('#FDEBD0');
    var rows = Object.keys(CFG_DEFAULT).map(function (k) { return [k, CFG_DEFAULT[k]]; });
    c.getRange(2, 1, rows.length, 2).setValues(rows);
    c.getRange('B:B').setNumberFormat('@');
    c.getRange(2, 2, rows.length, 1).setValues(rows.map(function (r) { return [String(r[1])]; }));
  }
  return { day: d, cfg: c };
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
function pinOk_(p) { return String(p || '') === String(PIN); }

// 시트 수식 주입 방지 + 길이 제한
function clean_(v, max) {
  var s = String(v == null ? '' : v).replace(/[\r\n\t]+/g, ' ').trim().slice(0, max || 100);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}
function num_(v) {
  if (v === null || v === undefined || v === '') return '';
  var n = Number(v);
  return isFinite(n) ? n : '';
}
function dateStr_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, TZ, 'yyyy-MM-dd');
  return String(v == null ? '' : v).replace(/^'/, '');
}

// 확인용(주소만 열었을 때) + 전체 읽기(PIN 필요)
function doGet(e) {
  var p = (e && e.parameter) || {};
  if (!p.action) return json_({ ok: true, name: '사랑관 태양광 점검', note: 'PIN과 action이 있어야 읽을 수 있습니다.' });
  if (!pinOk_(p.pin)) return json_({ ok: false, error: 'pin' });
  if (p.action === 'all') return all_();
  return json_({ ok: false, error: '알 수 없는 요청' });
}

function all_() {
  var sh = ensure_();
  var d = sh.day, last = d.getLastRow(), days = [];
  if (last >= 2) {
    var vals = d.getRange(2, 1, last - 1, KEYS.length).getValues();
    vals.forEach(function (r) {
      var date = dateStr_(r[0]);
      if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) return;
      var o = {};
      KEYS.forEach(function (k, i) {
        var v = r[i];
        if (k === 'date') o[k] = date;
        else if (k === 'full') o[k] = String(v).toUpperCase() !== 'N';
        else if (NUM_KEYS.indexOf(k) > -1) o[k] = (v === '' || v === null) ? null : Number(v);
        else o[k] = String(v == null ? '' : v).replace(/^'/, '');
      });
      days.push(o);
    });
  }
  var cfg = {}, c = sh.cfg, cl = c.getLastRow();
  if (cl >= 2) {
    c.getRange(2, 1, cl - 1, 2).getValues().forEach(function (r) {
      var k = String(r[0]);
      if (!(k in CFG_DEFAULT)) return;
      var v = r[1];
      if (CFG_NUM.indexOf(k) > -1) cfg[k] = (v === '' || v === null) ? null : Number(v);
      else cfg[k] = dateStr_(v);
    });
  }
  return json_({ ok: true, days: days, cfg: cfg });
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    var body = JSON.parse(e.postData.contents);
    if (!pinOk_(body.pin)) return json_({ ok: false, error: 'pin' });
    var sh = ensure_();
    if (body.action === 'saveDay') return saveDay_(sh.day, body.day || {});
    if (body.action === 'deleteDay') return deleteDay_(sh.day, body.date);
    if (body.action === 'saveCfg') return saveCfg_(sh.cfg, body.cfg || {});
    return json_({ ok: false, error: '알 수 없는 요청' });
  } catch (err) {
    return json_({ ok: false, error: '저장 중 오류' });
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
}

function findRow_(d, date) {
  var last = d.getLastRow();
  if (last < 2) return 0;
  var col = d.getRange(2, 1, last - 1, 1).getValues();
  for (var i = 0; i < col.length; i++) if (dateStr_(col[i][0]) === date) return i + 2;
  return 0;
}

function saveDay_(d, day) {
  var date = String(day.date || '');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) return json_({ ok: false, error: '날짜가 올바르지 않습니다' });
  var row = KEYS.map(function (k) {
    if (k === 'date') return date;
    if (k === 'full') return day.full === false ? 'N' : 'Y';
    if (NUM_KEYS.indexOf(k) > -1) return num_(day[k]);
    return clean_(day[k], k === 'note' ? 200 : 100);
  });
  var r = findRow_(d, date);
  if (!r) r = d.getLastRow() + 1;
  d.getRange(r, 1, 1, KEYS.length).setValues([row]);
  return json_({ ok: true });
}

function deleteDay_(d, date) {
  var r = findRow_(d, String(date || ''));
  if (r) d.deleteRow(r);
  return json_({ ok: true });
}

function saveCfg_(c, cfg) {
  var keys = Object.keys(CFG_DEFAULT);
  var rows = keys.map(function (k) {
    var v = cfg[k];
    if (CFG_NUM.indexOf(k) > -1) v = num_(v);
    else v = /^\d{4}-\d{2}-\d{2}$/.test(String(v || '')) ? String(v) : '';
    return [k, String(v)];
  });
  c.getRange('B:B').setNumberFormat('@');
  c.getRange(2, 1, rows.length, 2).setValues(rows);
  return json_({ ok: true });
}
