/**
 * 전도폭발 사역 평가표 - Google Apps Script 백엔드 (폼 + 대시보드 겸용)
 *
 * 시트: 수지주일반_전도폭발사역평가표 / 탭: 61기_평가표응답
 *  - 시트 1행의 "한글 헤더"를 기준으로 열을 찾아 저장하므로 열 순서가 바뀌어도 안전합니다.
 *  - 조회(GET)는 JSONP(callback 파라미터)와 일반 JSON을 모두 지원합니다.
 *
 * [적용]
 * 1) 이 스프레드시트의 확장 프로그램 → Apps Script 에 붙여 넣기
 * 2) 프로젝트 설정 → 스크립트 속성 → ADMIN_PASSWORD 추가 (관리자/대시보드 비밀번호)
 * 3) 배포 → 배포 관리 → (기존 배포) 편집 → 새 버전 → 배포
 *    실행 계정: 나 / 액세스 권한: 모든 사용자
 */

// 기수별 시트 탭 이름 = "{기수}_평가표응답"  (예: 61기_평가표응답, 62기_평가표응답)
var SHEET_SUFFIX = '_평가표응답';
var COHORTS = ['61기', '62기', '63기', '64기'];   // 허용 기수
var DEFAULT_COHORT_GET = '61기';                  // 조회 시 cohort 파라미터가 없을 때(기존 대시보드 호환)
var DEFAULT_COHORT_POST = '62기';                 // 제출 시 cohort 가 없을 때
var TOTALS_SHEET = '총원설정';                      // 기수별 단계별 총원을 저장하는 탭 (자동 생성)

// payload 키 → 시트 1행 한글 헤더
var COLS = {
  timestamp: '제출시각', class_group: '반별', group: '조', stage: '단계',
  name: '이름', date: '날짜', role: '역할',
  class_1: '강의내용', class_2: '교사전달', class_3: '시범자전달',
  class_4: '연습시간', class_5: '공개보고',
  train_1: '전도자무장', train_2: '팀원관계', train_3: '현장실습평가',
  lifestyle_1: '생활화무장', lifestyle_2: '비신자제시횟수',
  field_1: '실습참석횟수', field_2: '복음제시횟수', field_3: '재방문횟수',
  field_4: '결신자보고횟수', field_5a: '유익경험', field_5b: '난처경험및지혜',
  my_trainer: '훈련자1', trainer_name: '훈련자2',
  trainer_1: '훈련자복음지원', trainer_2: '훈련자무장지원', trainer_3: '훈련자현장참여',
  trainer_reason1: '못한이유1', trainer_reason2: '못한이유2',
  prayer_name1: '기도후원자1', prayer_name2: '기도후원자2',
  prayer_1: '기도접촉횟수', prayer_2: '기도관계점수',
  absent: '결석횟수', absent_reason: '결석이유',
  thanks_1: '감사1', thanks_2: '감사2',
  suggest_1: '제안1', suggest_2: '제안2', suggest_3: '제안3'
};

// 숫자로 저장할 항목 (나머지는 텍스트)
var NUMERIC = {
  class_1:1, class_2:1, class_3:1, class_4:1, class_5:1,
  train_1:1, train_2:1, train_3:1, lifestyle_1:1, lifestyle_2:1,
  field_1:1, field_2:1, field_3:1, field_4:1,
  trainer_1:1, trainer_2:1, prayer_1:1, prayer_2:1, absent:1
};

var MAX_LEN = 2000;

/* ───────── 제출 (POST) ───────── */
function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    var data = JSON.parse((e && e.postData && e.postData.contents) || '{}');

    if (!data.name || !data.stage || !data.group) {
      return json_({ success: false, error: '필수 항목(이름/단계/조)이 없습니다.' });
    }

    var cohort = data.cohort || DEFAULT_COHORT_POST;
    if (COHORTS.indexOf(cohort) < 0) {
      return json_({ success: false, error: '허용되지 않은 기수입니다: ' + cohort });
    }
    var sheet = getSheet_(cohort, true);   // 탭이 없으면 헤더와 함께 자동 생성
    var headers = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 1)).getValues()[0];
    var idx = {};                       // 헤더명 → 열 번호(0부터)
    headers.forEach(function (h, i) { idx[String(h).trim()] = i; });

    var row = headers.map(function () { return ''; });
    var fmt = headers.map(function () { return '@'; });

    Object.keys(COLS).forEach(function (key) {
      var col = idx[COLS[key]];
      if (col === undefined) return;    // 시트에 없는 헤더는 건너뜀
      var v = data[key];
      v = (v === undefined || v === null) ? '' : String(v);
      if (v.length > MAX_LEN) v = v.substring(0, MAX_LEN);
      if (NUMERIC[key] && v !== '' && !isNaN(Number(v))) {
        row[col] = Number(v);
        fmt[col] = '0';
      } else {
        row[col] = v;
      }
    });

    var next = sheet.getLastRow() + 1;
    var range = sheet.getRange(next, 1, 1, headers.length);
    range.setNumberFormats([fmt]).setValues([row]);

    return json_({ success: true });
  } catch (err) {
    return json_({ success: false, error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
}

/* ───────── 조회 (GET ?pw=...&callback=...) ───────── */
function doGet(e) {
  var cb = (e && e.parameter && e.parameter.callback) || '';
  try {
    var pw = (e && e.parameter && e.parameter.pw) || '';
    var saved = PropertiesService.getScriptProperties().getProperty('ADMIN_PASSWORD');

    if (!saved) return out_({ error: '서버에 ADMIN_PASSWORD가 설정되지 않았습니다.' }, cb);
    if (pw !== saved) return out_({ error: '비밀번호가 올바르지 않습니다.' }, cb);

    var cohort = (e && e.parameter && e.parameter.cohort) || DEFAULT_COHORT_GET;
    if (COHORTS.indexOf(cohort) < 0) return out_({ error: '허용되지 않은 기수입니다: ' + cohort }, cb);

    // 총원 저장 (대시보드 ⚙️ 설정): ?action=setTotals&cohort=62기&t1=11&t2=9&t3=6&t4=3&t5=3
    if (e.parameter.action === 'setTotals') {
      var tot = {};
      for (var i = 1; i <= 5; i++) {
        var n = parseInt(e.parameter['t' + i], 10);
        tot[i + '단계'] = (isNaN(n) || n < 0) ? 0 : Math.min(n, 99);
      }
      setTotals_(cohort, tot);
      return out_({ success: true, totals: tot }, cb);
    }

    var totals = getTotals_(cohort);
    var sheet = getSheet_(cohort, false);
    if (!sheet) return out_({ rows: [], totals: totals }, cb);   // 아직 탭이 없으면 빈 목록
    var last = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();
    if (last < 2) return out_({ rows: [], totals: totals }, cb);

    var values = sheet.getRange(1, 1, last, lastCol).getValues();
    var headers = values[0].map(function (h) { return String(h).trim(); });

    // 한글 헤더 → payload 키
    var rev = {};
    Object.keys(COLS).forEach(function (k) { rev[COLS[k]] = k; });

    var tz = Session.getScriptTimeZone();
    var rows = [];
    for (var r = 1; r < values.length; r++) {
      var o = {}, empty = true;
      for (var c = 0; c < headers.length; c++) {
        var key = rev[headers[c]];
        if (!key) continue;
        var v = values[r][c];
        if (v instanceof Date) v = Utilities.formatDate(v, tz, key === 'date' ? 'yyyy-MM-dd' : 'yyyy. M. d. a h:mm:ss');
        if (v !== '' && v !== null) empty = false;
        o[key] = v;
      }
      if (!empty) rows.push(o);
    }
    return out_({ rows: rows, totals: totals }, cb);
  } catch (err) {
    return out_({ error: String(err) }, cb);
  }
}

/* ───────── 도우미 ───────── */
function getSheet_(cohort, create) {
  var name = cohort + SHEET_SUFFIX;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(name);
  if (!sheet && create) {
    sheet = ss.insertSheet(name);
    var heads = Object.keys(COLS).map(function (k) { return COLS[k]; });
    sheet.getRange(1, 1, 1, heads.length).setValues([heads])
      .setFontWeight('bold').setBackground('#EEF3F9');
    sheet.setFrozenRows(1);
  }
  return sheet;   // create=false 이고 탭이 없으면 null
}

// ── 기수별 단계별 총원 (탭 "총원설정": 기수 | 1단계 | 2단계 | 3단계 | 4단계 | 5단계) ──
function getTotals_(cohort) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TOTALS_SHEET);
  if (!sh || sh.getLastRow() < 2) return null;
  var vals = sh.getRange(2, 1, sh.getLastRow() - 1, 6).getValues();
  for (var r = 0; r < vals.length; r++) {
    if (String(vals[r][0]).trim() === cohort) {
      var t = {};
      for (var i = 1; i <= 5; i++) t[i + '단계'] = Number(vals[r][i]) || 0;
      return t;
    }
  }
  return null;
}

function setTotals_(cohort, t) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName(TOTALS_SHEET);
    if (!sh) {
      sh = ss.insertSheet(TOTALS_SHEET);
      sh.getRange(1, 1, 1, 6).setValues([['기수', '1단계', '2단계', '3단계', '4단계', '5단계']])
        .setFontWeight('bold').setBackground('#EEF3F9');
      sh.setFrozenRows(1);
    }
    var row = [cohort, t['1단계'], t['2단계'], t['3단계'], t['4단계'], t['5단계']];
    var last = sh.getLastRow();
    var target = last + 1;
    if (last >= 2) {
      var names = sh.getRange(2, 1, last - 1, 1).getValues();
      for (var r = 0; r < names.length; r++) {
        if (String(names[r][0]).trim() === cohort) { target = r + 2; break; }
      }
    }
    sh.getRange(target, 1, 1, 6).setValues([row]);
  } finally {
    lock.releaseLock();
  }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// callback 이 있으면 JSONP, 없으면 JSON
function out_(obj, cb) {
  if (cb && /^[A-Za-z_$][\w$.]*$/.test(cb)) {
    return ContentService.createTextOutput(cb + '(' + JSON.stringify(obj) + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return json_(obj);
}

/** (선택) 에디터에서 실행: 권한 승인 + 헤더 매핑 점검 */
function checkHeaders() {
  var sheet = getSheet_(DEFAULT_COHORT_GET, false);
  if (!sheet) { Logger.log('탭이 없습니다: ' + DEFAULT_COHORT_GET + SHEET_SUFFIX); return; }
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
    .map(function (h) { return String(h).trim(); });
  var missing = Object.keys(COLS).filter(function (k) { return headers.indexOf(COLS[k]) < 0; });
  Logger.log(missing.length ? '시트에 없는 헤더: ' + missing.map(function (k) { return COLS[k]; }).join(', ')
                            : '모든 헤더 매핑 정상 (' + headers.length + '열)');
}
