/**
 * 전도폭발 사역 평가표 - Google Apps Script 백엔드 (폼 + 대시보드 겸용)
 * v2.5: 제출 수정 지원 (같은 이름+날짜 재제출 시 덮어쓰기, 수정용 확인번호=전화번호 뒤4자리(해시 저장),
 *       기존(확인번호 없는) 응답은 번호 없이 불러온 뒤 제출 시 번호 등록, 5회 오입력 시 10분 잠금)
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
var GROUPS_SUFFIX = '_조원리스트';                 // 기수별 조원 목록 탭 (자동 생성): 조 | 조원1 | 조원2 | ...
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
var PIN_HEADER = '수정PIN';      // 확인번호 해시 저장 열 (없으면 자동 생성, 대시보드에는 노출되지 않음)
var PIN_MAX_FAIL = 5;            // 오입력 허용 횟수
var PIN_LOCK_SEC = 600;          // 초과 시 잠금 시간(초)

/* ───────── 제출 (POST) ───────── */
function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    var data = JSON.parse((e && e.postData && e.postData.contents) || '{}');

    // 내 제출 불러오기 (PIN이 URL에 남지 않도록 POST 사용)
    if (data.action === 'mine') {
      var mc = data.cohort || DEFAULT_COHORT_POST;
      if (COHORTS.indexOf(mc) < 0) return json_({ error: '허용되지 않은 기수입니다: ' + mc });
      return json_(mineResult_(mc, String(data.name || '').trim(), String(data.pin || '').trim()));
    }

    if (!data.name || !data.stage || !data.group) {
      return json_({ success: false, error: '필수 항목(이름/단계/조)이 없습니다.' });
    }

    var cohort = data.cohort || DEFAULT_COHORT_POST;
    if (COHORTS.indexOf(cohort) < 0) {
      return json_({ success: false, error: '허용되지 않은 기수입니다: ' + cohort });
    }
    var pin = String(data.pin || '').trim();
    if (!/^\d{4}$/.test(pin)) {
      return json_({ success: false, error: '확인번호(전화번호 뒤 4자리)를 입력하세요.' });
    }
    var sheet = getSheet_(cohort, true);   // 탭이 없으면 헤더와 함께 자동 생성
    ensurePinCol_(sheet);
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

    // 같은 이름+날짜 응답이 이미 있으면 그 행을 덮어쓰고(수정), 없으면 새 행으로 추가
    var pinCol = idx[PIN_HEADER];
    var newHash = hashPin_(cohort, data.name, pin);
    row[pinCol] = newHash;
    var target = findRow_(sheet, idx, data.name, data.date);
    var updated = target > 0;
    if (updated) {
      if (isLocked_(cohort, data.name)) {
        return json_({ success: false, error: '확인번호를 여러 번 틀렸습니다. 10분 후 다시 시도하세요.' });
      }
      var stored = String(sheet.getRange(target, pinCol + 1).getValue()).trim();
      if (stored && stored !== newHash) {     // 기존 응답에 번호가 있으면 일치해야 덮어쓰기 가능
        addFail_(cohort, data.name);
        return json_({ success: false, error: '확인번호가 일치하지 않습니다.' });
      }
      clearFail_(cohort, data.name);
    } else {
      target = sheet.getLastRow() + 1;
    }
    sheet.getRange(target, 1, 1, headers.length).setNumberFormats([fmt]).setValues([row]);

    return json_({ success: true, updated: updated });
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
    // 조원 목록 (설문 폼의 '조 선택'용, 비밀번호 불필요): ?action=groups&cohort=62기
    if (e && e.parameter && e.parameter.action === 'groups') {
      var gc = e.parameter.cohort || DEFAULT_COHORT_POST;
      if (COHORTS.indexOf(gc) < 0) return out_({ error: '허용되지 않은 기수입니다: ' + gc }, cb);
      return out_({ cohort: gc, groups: getGroups_(gc) }, cb);
    }

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
function cellDate_(v, tz) {
  return (v instanceof Date) ? Utilities.formatDate(v, tz, 'yyyy-MM-dd') : String(v).trim();
}

// 이름+날짜가 같은 가장 최근 행 번호(1부터). 없으면 0
function findRow_(sheet, idx, name, date) {
  var last = sheet.getLastRow();
  var ni = idx[COLS.name], di = idx[COLS.date];
  if (last < 2 || ni === undefined || di === undefined) return 0;
  var vals = sheet.getRange(2, 1, last - 1, sheet.getLastColumn()).getValues();
  var tz = Session.getScriptTimeZone();
  var nm = String(name).trim(), dt = String(date || '').trim();
  for (var r = vals.length - 1; r >= 0; r--) {
    if (String(vals[r][ni]).trim() !== nm) continue;
    if (cellDate_(vals[r][di], tz) === dt) return r + 2;
  }
  return 0;
}

// ── 확인번호(전화번호 뒤 4자리) ──
function hashPin_(cohort, name, pin) {
  var raw = cohort + '|' + String(name).trim() + '|' + String(pin).trim() + '|EE-PIN';
  var bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, raw, Utilities.Charset.UTF_8);
  return bytes.map(function (b) { return ('0' + (b & 0xff).toString(16)).slice(-2); }).join('');
}

// '수정PIN' 열이 없으면 맨 끝에 추가, 열 번호(0부터) 반환
function ensurePinCol_(sheet) {
  var lc = Math.max(sheet.getLastColumn(), 1);
  var hs = sheet.getRange(1, 1, 1, lc).getValues()[0].map(function (h) { return String(h).trim(); });
  var i = hs.indexOf(PIN_HEADER);
  if (i >= 0) return i;
  sheet.getRange(1, lc + 1).setValue(PIN_HEADER).setFontWeight('bold').setBackground('#EEF3F9');
  return lc;
}

function failKey_(cohort, name) { return 'pinfail|' + cohort + '|' + String(name).trim(); }
function isLocked_(cohort, name) {
  return Number(CacheService.getScriptCache().get(failKey_(cohort, name)) || 0) >= PIN_MAX_FAIL;
}
function addFail_(cohort, name) {
  var c = CacheService.getScriptCache(), k = failKey_(cohort, name);
  c.put(k, String(Number(c.get(k) || 0) + 1), PIN_LOCK_SEC);
}
function clearFail_(cohort, name) { CacheService.getScriptCache().remove(failKey_(cohort, name)); }

// 이름 일치: 정확히 같거나, 입력값 뒤에 식별기호(영문 1~3자)만 더 붙은 경우 (예: '최창숙' ↔ '최창숙A')
function nameMatch_(cell, name) {
  var c = String(cell).trim();
  if (c === name) return true;
  return c.indexOf(name) === 0 && /^[A-Za-z]{1,3}$/.test(c.substring(name.length));
}

// 해당 이름의 제출 목록(최신순, 최대 10건). 확인번호가 등록된 응답은 번호가 맞아야 반환, 번호 없는 기존 응답은 그대로 반환(_legacy)
function mineResult_(cohort, name, pin) {
  if (!name) return { error: '이름을 입력하세요.' };
  if (isLocked_(cohort, name)) return { error: '확인번호를 여러 번 틀렸습니다. 10분 후 다시 시도하세요.' };
  var sheet = getSheet_(cohort, false);
  if (!sheet) return { rows: [] };
  var last = sheet.getLastRow(), lastCol = sheet.getLastColumn();
  if (last < 2) return { rows: [] };
  var values = sheet.getRange(1, 1, last, lastCol).getValues();
  var headers = values[0].map(function (h) { return String(h).trim(); });
  var rev = {};
  Object.keys(COLS).forEach(function (k) { rev[COLS[k]] = k; });
  var ni = headers.indexOf(COLS.name);
  var pi = headers.indexOf(PIN_HEADER);
  if (ni < 0) return { rows: [] };
  var pinOk = /^\d{4}$/.test(pin);
  var anyWant = false;
  var tz = Session.getScriptTimeZone();
  var out = [], denied = 0;
  for (var r = values.length - 1; r >= 1 && out.length < 10; r--) {
    if (!nameMatch_(values[r][ni], name)) continue;
    var stored = pi >= 0 ? String(values[r][pi]).trim() : '';
    var want = pinOk ? hashPin_(cohort, String(values[r][ni]).trim(), pin) : '';
    if (stored && stored !== want) { denied++; continue; }
    if (stored) anyWant = true;
    var o = {};
    for (var c = 0; c < headers.length; c++) {
      var key = rev[headers[c]];
      if (!key) continue;
      var v = values[r][c];
      if (v instanceof Date) v = Utilities.formatDate(v, tz, key === 'date' ? 'yyyy-MM-dd' : 'yyyy. M. d. a h:mm:ss');
      o[key] = v;
    }
    o._legacy = !stored;
    out.push(o);
  }
  if (!out.length && denied) {
    if (pinOk) { addFail_(cohort, name); return { error: '확인번호가 일치하지 않습니다.' }; }
    return { error: '확인번호(전화번호 뒤 4자리)를 입력한 뒤 다시 눌러주세요.' };
  }
  if (out.length && anyWant) clearFail_(cohort, name);
  return { rows: out };
}

function getSheet_(cohort, create) {
  var name = cohort + SHEET_SUFFIX;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(name);
  if (!sheet && create) {
    sheet = ss.insertSheet(name);
    var heads = Object.keys(COLS).map(function (k) { return COLS[k]; });
    heads.push(PIN_HEADER);
    sheet.getRange(1, 1, 1, heads.length).setValues([heads])
      .setFontWeight('bold').setBackground('#EEF3F9');
    sheet.setFrozenRows(1);
  }
  return sheet;   // create=false 이고 탭이 없으면 null
}

// ── 기수별 조원 목록 (탭 "{기수}_조원리스트": 조 | 조원1 | 조원2 | ...) ──
function getGroups_(cohort) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var name = cohort + GROUPS_SUFFIX;
  var sh = ss.getSheetByName(name);
  if (!sh) {                                   // 없으면 빈 양식 자동 생성 (1~11조)
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, 4).setValues([['조', '조원1', '조원2', '조원3']])
      .setFontWeight('bold').setBackground('#EEF3F9');
    var rows = [];
    for (var i = 1; i <= 11; i++) rows.push([i + '조']);
    sh.getRange(2, 1, rows.length, 1).setNumberFormat('@').setValues(rows);
    sh.setFrozenRows(1);
    return [];
  }
  var lastR = sh.getLastRow(), lastC = sh.getLastColumn();
  if (lastR < 2 || lastC < 1) return [];
  var vals = sh.getRange(2, 1, lastR - 1, lastC).getValues();
  var out = [];
  vals.forEach(function (row) {
    var g = String(row[0]).trim();
    if (!g) return;
    if (/^\d+$/.test(g)) g += '조';            // '1' → '1조'
    var members = row.slice(1).map(function (v) { return String(v).trim(); })
      .filter(function (v) { return v; });
    out.push({ group: g, members: members });
  });
  out.sort(function (a, b) { return (parseInt(a.group, 10) || 999) - (parseInt(b.group, 10) || 999); });
  return out;
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
