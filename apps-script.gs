/**
 * DR K's Ed-Venture & E-Career Lab — 구글 스프레드시트 양방향 연동 스크립트
 *  - doPost: 사이트에서 제출한 성찰·상담 신청을 시트에 한 줄씩 쌓습니다.
 *  - doGet : 교사 관제실의 [시트에서 불러오기]가 호출합니다. SYNC_KEY가 일치할 때만 응답합니다.
 * 사용법: 스프레드시트 [확장 프로그램] → [Apps Script]에 전체를 붙여 넣고
 *         [배포] → [새 배포] → 유형 '웹 앱', 실행 '나', 액세스 '모든 사용자'로 배포합니다.
 */
var SYNC_KEY = "여기에_관제실에서_만든_동기화_키를_넣으세요";

var SHEETS = {
  reflection: { name: "성찰기록", head: ["작성 일시", "학년", "반", "번호", "이름", "학기", "차시", "핵심 활동 주제", "배운 점", "느낀 점", "적용 계획", "SEL 마음날씨", "자기주도 체크"] },
  counseling: { name: "상담신청", head: ["접수일", "신청구분", "학번·학생명", "연락처", "상담방식", "상담분야", "희망일", "희망시간", "진로확신도", "스트레스", "시급도", "사전 질문·고민"] }
};

function getSheet_(def) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(def.name);
  if (!sh) {
    sh = ss.insertSheet(def.name);
    sh.appendRow(def.head);
    sh.getRange(1, 1, 1, def.head.length).setFontWeight("bold").setBackground("#0F172A").setFontColor("#FFFFFF");
    sh.setFrozenRows(1);
  }
  return sh;
}

/* 수식으로 해석될 수 있는 값은 문자로 저장 */
function safe_(v) {
  v = v == null ? "" : String(v);
  return /^[=+\-@]/.test(v) ? "'" + v : v;
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var d = JSON.parse(e.postData.contents);
    var isCns = d.type === "counseling";
    var row = isCns
      ? [d.createdAt, d.applicantType, d.student, d.contact, d.method, d.category, d.date, d.time, d.clarity, d.stress, d.urgency, d.note]
      : [d.date, d.grade, d.cls, d.num, d.name, d.sem, d.no, d.topic, d.learned, d.felt, d.plan, d.weather, d.checks];
    getSheet_(isCns ? SHEETS.counseling : SHEETS.reflection).appendRow(row.map(safe_));
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function rows_(def) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(def.name);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, def.head.length).getDisplayValues();
}

function doGet(e) {
  var p = (e && e.parameter) || {};
  if (!SYNC_KEY || SYNC_KEY.indexOf("__") === 0 || p.key !== SYNC_KEY) return json_({ ok: false, error: "unauthorized" });
  try {
    return json_({ ok: true, reflections: rows_(SHEETS.reflection), counseling: rows_(SHEETS.counseling) });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}
