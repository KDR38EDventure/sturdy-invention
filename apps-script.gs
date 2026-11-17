/**
 * DR K's Ed-Venture & E-Career Lab — 구글 스프레드시트 양방향 연동 스크립트 (v10)
 *  - doPost: 학생 제출(성찰·상담·역량진단)을 시트에 쌓고, 교사 관리 데이터(상담 상태·메모, 생기부 초안,
 *            차시 메모, 진도·링크·편집 내용)를 [교사관리] 탭에 저장합니다. 교사 데이터는 SYNC_KEY가 맞아야 저장됩니다.
 *  - doGet : 교사 관제실의 [시트와 동기화]가 호출합니다. SYNC_KEY가 일치할 때만 응답합니다.
 * 사용법: 스프레드시트 [확장 프로그램] → [Apps Script]에 전체를 붙여 넣고
 *         [배포] → [새 배포] → 유형 '웹 앱', 실행 '나', 액세스 '모든 사용자'로 배포합니다.
 */
var SYNC_KEY = "여기에_관제실에서_만든_동기화_키를_넣으세요";

var SHEETS = {
  reflection: { name: "성찰기록", head: ["작성 일시", "학년", "반", "번호", "이름", "학기", "차시", "핵심 활동 주제", "배운 점", "느낀 점", "적용 계획", "SEL 마음날씨", "자기주도 체크"] },
  counseling: { name: "상담신청", head: ["접수일", "신청구분", "학번·학생명", "연락처", "상담방식", "상담분야", "희망일", "희망시간", "진로확신도", "스트레스", "시급도", "사전 질문·고민"] },
  survey: { name: "역량진단", head: ["작성 일시", "구분", "학년", "반", "번호", "이름", "문항1", "문항2", "문항3", "문항4", "문항5", "문항6", "문항7", "문항8", "문항9", "문항10", "문항11", "문항12"] },
  teacher: { name: "교사관리", head: ["종류", "키", "값(JSON)", "수정 시각(ms)", "수정 일시"] }
};

function keyOk_(k) {
  return !!SYNC_KEY && SYNC_KEY.indexOf("__") !== 0 && k === SYNC_KEY;
}

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

/* 교사 관리 데이터: (종류, 키)가 같으면 더 최신 것만 덮어씀 */
function upsertTeacher_(items) {
  var sh = getSheet_(SHEETS.teacher);
  var last = sh.getLastRow();
  var rows = last > 1 ? sh.getRange(2, 1, last - 1, 4).getValues() : [];
  var index = {};
  for (var i = 0; i < rows.length; i++) index[rows[i][0] + "\u0001" + rows[i][1]] = i;
  var saved = 0;
  for (var j = 0; j < items.length; j++) {
    var it = items[j] || {};
    if (!it.kind || it.key == null) continue;
    var t = Number(it.updatedAt) || 0;
    var row = [String(it.kind), String(it.key), JSON.stringify(it.value == null ? null : it.value), t, new Date(t || Date.now())];
    var id = row[0] + "\u0001" + row[1];
    if (id in index) {
      if (t >= (Number(rows[index[id]][3]) || 0)) { sh.getRange(index[id] + 2, 1, 1, 5).setValues([row]); saved++; }
    } else {
      sh.appendRow(row); index[id] = rows.length; rows.push(row); saved++;
    }
  }
  return saved;
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var d = JSON.parse(e.postData.contents);
    if (d.type === "teacher") {
      if (!keyOk_(d.key)) return json_({ ok: false, error: "unauthorized" });
      return json_({ ok: true, saved: upsertTeacher_(d.items || []) });
    }
    var def, row;
    if (d.type === "counseling") {
      def = SHEETS.counseling;
      row = [d.createdAt, d.applicantType, d.student, d.contact, d.method, d.category, d.date, d.time, d.clarity, d.stress, d.urgency, d.note];
    } else if (d.type === "survey") {
      def = SHEETS.survey;
      row = [d.date, d.phase, d.grade, d.cls, d.num, d.name].concat(d.answers || []);
    } else {
      def = SHEETS.reflection;
      row = [d.date, d.grade, d.cls, d.num, d.name, d.sem, d.no, d.topic, d.learned, d.felt, d.plan, d.weather, d.checks];
    }
    getSheet_(def).appendRow(row.map(safe_));
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function rows_(def, cols) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(def.name);
  if (!sh || sh.getLastRow() < 2) return [];
  var range = sh.getRange(2, 1, sh.getLastRow() - 1, cols || def.head.length);
  return cols ? range.getValues() : range.getDisplayValues();
}

function doGet(e) {
  var p = (e && e.parameter) || {};
  if (!keyOk_(p.key)) return json_({ ok: false, error: "unauthorized" });
  try {
    return json_({
      ok: true,
      reflections: rows_(SHEETS.reflection),
      counseling: rows_(SHEETS.counseling),
      surveys: rows_(SHEETS.survey),
      teacher: rows_(SHEETS.teacher, 4)
    });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}
