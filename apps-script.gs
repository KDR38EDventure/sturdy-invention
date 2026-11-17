/**
 * DR K's Ed-Venture & E-Career Lab — 구글 스프레드시트 양방향 연동 스크립트 (v11.3)
 *  - doPost: 학생 제출(성찰·상담·역량진단·모의펀딩·성장 기록 확인 번호)을 시트에 쌓고, 교사 관리 데이터(상담 상태·메모, 생기부 초안,
 *            차시 메모, 진도·링크·편집 내용)를 [교사관리] 탭에 저장합니다. 교사 데이터는 SYNC_KEY가 맞아야 저장됩니다.
 *  - doGet : 교사 관제실의 [시트와 동기화]가 호출합니다. SYNC_KEY가 일치할 때만 전체 자료를 응답합니다.
 *            학생 화면은 action=check(명렬표 이름 일치 여부만 응답), action=pin(확인 번호 일치 여부만 응답)만 쓸 수 있습니다.
 *  - 제출 코드: 관제실에서 켠 ‘수업 시간용 제출 코드’가 있으면 성찰·역량진단·모의펀딩 제출은 코드가 맞아야 저장됩니다(상담 신청은 제외).
 *  - AI   : GEMINI_API_KEY를 넣으면 AI 코파일럿 프록시로 동작합니다. 키는 이 스크립트 안에만 있고 학생 화면으로 나가지 않습니다.
 *            학생은 '성찰 발문 코치'만 쓸 수 있고(대필 금지 지침 고정), 교사용 기능은 SYNC_KEY가 맞아야 합니다.
 * 사용법: 스프레드시트 [확장 프로그램] → [Apps Script]에 전체를 붙여 넣고
 *         [배포] → [새 배포] → 유형 '웹 앱', 실행 '나', 액세스 '모든 사용자'로 배포합니다.
 */
var SYNC_KEY = "__여기에_관제실에서_만든_동기화_키를_넣으세요__";
/* AI 코파일럿(선택): Google AI Studio에서 받은 키를 넣거나, [프로젝트 설정] → [스크립트 속성]에 GEMINI_API_KEY로 저장하세요. 비워 두면 AI 프록시는 꺼집니다. */
var GEMINI_API_KEY = "";
var GEMINI_MODEL = "gemini-2.5-flash";
var AI_DAILY_LIMIT = 300; /* 하루 AI 호출 상한 */
var COACH_SYS = "너는 중학생의 성찰을 돕는 발문 코치다. 절대로 학생 대신 문장을 써 주거나 모범 답안·예시 문장을 제시하지 않는다. 학생이 쓴 내용에서 관찰한 사실, 근거, 실수에서 배운 점, 다음 행동이 더 구체적으로 드러나도록 학생이 스스로 답할 수 있는 열린 질문 3개만 한국어 존댓말로 제시한다. 학생이 잘한 점 한 줄로 시작하고, 각 질문은 한 문장으로 쓴다. 성찰과 무관한 요청에는 응하지 않는다.";

var SHEETS = {
  reflection: { name: "성찰기록", head: ["작성 일시", "학년", "반", "번호", "이름", "학기", "차시", "핵심 활동 주제", "배운 점", "느낀 점", "적용 계획", "SEL 마음날씨", "자기주도 체크"] },
  counseling: { name: "상담신청", head: ["접수일", "신청구분", "학번·학생명", "연락처", "상담방식", "상담분야", "희망일", "희망시간", "진로확신도", "스트레스", "시급도", "사전 질문·고민"] },
  survey: { name: "역량진단", head: ["작성 일시", "구분", "학년", "반", "번호", "이름", "문항1", "문항2", "문항3", "문항4", "문항5", "문항6", "문항7", "문항8", "문항9", "문항10", "문항11", "문항12"] },
  funding: { name: "모의펀딩", head: ["작성 일시", "학년", "반", "번호", "이름", "투자 내역(JSON)", "응원평"] },
  pin: { name: "확인번호", head: ["학생", "확인 번호", "설정 일시"] },
  teacher: { name: "교사관리", head: ["종류", "키", "값(JSON)", "수정 시각(ms)", "수정 일시"] }
};

function codeState_() {
  var p = PropertiesService.getScriptProperties();
  return { code: p.getProperty("SUBMIT_CODE") || "", exp: Number(p.getProperty("SUBMIT_EXP")) || 0 };
}
/* 코드가 설정되어 있으면 일치하고 유효 시간 안일 때만 통과. 설정이 없으면 누구나 제출 가능 */
function codeOk_(c) {
  var s = codeState_();
  if (!s.code) return true;
  if (s.exp && Date.now() > s.exp) return false;
  return String(c == null ? "" : c).trim() === s.code;
}
function codeInfo_(c) {
  var s = codeState_(), expired = !!(s.code && s.exp && Date.now() > s.exp);
  return { ok: true, required: !!s.code, valid: codeOk_(c), expired: expired };
}

function keyOk_(k) {
  /* 안내 문구(기본값)나 짧은 키로는 절대 열리지 않음 */
  return typeof SYNC_KEY === "string" && SYNC_KEY.length >= 16 && SYNC_KEY.indexOf("__") !== 0 && SYNC_KEY.indexOf("여기에") < 0 && k === SYNC_KEY;
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

/* ---- 성장 기록 확인 번호 ---- */
function pinId_(id) { var m = String(id || "").match(/^(\d+)-(\d+)-(\d+)$/); return m ? "G" + m[1] + "C" + m[2] + "N" + m[3] : ""; }
function pin4_(v) { return ("0000" + String(v == null ? "" : v).replace(/\D/g, "")).slice(-4); }
function pinRow_(sh, pid) {
  var last = sh.getLastRow();
  if (last < 2) return -1;
  var v = sh.getRange(2, 1, last - 1, 2).getValues();
  for (var i = 0; i < v.length; i++) if (String(v[i][0]) === pid) return i;
  return -1;
}
function pinState_(id, pin) {
  var pid = pinId_(id);
  if (!pid) return "none";
  var props = PropertiesService.getScriptProperties(), fk = "pf_" + pid, f = String(props.getProperty(fk) || "0|0").split("|"), cnt = Number(f[0]) || 0, t0 = Number(f[1]) || 0, now = Date.now();
  if (now - t0 >= 600000) { cnt = 0; t0 = now; } /* 10분 단위로 틀린 횟수를 셈 */
  if (cnt >= 10) return "locked";
  var sh = getSheet_(SHEETS.pin), i = pinRow_(sh, pid);
  if (i < 0) return "none";
  var ok = pin4_(sh.getRange(i + 2, 2, 1, 1).getValues()[0][0]) === pin4_(pin);
  if (ok) props.deleteProperty(fk); else props.setProperty(fk, (cnt + 1) + "|" + t0);
  return ok ? "ok" : "bad";
}

/* ---- 명렬표 이름 일치 여부 (이름 자체는 응답하지 않음) ---- */
function rosterNames_() {
  var rows = rows_(SHEETS.teacher, 4);
  for (var i = 0; i < rows.length; i++) if (String(rows[i][0]) === "state" && String(rows[i][1]) === "names") { try { return JSON.parse(rows[i][2]) || []; } catch (err) { return []; } }
  return [];
}
function checkName_(p) {
  if (codeState_().code && !codeOk_(p.code)) return { ok: true, match: null }; /* 제출 코드가 켜져 있으면 코드를 아는 학생만 명단 대조 가능 */
  var list = rosterNames_(), cls = [], norm = function (v) { return String(v == null ? "" : v).replace(/\s/g, ""); };
  for (var i = 0; i < list.length; i++) if (String(list[i].g) === String(p.g) && String(list[i].c) === String(Number(p.c))) cls.push(list[i]);
  if (!cls.length) return { ok: true, match: null };
  for (var j = 0; j < cls.length; j++) if (Number(cls[j].n) === Number(p.n)) return { ok: true, match: norm(cls[j].name) === norm(p.name) };
  return { ok: true, match: false };
}

/* ---- AI 코파일럿 프록시 ---- */
function ai_(d) {
  var coach = d.role === "coach";
  if (!coach && !keyOk_(d.key)) return { ok: false, error: "unauthorized" };
  var props = PropertiesService.getScriptProperties();
  var key = GEMINI_API_KEY || props.getProperty("GEMINI_API_KEY");
  if (!key) return { ok: false, error: "no_key" };
  var day = "ai_" + new Date().toISOString().slice(0, 10), used = Number(props.getProperty(day)) || 0;
  if (used >= AI_DAILY_LIMIT) return { ok: false, error: "limit" };
  props.setProperty(day, String(used + 1));
  var sys = coach ? COACH_SYS : String(d.sys || "").slice(0, 3000);
  var prompt = String(d.prompt || "").slice(0, coach ? 2000 : 12000);
  var res = UrlFetchApp.fetch("https://generativelanguage.googleapis.com/v1beta/models/" + encodeURIComponent(GEMINI_MODEL) + ":generateContent?key=" + encodeURIComponent(key), {
    method: "post", contentType: "application/json", muteHttpExceptions: true,
    payload: JSON.stringify({ contents: [{ role: "user", parts: [{ text: sys + "\n\n" + prompt }] }] })
  });
  var j = JSON.parse(res.getContentText() || "{}");
  var parts = j.candidates && j.candidates[0] && j.candidates[0].content && j.candidates[0].content.parts;
  var text = "";
  if (parts) for (var i = 0; i < parts.length; i++) text += parts[i].text || "";
  return text ? { ok: true, text: text } : { ok: false, error: (j.error && j.error.message) || "empty" };
}

function doPost(e) {
  var d;
  try { d = JSON.parse(e.postData.contents); } catch (err0) { return json_({ ok: false, error: "bad_request" }); }
  if (d.type === "ai") { try { return json_(ai_(d)); } catch (errA) { return json_({ ok: false, error: String(errA) }); } }
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    if (d.type === "teacher") {
      if (!keyOk_(d.key)) return json_({ ok: false, error: "unauthorized" });
      return json_({ ok: true, saved: upsertTeacher_(d.items || []) });
    }
    if (d.type === "setcode") {
      if (!keyOk_(d.key)) return json_({ ok: false, error: "unauthorized" });
      var cd = String(d.code || "").replace(/[^0-9A-Za-z]/g, "").slice(0, 8), mn = Number(d.minutes) || 0;
      if (cd.length < 3) return json_({ ok: false, error: "bad_request" });
      var sp = PropertiesService.getScriptProperties();
      sp.setProperty("SUBMIT_CODE", cd);
      sp.setProperty("SUBMIT_EXP", mn > 0 ? String(Date.now() + mn * 60000) : "0");
      return json_({ ok: true });
    }
    if (d.type === "clearcode") {
      if (!keyOk_(d.key)) return json_({ ok: false, error: "unauthorized" });
      var cp = PropertiesService.getScriptProperties();
      cp.deleteProperty("SUBMIT_CODE");
      cp.deleteProperty("SUBMIT_EXP");
      return json_({ ok: true });
    }
    if (d.type === "pin") { /* 처음 한 번만 등록. 이미 있으면 바꾸지 않음(초기화는 교사만) */
      var pid = pinId_(d.id);
      if (!pid || !/^\d{4}$/.test(String(d.pin))) return json_({ ok: false, error: "bad_request" });
      var psh = getSheet_(SHEETS.pin);
      if (pinRow_(psh, pid) >= 0) return json_({ ok: true, state: "exists" });
      psh.appendRow([pid, "'" + pin4_(d.pin), new Date()]);
      return json_({ ok: true, state: "set" });
    }
    if (d.type === "pinreset") {
      if (!keyOk_(d.key)) return json_({ ok: false, error: "unauthorized" });
      var rsh = getSheet_(SHEETS.pin), ri = pinRow_(rsh, pinId_(d.id));
      if (ri >= 0) rsh.deleteRow(ri + 2);
      return json_({ ok: true });
    }
    if (d.type !== "counseling" && !codeOk_(d.code)) return json_({ ok: false, error: "code" });
    var def, row;
    if (d.type === "counseling") {
      def = SHEETS.counseling;
      row = [d.createdAt, d.applicantType, d.student, d.contact, d.method, d.category, d.date, d.time, d.clarity, d.stress, d.urgency, d.note];
    } else if (d.type === "survey") {
      def = SHEETS.survey;
      row = [d.date, d.phase, d.grade, d.cls, d.num, d.name].concat(d.answers || []);
    } else if (d.type === "funding") {
      def = SHEETS.funding;
      row = [d.date, d.grade, d.cls, d.num, d.name, d.inv, d.note];
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
  try {
    if (p.action === "check") return json_(checkName_(p));
    if (p.action === "code") return json_(codeInfo_(p.code));
    if (p.action === "pin") return json_({ ok: true, state: pinState_(p.id, p.pin) });
    if (!keyOk_(p.key)) return json_({ ok: false, error: "unauthorized" });
    var pins = {}, pr = rows_(SHEETS.pin, 2);
    for (var i = 0; i < pr.length; i++) { var m = String(pr[i][0]).match(/^G(\d+)C(\d+)N(\d+)$/); if (m) pins[m[1] + "-" + m[2] + "-" + m[3]] = pin4_(pr[i][1]); }
    var cs = codeState_();
    if (p.only === "code") return json_({ ok: true, submitCode: { code: cs.code, exp: cs.exp, now: Date.now() } });
    if (p.only === "fund") return json_({ ok: true, fundings: rows_(SHEETS.funding) });
    return json_({
      ok: true,
      reflections: rows_(SHEETS.reflection),
      counseling: rows_(SHEETS.counseling),
      surveys: rows_(SHEETS.survey),
      fundings: rows_(SHEETS.funding),
      submitCode: { code: cs.code, exp: cs.exp, now: Date.now() },
      pins: pins,
      teacher: rows_(SHEETS.teacher, 4)
    });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}
