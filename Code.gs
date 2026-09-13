/**
 * 팀별 상호 피드백 수업 지원 웹앱 (RE:SEARCH를 일반화한 버전)
 * 이 스크립트가 연결된 구글 시트를 데이터베이스로 사용한다.
 *  - Feedback 시트: 팀 간 피드백 기록
 *  - Reflections 시트: 발표팀의 최종 성찰 기록
 *  - 스크립트 속성(PropertiesService): 현재 수업 진행 상태(state) — 앱 이름, 피드백 선택지, 팀 목록 등도 여기 포함
 *
 * research-app과의 차이: 앱 이름/부제, 강점·개선점 선택지가 하드코딩이 아니라
 * state 안에 들어있어 교사가 화면에서 직접 수정할 수 있다. 다른 수업(다른 발표 주제)에
 * 재사용하려면 이 프로젝트를 통째로 복사해 새 배포를 만들면 된다.
 */

var STATE_KEY = 'TEAM_FEEDBACK_STATE';

var BAD_WORDS = ['씨발', '시발', '병신', '개새끼', '좆같', '지랄', '미친놈', '미친년', 'ㅅㅂ', 'ㅄ'];

var FEEDBACK_SHEET = 'Feedback';
var REFLECTION_SHEET = 'Reflections';
var FEEDBACK_HEADERS = ['id', 'sessionId', 'fromTeamId', 'toTeamId', 'strengths', 'improvements', 'comment', 'approved', 'submittedAt'];
var REFLECTION_HEADERS = ['id', 'sessionId', 'teamId', 'selectedArea', 'reason', 'submittedAt'];

function doGet(e) {
  var role = e && e.parameter && e.parameter.role === 'teacher' ? 'teacher'
    : (e && e.parameter && e.parameter.role === 'student' ? 'student' : '');
  var view = e && e.parameter && e.parameter.view === 'stage' ? 'stage' : '';
  var state = ensureState_();
  var tmpl = HtmlService.createTemplateFromFile('index');
  tmpl.initialRole = role;
  tmpl.initialView = view;
  tmpl.appName = state.appName;
  tmpl.appSubtitle = state.appSubtitle;
  return tmpl.evaluate()
    .setTitle(state.appName || '팀 피드백')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/* ---------- 시트 준비 ---------- */

function getSheet_(name, headers) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function feedbackSheet_() { return getSheet_(FEEDBACK_SHEET, FEEDBACK_HEADERS); }
function reflectionSheet_() { return getSheet_(REFLECTION_SHEET, REFLECTION_HEADERS); }

/* ---------- 수업 상태(State) ---------- */

function defaultState_() {
  return {
    sessionId: 's' + Date.now(),
    phase: 'PRESENTATION',
    currentTeamId: null,
    presentedTeamIds: [],
    pin: '0000',
    appName: 'RE:SEARCH',
    appSubtitle: '주제탐구독서 팀 프로젝트',
    strengths: [
      {id: 's1', label: '문제 분석이 명확했다.'},
      {id: 's2', label: '근거/자료 활용이 좋았다.'},
      {id: 's3', label: '대안이 구체적이었다.'},
      {id: 's4', label: '대안의 실현 가능성을 잘 고려했다.'},
      {id: 's5', label: '매체와 주제가 잘 연결되었다.'},
      {id: 's6', label: '발표 전달 방식이 효과적이었다.'},
      {id: 's7', label: '새로운 관점이 돋보였다.'}
    ],
    improvements: [
      {id: 'i1', label: '문제의 범위를 더 구체화하면 좋겠다.'},
      {id: 'i2', label: '문제의 원인을 더 조사하면 좋겠다.'},
      {id: 'i3', label: '근거 자료를 더 보강하면 좋겠다.'},
      {id: 'i4', label: '다른 관점도 살펴보면 좋겠다.'},
      {id: 'i5', label: '대안의 실현 가능성을 더 검토하면 좋겠다.'},
      {id: 'i6', label: '대안의 부작용/한계도 살펴보면 좋겠다.'},
      {id: 'i7', label: '매체와 내용의 연결을 강화하면 좋겠다.'},
      {id: 'i8', label: '발표 내용을 더 명확하게 구조화하면 좋겠다.'}
    ],
    teams: [
      {id: 't1', name: '1팀', code: '1A7', slideUrl: ''},
      {id: 't2', name: '2팀', code: '2B4', slideUrl: ''},
      {id: 't3', name: '3팀', code: '3C9', slideUrl: ''},
      {id: 't4', name: '4팀', code: '4D2', slideUrl: ''},
      {id: 't5', name: '5팀', code: '5E6', slideUrl: ''},
      {id: 't6', name: '6팀', code: '6F3', slideUrl: ''}
    ]
  };
}

function getState_() {
  var raw = PropertiesService.getScriptProperties().getProperty(STATE_KEY);
  if (!raw) return null;
  try { return JSON.parse(raw); } catch (e) { return null; }
}

function saveState_(state) {
  PropertiesService.getScriptProperties().setProperty(STATE_KEY, JSON.stringify(state));
  return state;
}

function ensureState_() {
  var state = getState_();
  if (!state) state = saveState_(defaultState_());
  return state;
}

function containsBadWord_(text) {
  return BAD_WORDS.some(function (w) { return text.indexOf(w) > -1; });
}

/* ---------- 클라이언트 호출용 함수 ---------- */

function getSnapshot() {
  var state = ensureState_();
  return {
    state: state,
    feedback: readSheetForSession_(feedbackSheet_(), FEEDBACK_HEADERS, state.sessionId),
    reflections: readSheetForSession_(reflectionSheet_(), REFLECTION_HEADERS, state.sessionId),
    spreadsheetUrl: SpreadsheetApp.getActiveSpreadsheet().getUrl(),
    webAppUrl: ScriptApp.getService().getUrl()
  };
}

function readSheetForSession_(sheet, headers, sessionId) {
  var values = sheet.getDataRange().getValues();
  var isFeedback = headers === FEEDBACK_HEADERS;
  var rows = [];
  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    if (String(row[1]) !== String(sessionId)) continue;
    var obj = {};
    headers.forEach(function (h, idx) { obj[h] = row[idx]; });
    if (isFeedback) {
      obj.strengths = obj.strengths ? String(obj.strengths).split('|').filter(Boolean) : [];
      obj.improvements = obj.improvements ? String(obj.improvements).split('|').filter(Boolean) : [];
      obj.approved = (obj.approved === true || obj.approved === false) ? obj.approved : null;
    }
    rows.push(obj);
  }
  return rows;
}

function teacherLogin(pin) {
  var state = getState_();
  if (!state) {
    if (String(pin) === '0000') { saveState_(defaultState_()); return {ok: true}; }
    return {ok: false, message: '최초 PIN은 0000이에요.'};
  }
  if (String(pin) === String(state.pin)) return {ok: true};
  return {ok: false, message: 'PIN이 올바르지 않아요.'};
}

function updateState(partial) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var state = ensureState_();
    Object.keys(partial).forEach(function (k) { state[k] = partial[k]; });
    return saveState_(state);
  } finally {
    lock.releaseLock();
  }
}

function newSession() {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var state = ensureState_();
    state.sessionId = 's' + Date.now();
    state.phase = 'PRESENTATION';
    state.currentTeamId = null;
    state.presentedTeamIds = [];
    return saveState_(state);
  } finally {
    lock.releaseLock();
  }
}

function saveTeams(teams) {
  var cleaned = (teams || [])
    .map(function (t) { return {id: String(t.id), name: String(t.name || '').trim(), code: String(t.code || '').trim(), slideUrl: String(t.slideUrl || '').trim()}; })
    .filter(function (t) { return t.name && t.code; });
  if (!cleaned.length) throw new Error('팀은 최소 1개 이상 필요해요.');
  return updateState({teams: cleaned});
}

function saveOptionCatalog(payload) {
  function clean(list, prefix) {
    return (list || [])
      .map(function (o, i) { return {id: o.id || (prefix + (i + 1)), label: String(o.label || '').trim()}; })
      .filter(function (o) { return o.label; });
  }
  var strengths = clean(payload.strengths, 's');
  var improvements = clean(payload.improvements, 'i');
  if (!strengths.length) throw new Error('강점 선택지가 1개 이상 필요해요.');
  if (!improvements.length) throw new Error('개선점 선택지가 1개 이상 필요해요.');
  return updateState({strengths: strengths, improvements: improvements});
}

function saveAppSettings(payload) {
  var appName = String(payload.appName || '').trim();
  var appSubtitle = String(payload.appSubtitle || '').trim();
  if (!appName) throw new Error('앱 이름을 입력해주세요.');
  return updateState({appName: appName, appSubtitle: appSubtitle});
}

function setTeamSlideUrl(payload) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var state = ensureState_();
    var teamId = String(payload.teamId || '');
    var url = String(payload.slideUrl || '').trim();
    var team = null;
    for (var i = 0; i < state.teams.length; i++) { if (state.teams[i].id === teamId) { team = state.teams[i]; break; } }
    if (!team) throw new Error('팀을 찾을 수 없어요.');
    team.slideUrl = url;
    return saveState_(state);
  } finally {
    lock.releaseLock();
  }
}

function changePin(newPin) {
  var v = String(newPin || '').trim();
  if (!v) throw new Error('새 PIN을 입력해주세요.');
  return updateState({pin: v});
}

function setApproval(id, approved) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var sheet = feedbackSheet_();
    var values = sheet.getDataRange().getValues();
    for (var i = 1; i < values.length; i++) {
      if (values[i][0] === id) { sheet.getRange(i + 1, 8).setValue(!!approved); break; }
    }
    return {ok: true};
  } finally {
    lock.releaseLock();
  }
}

function submitFeedback(payload) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var state = ensureState_();
    var strengths = (payload.strengths || []).slice(0, 2);
    var improvements = (payload.improvements || []).slice(0, 2);
    var comment = String(payload.comment || '').trim();
    var fromTeamId = String(payload.fromTeamId || '');
    var toTeamId = String(payload.toTeamId || '');
    if (!fromTeamId || !toTeamId) throw new Error('팀 정보가 올바르지 않아요.');
    if (fromTeamId === toTeamId) throw new Error('자기 팀에는 피드백을 보낼 수 없어요.');
    if (strengths.length < 1) throw new Error('강점을 1개 이상 선택해주세요.');
    if (improvements.length < 1) throw new Error('개선점을 1개 이상 선택해주세요.');
    if (comment.length < 5 || comment.length > 150) throw new Error('의견은 5자 이상 150자 이내로 적어주세요.');
    if (containsBadWord_(comment)) throw new Error('바르고 고운 말로 다시 적어주세요.');

    var sheet = feedbackSheet_();
    var values = sheet.getDataRange().getValues();
    var id = state.sessionId + '__' + fromTeamId + '__' + toTeamId;
    var rowIndex = -1;
    for (var i = 1; i < values.length; i++) {
      if (values[i][0] === id) { rowIndex = i + 1; break; }
    }
    var rowData = [id, state.sessionId, fromTeamId, toTeamId, strengths.join('|'), improvements.join('|'), comment, '', new Date().toISOString()];
    if (rowIndex > 0) sheet.getRange(rowIndex, 1, 1, rowData.length).setValues([rowData]);
    else sheet.appendRow(rowData);
    return {ok: true};
  } finally {
    lock.releaseLock();
  }
}

function submitReflection(payload) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var state = ensureState_();
    var teamId = String(payload.teamId || '');
    var area = String(payload.selectedArea || '').trim();
    var reason = String(payload.reason || '').trim();
    if (!teamId) throw new Error('팀 정보가 올바르지 않아요.');
    if (!area) throw new Error('되돌아볼 영역을 선택해주세요.');
    if (reason.length < 5) throw new Error('이유를 조금 더 자세히 적어주세요.');
    if (containsBadWord_(reason)) throw new Error('바르고 고운 말로 다시 적어주세요.');

    var sheet = reflectionSheet_();
    var values = sheet.getDataRange().getValues();
    var id = state.sessionId + '__' + teamId;
    var rowIndex = -1;
    for (var i = 1; i < values.length; i++) {
      if (values[i][0] === id) { rowIndex = i + 1; break; }
    }
    var rowData = [id, state.sessionId, teamId, area, reason, new Date().toISOString()];
    if (rowIndex > 0) sheet.getRange(rowIndex, 1, 1, rowData.length).setValues([rowData]);
    else sheet.appendRow(rowData);
    return {ok: true};
  } finally {
    lock.releaseLock();
  }
}
