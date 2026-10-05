/**
 * 팀별 상호 피드백 수업 지원 웹앱 (RE:SEARCH를 일반화한 버전)
 * 이 스크립트가 연결된 구글 시트를 데이터베이스로 사용한다.
 *  - Feedback_<반> 시트: 팀 간 피드백 기록 (반마다 별도 탭)
 *  - Reflections_<반> 시트: 발표팀의 최종 성찰 기록 (반마다 별도 탭)
 *  - 스크립트 속성(PropertiesService): 반마다 별도의 수업 진행 상태(state) — 앱 이름, 피드백 선택지, 팀 목록 등도 여기 포함
 *
 * research-app과의 차이: 앱 이름/부제, 강점·개선점 선택지가 하드코딩이 아니라
 * state 안에 들어있어 교사가 화면에서 직접 수정할 수 있다. 다른 수업(다른 발표 주제)에
 * 재사용하려면 URL의 ?class= 값만 바꿔서 새 반을 만들면 된다(코드 복사 불필요).
 */

var CLASS_LIST_KEY = 'TFS_CLASSES';

var BAD_WORDS = ['씨발', '시발', '병신', '개새끼', '좆같', '지랄', '미친놈', '미친년', 'ㅅㅂ', 'ㅄ'];

var FEEDBACK_HEADERS = ['id', 'sessionId', 'fromTeamId', 'toTeamId', 'strengths', 'improvements', 'comment', 'approved', 'submittedAt', 'fromStudent'];
var REFLECTION_HEADERS = ['id', 'sessionId', 'teamId', 'selectedArea', 'reason', 'submittedAt'];

function doGet(e) {
  var role = e && e.parameter && e.parameter.role === 'teacher' ? 'teacher'
    : (e && e.parameter && e.parameter.role === 'student' ? 'student' : '');
  var view = e && e.parameter && e.parameter.view === 'stage' ? 'stage' : '';
  var classId = normalizeClassId_(e && e.parameter && e.parameter.class);
  var state = ensureState_(classId);
  var tmpl = HtmlService.createTemplateFromFile('index');
  tmpl.initialRole = role;
  tmpl.initialView = view;
  tmpl.initialClass = classId;
  tmpl.appName = state.appName;
  tmpl.appSubtitle = state.appSubtitle;
  return tmpl.evaluate()
    .setTitle((state.appName || '팀 피드백') + (classId === 'default' ? '' : ' · ' + classId))
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function normalizeClassId_(raw) {
  var v = String(raw || '').trim();
  return v ? v.slice(0, 40) : 'default';
}

/* ---------- 반(class) 목록 ---------- */

function listClasses() {
  var raw = PropertiesService.getScriptProperties().getProperty(CLASS_LIST_KEY);
  var list = [];
  if (raw) { try { list = JSON.parse(raw) || []; } catch (e) { list = []; } }
  return list;
}

function registerClass_(classId) {
  var list = listClasses();
  if (list.indexOf(classId) === -1) {
    list.push(classId);
    PropertiesService.getScriptProperties().setProperty(CLASS_LIST_KEY, JSON.stringify(list));
  }
}

function deleteClass(classId) {
  classId = normalizeClassId_(classId);
  if (classId === 'default') throw new Error('처음 만든 반은 삭제할 수 없어요.');
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    PropertiesService.getScriptProperties().deleteProperty(stateKey_(classId));
    var list = listClasses().filter(function (c) { return c !== classId; });
    PropertiesService.getScriptProperties().setProperty(CLASS_LIST_KEY, JSON.stringify(list));
    return {ok: true};
  } finally {
    lock.releaseLock();
  }
}

/* ---------- 시트 준비 (반마다 별도 탭) ---------- */

function sanitizeSheetName_(s) {
  return String(s).replace(/[\[\]\*\?\/\\:]/g, '_').slice(0, 90);
}

function feedbackSheetName_(classId) { return classId === 'default' ? 'Feedback' : 'Feedback_' + sanitizeSheetName_(classId); }
function reflectionSheetName_(classId) { return classId === 'default' ? 'Reflections' : 'Reflections_' + sanitizeSheetName_(classId); }

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

function feedbackSheet_(classId) { return getSheet_(feedbackSheetName_(classId), FEEDBACK_HEADERS); }
function reflectionSheet_(classId) { return getSheet_(reflectionSheetName_(classId), REFLECTION_HEADERS); }

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

function stateKey_(classId) { return 'TFS_STATE__' + classId; }

function getState_(classId) {
  var raw = PropertiesService.getScriptProperties().getProperty(stateKey_(classId));
  if (!raw) return null;
  try { return JSON.parse(raw); } catch (e) { return null; }
}

function saveState_(classId, state) {
  PropertiesService.getScriptProperties().setProperty(stateKey_(classId), JSON.stringify(state));
  return state;
}

function ensureState_(classId) {
  var state = getState_(classId);
  if (!state) {
    // classId='default'는 반 구분을 도입하기 전 쓰던 옛 단일 상태를 이어받는다.
    state = classId === 'default' ? getLegacyState_() : null;
    if (!state) state = defaultState_();
    saveState_(classId, state);
  }
  registerClass_(classId);
  return state;
}

function getLegacyState_() {
  var raw = PropertiesService.getScriptProperties().getProperty('TEAM_FEEDBACK_STATE');
  if (!raw) return null;
  try { return JSON.parse(raw); } catch (e) { return null; }
}

function containsBadWord_(text) {
  return BAD_WORDS.some(function (w) { return text.indexOf(w) > -1; });
}

/* ---------- 잠금·캐시 ---------- */

// 학생 제출(시트 쓰기)은 문서 잠금, 교사 단계 변경(상태 쓰기)은 스크립트 잠금으로 분리한다.
// 예전엔 전부 스크립트 잠금 하나라서 학생 30명이 동시에 제출하면 교사의 단계 변경이 줄 뒤에서 시간 초과로 실패했다.
// ponytail: 문서 잠금은 반 구분 없이 하나. 여러 반 동시 진행이 잦으면 반별 시트 잠금을 검토.
function sheetLock_() { return LockService.getDocumentLock() || LockService.getScriptLock(); }

function snapCacheKey_(classId, sessionId) { return 'snap__' + classId + '__' + sessionId; }
function clearSnapCache_(classId) {
  try { CacheService.getScriptCache().remove(snapCacheKey_(classId, ensureState_(classId).sessionId)); } catch (e) {}
}

function pinOk_(classId, pin) {
  var st = getState_(classId) || (classId === 'default' ? getLegacyState_() : null);
  return !!st && String(pin) === String(st.pin);
}

/* ---------- 학생용 피드백 조회·받은 피드백 공개 ---------- */

// 학생·프로젝터 화면용 피드백: 쓴 모둠·학생 이름은 본인 것만 남기고, 한마디(의견)는 교사가 공개한 것과 본인 것만 남긴다.
function studentFeedbackView_(docs, student) {
  student = String(student || '');
  return docs.map(function (d) {
    var own = !!student && String(d.fromStudent) === student;
    return {
      fromTeamId: own ? d.fromTeamId : '',
      fromStudent: own ? d.fromStudent : '',
      toTeamId: d.toTeamId,
      strengths: d.strengths,
      improvements: d.improvements,
      comment: (own || d.approved === true) ? d.comment : '',
      approved: d.approved,
      submittedAt: d.submittedAt
    };
  });
}

// 교사가 이번 수업의 '받은 피드백'을 학생에게 공개/비공개한다. 공개한 수업 id는 새 수업을 시작해도 남아, 학생이 지난 수업 기록을 다시 볼 수 있다.
function setReceivedRelease(classId, pin, on) {
  classId = normalizeClassId_(classId);
  if (!pinOk_(classId, pin)) return {ok: false, message: 'PIN이 올바르지 않아요.'};
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var state = ensureState_(classId);
    var list = (state.releasedSessions || []).filter(function (id) { return id !== state.sessionId; });
    if (on) list.push(state.sessionId);
    state.releasedSessions = list.slice(-50);
    saveState_(classId, state);
    var safe = JSON.parse(JSON.stringify(state));
    delete safe.pin;
    return {ok: true, state: safe};
  } finally {
    lock.releaseLock();
  }
}

// 학생이 '우리 모둠이 받은 피드백'을 본다. 교사가 공개한 수업의 것만, 쓴 사람 정보 없이, 교사가 공개한 한마디만 내려간다(프로젝터에 공개되는 범위와 같다).
function getMyFeedback(classId, teamId) {
  classId = normalizeClassId_(classId);
  var state = ensureState_(classId);
  var released = state.releasedSessions || [];
  teamId = String(teamId || '');
  var values = feedbackSheet_(classId).getDataRange().getValues();
  var by = {};
  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    var sid = String(row[1]);
    if (String(row[3]) !== teamId || released.indexOf(sid) < 0) continue;
    var s = by[sid] || (by[sid] = {sessionId: sid, startedAt: Number(sid.slice(1)) || 0, count: 0, strengths: {}, improvements: {}, comments: []});
    s.count++;
    String(row[4] || '').split('|').filter(Boolean).forEach(function (id) { s.strengths[id] = (s.strengths[id] || 0) + 1; });
    String(row[5] || '').split('|').filter(Boolean).forEach(function (id) { s.improvements[id] = (s.improvements[id] || 0) + 1; });
    if (row[7] === true && String(row[6] || '').trim()) s.comments.push(String(row[6]).trim());
  }
  var sessions = Object.keys(by).map(function (k) { return by[k]; }).sort(function (a, b) { return b.startedAt - a.startedAt; });
  return {ok: true, currentSessionId: state.sessionId, currentReleased: released.indexOf(state.sessionId) > -1, sessions: sessions};
}

/* ---------- 클라이언트 호출용 함수 ---------- */

function getSnapshot(classId, auth) {
  classId = normalizeClassId_(classId);
  var state = ensureState_(classId);
  // 모두가 3초마다 부르는 함수라 시트 읽기 결과를 4초 캐시한다(제출·승인 때는 캐시를 비운다).
  var cache = CacheService.getScriptCache(), key = snapCacheKey_(classId, state.sessionId), rows = null;
  try { var hit = cache.get(key); if (hit) rows = JSON.parse(hit); } catch (e) {}
  if (!rows) {
    rows = {
      feedback: readSheetForSession_(feedbackSheet_(classId), FEEDBACK_HEADERS, state.sessionId),
      reflections: readSheetForSession_(reflectionSheet_(classId), REFLECTION_HEADERS, state.sessionId)
    };
    try { cache.put(key, JSON.stringify(rows), 4); } catch (e) {} // 100KB 넘으면 캐시 없이 그대로 동작
  }
  var safeState = JSON.parse(JSON.stringify(state));
  delete safeState.pin; // 학생에게도 가는 조회라 교사 PIN은 내려보내지 않는다
  // 교사(PIN 확인)에게만 피드백 원본을 주고, 학생·프로젝터 화면에는 공개된 것만 내려보낸다.
  var teacher = !!(auth && auth.pin !== undefined && pinOk_(classId, auth.pin));
  var feedback = teacher ? rows.feedback : studentFeedbackView_(rows.feedback, auth && auth.student);
  return {
    state: safeState,
    teacher: teacher, // 화면이 '교사로 인정됐는지' 알 수 있게(저장된 PIN이 낡았으면 PIN 입력으로 돌아간다)
    feedback: feedback,
    reflections: rows.reflections,
    spreadsheetUrl: SpreadsheetApp.getActiveSpreadsheet().getUrl(),
    webAppUrl: ScriptApp.getService().getUrl(),
    classes: listClasses()
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

function teacherLogin(classId, pin) {
  classId = normalizeClassId_(classId);
  var state = getState_(classId) || (classId === 'default' ? getLegacyState_() : null);
  if (!state) {
    if (String(pin) === '0000') { ensureState_(classId); return {ok: true}; }
    return {ok: false, message: '최초 PIN은 0000이에요.'};
  }
  if (String(pin) === String(state.pin)) return {ok: true};
  return {ok: false, message: 'PIN이 올바르지 않아요.'};
}

function updateState(classId, partial) {
  classId = normalizeClassId_(classId);
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var state = ensureState_(classId);
    delete partial.releasedSessions; // 공개 목록은 setReceivedRelease(PIN 확인)로만 바꾼다
    Object.keys(partial).forEach(function (k) { state[k] = partial[k]; });
    return saveState_(classId, state);
  } finally {
    lock.releaseLock();
  }
}

function newSession(classId) {
  classId = normalizeClassId_(classId);
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var state = ensureState_(classId);
    state.sessionId = 's' + Date.now();
    state.phase = 'PRESENTATION';
    state.currentTeamId = null;
    state.presentedTeamIds = [];
    return saveState_(classId, state);
  } finally {
    lock.releaseLock();
  }
}

function saveTeams(classId, teams) {
  var cleaned = (teams || [])
    .map(function (t) { return {id: String(t.id), name: String(t.name || '').trim(), code: String(t.code || '').trim(), slideUrl: String(t.slideUrl || '').trim()}; })
    .filter(function (t) { return t.name && t.code; });
  if (!cleaned.length) throw new Error('팀은 최소 1개 이상 필요해요.');
  return updateState(classId, {teams: cleaned});
}

function saveOptionCatalog(classId, payload) {
  function clean(list, prefix) {
    return (list || [])
      .map(function (o, i) { return {id: o.id || (prefix + (i + 1)), label: String(o.label || '').trim()}; })
      .filter(function (o) { return o.label; });
  }
  var strengths = clean(payload.strengths, 's');
  var improvements = clean(payload.improvements, 'i');
  if (!strengths.length) throw new Error('강점 선택지가 1개 이상 필요해요.');
  if (!improvements.length) throw new Error('개선점 선택지가 1개 이상 필요해요.');
  return updateState(classId, {strengths: strengths, improvements: improvements});
}

function saveAppSettings(classId, payload) {
  var appName = String(payload.appName || '').trim();
  var appSubtitle = String(payload.appSubtitle || '').trim();
  if (!appName) throw new Error('앱 이름을 입력해주세요.');
  return updateState(classId, {appName: appName, appSubtitle: appSubtitle});
}

function changePin(classId, newPin) {
  var v = String(newPin || '').trim();
  if (!v) throw new Error('새 PIN을 입력해주세요.');
  return updateState(classId, {pin: v});
}

function setTeamSlideUrl(classId, payload) {
  classId = normalizeClassId_(classId);
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var state = ensureState_(classId);
    var teamId = String(payload.teamId || '');
    var url = String(payload.slideUrl || '').trim();
    var team = null;
    for (var i = 0; i < state.teams.length; i++) { if (state.teams[i].id === teamId) { team = state.teams[i]; break; } }
    if (!team) throw new Error('팀을 찾을 수 없어요.');
    team.slideUrl = url;
    return saveState_(classId, state);
  } finally {
    lock.releaseLock();
  }
}

/* ---------- 발표자료 파일 업로드 (교사 드라이브에 저장) ---------- */

var UPLOAD_ROOT_FOLDER_NAME = '팀 피드백 스튜디오 발표자료';
var MAX_UPLOAD_BASE64_LEN = 27 * 1024 * 1024; // base64는 원본보다 약 4/3배 커짐 → 대략 20MB 제한

/** 편집기에서 "실행"으로 눌러서 드라이브 권한 승인 창이 뜨는지 확인하는 용도. */
function testDriveAccess() {
  var folder = uploadRootFolder_();
  return folder.getUrl();
}

function uploadRootFolder_() {
  var it = DriveApp.getFoldersByName(UPLOAD_ROOT_FOLDER_NAME);
  if (it.hasNext()) return it.next();
  return DriveApp.createFolder(UPLOAD_ROOT_FOLDER_NAME);
}

function classUploadFolder_(classId) {
  var root = uploadRootFolder_();
  var name = sanitizeSheetName_(classId);
  var it = root.getFoldersByName(name);
  if (it.hasNext()) return it.next();
  return root.createFolder(name);
}

function uploadTeamSlide(classId, payload) {
  classId = normalizeClassId_(classId);
  var teamId = String(payload.teamId || '');
  var team = null, st = ensureState_(classId);
  for (var i = 0; i < st.teams.length; i++) { if (st.teams[i].id === teamId) { team = st.teams[i]; break; } }
  if (!team) throw new Error('팀을 찾을 수 없어요.');

  var base64 = String(payload.base64Data || '');
  if (!base64) throw new Error('파일 내용을 읽지 못했어요.');
  if (base64.length > MAX_UPLOAD_BASE64_LEN) throw new Error('파일이 너무 커요(20MB 이하로 올려주세요). 큰 파일은 링크 입력을 이용해주세요.');

  var mimeType = String(payload.mimeType || 'application/octet-stream');
  var filename = String(payload.filename || 'presentation').slice(0, 120);
  var blob = Utilities.newBlob(Utilities.base64Decode(base64), mimeType, filename);

  // 오래 걸리는 Drive 저장은 잠금 밖에서 하고, 상태에 주소를 적는 순간만 잠근다.
  var file = classUploadFolder_(classId).createFile(blob);
  file.setName(team.name + '_' + filename);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var state = ensureState_(classId);
    for (var k = 0; k < state.teams.length; k++) {
      if (state.teams[k].id === teamId) { state.teams[k].slideUrl = 'https://drive.google.com/file/d/' + file.getId() + '/preview'; break; }
    }
    saveState_(classId, state);
  } finally {
    lock.releaseLock();
  }
  return {ok: true, fileUrl: file.getUrl()};
}

function setApproval(classId, id, approved) {
  classId = normalizeClassId_(classId);
  var lock = sheetLock_();
  lock.waitLock(30000);
  try {
    var sheet = feedbackSheet_(classId);
    var values = sheet.getDataRange().getValues();
    var col = FEEDBACK_HEADERS.indexOf('approved') + 1;
    for (var i = 1; i < values.length; i++) {
      if (values[i][0] === id) { sheet.getRange(i + 1, col).setValue(!!approved); break; }
    }
    clearSnapCache_(classId);
    return {ok: true};
  } finally {
    lock.releaseLock();
  }
}

function submitFeedback(classId, payload) {
  classId = normalizeClassId_(classId);
  var lock = sheetLock_();
  lock.waitLock(30000);
  try {
    var state = ensureState_(classId);
    var strengths = (payload.strengths || []).slice(0, 2);
    var improvements = (payload.improvements || []).slice(0, 2);
    var comment = String(payload.comment || '').trim();
    var fromTeamId = String(payload.fromTeamId || '');
    var fromStudent = String(payload.fromStudent || '').trim();
    var toTeamId = String(payload.toTeamId || '');
    if (!fromTeamId || !toTeamId) throw new Error('팀 정보가 올바르지 않아요.');
    if (!fromStudent) throw new Error('학번과 이름을 먼저 입력해주세요.');
    if (fromTeamId === toTeamId) throw new Error('자기 팀에는 피드백을 보낼 수 없어요.');
    if (strengths.length < 1) throw new Error('강점을 1개 이상 선택해주세요.');
    if (improvements.length < 1) throw new Error('개선점을 1개 이상 선택해주세요.');
    if (comment.length < 5 || comment.length > 150) throw new Error('의견은 5자 이상 150자 이내로 적어주세요.');
    if (containsBadWord_(comment)) throw new Error('바르고 고운 말로 다시 적어주세요.');

    var sheet = feedbackSheet_(classId);
    var values = sheet.getDataRange().getValues();
    // 학생 개인 단위 식별: 같은 학생이 같은 팀에 다시 보내면 이전 제출을 덮어쓰고,
    // 같은 팀의 다른 학생은 서로 다른 id를 가져 각자 따로 제출할 수 있다.
    var id = state.sessionId + '__' + fromTeamId + '__' + fromStudent + '__' + toTeamId;
    var rowIndex = -1;
    for (var i = 1; i < values.length; i++) {
      if (values[i][0] === id) { rowIndex = i + 1; break; }
    }
    var rowData = [id, state.sessionId, fromTeamId, toTeamId, strengths.join('|'), improvements.join('|'), comment, '', new Date().toISOString(), fromStudent];
    if (rowIndex > 0) sheet.getRange(rowIndex, 1, 1, rowData.length).setValues([rowData]);
    else sheet.appendRow(rowData);
    clearSnapCache_(classId);
    return {ok: true};
  } finally {
    lock.releaseLock();
  }
}

function submitReflection(classId, payload) {
  classId = normalizeClassId_(classId);
  var lock = sheetLock_();
  lock.waitLock(30000);
  try {
    var state = ensureState_(classId);
    var teamId = String(payload.teamId || '');
    var area = String(payload.selectedArea || '').trim();
    var reason = String(payload.reason || '').trim();
    if (!teamId) throw new Error('팀 정보가 올바르지 않아요.');
    if (!area) throw new Error('되돌아볼 영역을 선택해주세요.');
    if (reason.length < 5) throw new Error('이유를 조금 더 자세히 적어주세요.');
    if (containsBadWord_(reason)) throw new Error('바르고 고운 말로 다시 적어주세요.');

    var sheet = reflectionSheet_(classId);
    var values = sheet.getDataRange().getValues();
    var id = state.sessionId + '__' + teamId;
    var rowIndex = -1;
    for (var i = 1; i < values.length; i++) {
      if (values[i][0] === id) { rowIndex = i + 1; break; }
    }
    var rowData = [id, state.sessionId, teamId, area, reason, new Date().toISOString()];
    if (rowIndex > 0) sheet.getRange(rowIndex, 1, 1, rowData.length).setValues([rowData]);
    else sheet.appendRow(rowData);
    clearSnapCache_(classId);
    return {ok: true};
  } finally {
    lock.releaseLock();
  }
}
