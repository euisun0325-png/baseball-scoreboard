/**
 * 야구 스코어보드 - 사용자 가입 · 승인 서버 (구글 시트)
 *
 * 시트 열: 가입일시 | 학교 | 이름 | 전화번호 | 상태 | 마지막 접속 | 앱 버전 | 메모 | (비번 확인값)
 * 메모 칸에 "관리자" 라고 적은 사람은 앱에서 '주소로 로고 가져오기' 를 쓸 수 있어요.
 * 명단 사진 읽기(글자 인식)를 쓰려면 편집기 왼쪽 '서비스 +' 에서 'Drive API' 를 추가해야 해요.
 * 상태 칸에 "승인" 이라고 바꾸면 그 사람이 앱을 쓸 수 있어요. "차단" 이면 막혀요.
 * 비밀번호를 잊은 사람은 맨 끝 칸(비번 확인값)을 지우면, 다음 로그인 때 새 비밀번호로 정해져요.
 *
 * 시트는 처음 가입이 들어올 때 내 구글 드라이브에 "스코어보드 사용자" 라는 이름으로 저절로 만들어져요.
 * (구글 시트의 확장 프로그램 → Apps Script 에 붙여 넣어도 그 시트를 그대로 써요)
 */

var SHEET = '사용자';
var HEAD = ['가입일시', '학교', '이름', '전화번호', '상태', '마지막 접속', '앱 버전', '메모', '비번 확인값', '지메일'];   // 지메일: 플레이스토어 테스터 등록용 (v17.9)
var AUTO_APPROVE = false;   // true 로 바꾸면 가입하자마자 바로 승인
var MAIL_ON_JOIN = true;    // 새 가입이 들어오면 메일로 알림
var LOGO_FOLDER = '스코어보드 로고';   // 이 폴더에 "학교이름.png" 로 올리면 앱의 로고 찾기에 나와요
var CFG_FOLDER = '스코어보드 설정 전달';  // 관리자가 '설정 보내기' 한 설정 파일이 여기에 (사람마다 1개)

function book_() {
  var ss = null;
  try { ss = SpreadsheetApp.getActiveSpreadsheet(); } catch (e) { }
  if (ss) return ss;
  var props = PropertiesService.getScriptProperties(), id = props.getProperty('SHEET_ID');
  if (id) { try { return SpreadsheetApp.openById(id); } catch (e) { } }
  ss = SpreadsheetApp.create('스코어보드 사용자');
  props.setProperty('SHEET_ID', ss.getId());
  return ss;
}

/** 편집기에서 이 함수를 한 번 실행하면 시트와 로고 폴더가 만들어지고 주소가 실행 로그에 나와요 */
function setup() {
  var ss = book_(); sheet_();
  Logger.log('사용자 시트: ' + ss.getUrl());
  Logger.log('로고 폴더: ' + logoFolder_().getUrl());
}

function logoFolder_() {
  var props = PropertiesService.getScriptProperties(), id = props.getProperty('LOGO_FOLDER_ID');
  if (id) { try { var f = DriveApp.getFolderById(id); if (!f.isTrashed()) return f; } catch (e) { } }
  var it = DriveApp.getFoldersByName(LOGO_FOLDER);
  var folder = it.hasNext() ? it.next() : DriveApp.createFolder(LOGO_FOLDER);
  props.setProperty('LOGO_FOLDER_ID', folder.getId());
  return folder;
}

/** 주소의 이미지를 받아 로고 폴더에 "이름.확장자" 로 저장 (같은 이름이 있으면 바꿔 씀) */
function saveLogoFromUrl_(it) {
  var name = String(it && it.name || '').replace(/[\\\/:*?"<>|]/g, '').trim().slice(0, 40);
  var url = String(it && it.url || '');
  if (!name || (!it.data && !/^https?:\/\//.test(url))) return { name: name, ok: false, error: '이름이나 그림이 없어요' };
  try {
    var bytes, mt;
    if (it.data) {
      // 앱(폰)이 받아서 보낸 그림
      bytes = Utilities.base64Decode(String(it.data)); mt = String(it.mime || '');
    } else {
      var res = UrlFetchApp.fetch(url, { muteHttpExceptions: true, followRedirects: true, headers: it.ref ? { Referer: String(it.ref) } : {} });
      if (res.getResponseCode() !== 200) return { name: name, ok: false, error: '받기 실패 (' + res.getResponseCode() + ')' };
      var blob = res.getBlob(); mt = blob.getContentType() || ''; bytes = blob.getBytes();
    }
    if (!/^image\//.test(mt)) {
      // 일부 서버는 형식을 알려주지 않음 → 파일 머리로 판단
      var b0 = bytes[0] & 255, b1 = bytes[1] & 255;
      mt = b0 === 0x89 && b1 === 0x50 ? 'image/png' : b0 === 0xff && b1 === 0xd8 ? 'image/jpeg' : b0 === 0x47 && b1 === 0x49 ? 'image/gif' : '';
      if (!mt) return { name: name, ok: false, error: '그림 파일이 아니에요' };
    }
    if (bytes.length > 1500000) return { name: name, ok: false, error: '그림이 너무 커요' };
    var ext = { 'image/png': '.png', 'image/jpeg': '.jpg', 'image/gif': '.gif', 'image/webp': '.webp', 'image/svg+xml': '.svg' }[mt] || '.png';
    var folder = logoFolder_(), old = folder.getFiles(), replaced = false;
    while (old.hasNext()) {
      var f = old.next();
      if (f.getName().replace(/\.[^.]+$/, '') === name) { f.setTrashed(true); replaced = true; }
    }
    folder.createFile(Utilities.newBlob(bytes, mt, name + ext));
    return { name: name, ok: true, replaced: replaced };
  } catch (e) {
    return { name: name, ok: false, error: String(e).slice(0, 80) };
  }
}

/**
 * 명단 사진 → 선수 목록 (구글 AI Gemini). 손글씨·표 모양이 제각각이어도 칸별로 정리해 줌.
 * 준비: 앱스 스크립트 → 프로젝트 설정 → 스크립트 속성 → 이름 GEMINI_KEY, 값 = AI 스튜디오에서 받은 키
 * 키가 없거나 실패하면 null → 아래 드라이브 글자 인식으로 넘어감
 */
var AI_ERR = '';
function aiRoster_(bytes, mime, names) {
  AI_ERR = '';
  var key = String(PropertiesService.getScriptProperties().getProperty('GEMINI_KEY') || '').trim();
  if (!key) { AI_ERR = 'GEMINI_KEY 없음'; return null; }
  var prompt = [
    '이 사진은 한국 초등학교 야구팀의 경기 기록지 또는 선발 명단(라인업)이다. 손글씨일 수 있고, 사진이 옆으로 돌아가 있을 수 있다.',
    '선수마다 한 줄씩 JSON 배열로만 답하라. 같은 선수는 한 번만 넣어라.',
    '- order: 타순 숫자(1~9). 타순 표에 없는 선수(후보·교체)는 null',
    '- pos: 수비 위치. 숫자로 적혀 있으면 야구 수비 번호다: 1=P 2=C 3=1B 4=2B 5=3B 6=SS 7=LF 8=CF 9=RF. 한글(투수·포수·1루수·2루수·3루수·유격수·좌익수·중견수·우익수·지명타자)이면 P,C,1B,2B,3B,SS,LF,CF,RF,DH 로 바꿔라. 없으면 빈 문자열.',
    '- no: 등번호(배번). 같은 종이에 "이름(배번)" 같은 선수 명단이 따로 있으면 거기서 같은 이름의 번호를 찾아 넣어라. 수비 번호를 등번호로 넣지 마라. 모르면 빈 문자열.',
    '- name: 선수 이름(한글 2~4자). 손글씨는 흔한 한국 이름으로 가장 그럴듯하게 읽어라. 인쇄된 선수 명단에 비슷한 이름이 있으면 그 이름을 써라.',
    '타순 표의 선수를 먼저 타순 순서대로, 그다음 선수 명단에만 있는 선수를 order null 로 넣어라.',
    '머리줄(타순·수비·이름·기록 등), 이닝 번호, 팀명, 대회명, 감독·코치, 빈 칸은 넣지 마라.'
  ].concat(names && names.length ? ['참고: 이 팀에 등록된 선수 이름은 ' + names.join(', ') + ' 이다. 손글씨가 이 중 하나와 비슷하면 그 이름으로 적어라. 목록에 없는 이름도 있을 수 있다.'] : []).join('\n');
  var body = {
    contents: [{ parts: [{ text: prompt }, { inline_data: { mime_type: mime || 'image/jpeg', data: Utilities.base64Encode(bytes) } }] }],
    generationConfig: { temperature: 0, responseMimeType: 'application/json',
      responseSchema: { type: 'ARRAY', items: { type: 'OBJECT', properties: {
        order: { type: 'INTEGER', nullable: true }, pos: { type: 'STRING' }, no: { type: 'STRING' }, name: { type: 'STRING' } }, required: ['name'] } } }
  };
  var models = aiModels_(key), tried = {}, t0 = Date.now(), n = 0;
  for (var i = 0; i < models.length; i++) {
    var md = models[i]; if (tried[md]) continue; tried[md] = 1;
    if (n >= 3 || Date.now() - t0 > 35000) { AI_ERR = (AI_ERR || '') + ' (시간이 걸려 예전 방식으로)'; break; }   // 앱이 오래 기다리지 않게
    n++;
    try {
      var res = aiCall_(key, md, body);
      if (res.code !== 200) {
        AI_ERR = md + ' ' + res.code + ': ' + res.text.replace(/\s+/g, ' ').slice(0, 160);
        var sug = res.text.match(/use (?:models\/)?(gemini-[\w.\-]+)/);        // "이 모델 대신 ○○를 쓰세요" 안내를 따라감
        if (sug && !tried[sug[1]]) models.splice(i + 1, 0, sug[1]);
        if (res.code === 400 && body.generationConfig.responseSchema) {       // 표 형식 지정을 못 받는 모델이면 형식 없이 다시
          var b2 = JSON.parse(JSON.stringify(body)); delete b2.generationConfig.responseSchema;
          res = aiCall_(key, md, b2); if (res.code !== 200) continue;
        } else continue;
      }
      var rows = aiRows_(res.text);
      if (!rows.length) { AI_ERR = md + ': 선수를 못 찾음'; continue; }
      PropertiesService.getScriptProperties().setProperty('GEMINI_OK_MODEL', md);   // 잘 된 모델을 다음에 먼저 씀
      return rows;
    } catch (e) { AI_ERR = md + ': ' + String(e).slice(0, 160); }
  }
  return null;
}

/** 쓸 모델 순서: 직접 정한 모델 → 지난번에 잘 된 모델 → 최신 flash (gemini-flash-latest, 3.x flash…) → 예전 이름 */
function aiModels_(key) {
  var P = PropertiesService.getScriptProperties(), out = [];
  [P.getProperty('GEMINI_MODEL'), P.getProperty('GEMINI_OK_MODEL')].forEach(function (m) { if (m) out.push(m); });
  if (out.length && !aiModels_.full) return out.concat(['gemini-2.5-flash']).filter(function (m, i, a) { return a.indexOf(m) === i; });   // 잘 된 모델이 있으면 목록 조회 생략 (빠르게)
  try {
    var lr = UrlFetchApp.fetch('https://generativelanguage.googleapis.com/v1beta/models?pageSize=200', { headers: { 'x-goog-api-key': key }, muteHttpExceptions: true });
    if (lr.getResponseCode() === 200) {
      var ms = (JSON.parse(lr.getContentText()).models || []).filter(function (m) {
        return (m.supportedGenerationMethods || []).indexOf('generateContent') >= 0 && /gemini/.test(m.name) && /flash/.test(m.name) && !/(embed|tts|audio|image|live|thinking)/.test(m.name);
      }).map(function (m) { return m.name.replace(/^models\//, ''); });
      var ver = function (n) { var v = n.match(/gemini-(\d+(?:\.\d+)?)/); return v ? parseFloat(v[1]) : 0; };
      // 번호가 붙은 최신 flash 먼저 (예: gemini-3.8-flash). '-latest' 별명은 응답이 멈추는 경우가 있어 맨 뒤로
      var score = function (n) { return (/latest/.test(n) ? -1000 : 0) + ver(n) * 10 - (/lite/.test(n) ? 5 : 0) - (/preview|exp/.test(n) ? 1 : 0); };
      out = out.concat(ms.sort(function (x, y) { return score(y) - score(x); }));
    }
  } catch (eL) { AI_ERR = '모델 목록: ' + String(eL).slice(0, 160); }
  return out.concat(['gemini-2.5-flash']).filter(function (m, i, a) { return m && a.indexOf(m) === i; });
}
function aiCall_(key, model, body) {
  var r = UrlFetchApp.fetch('https://generativelanguage.googleapis.com/v1beta/models/' + model + ':generateContent', {
    method: 'post', contentType: 'application/json', headers: { 'x-goog-api-key': key }, payload: JSON.stringify(body), muteHttpExceptions: true });
  return { code: r.getResponseCode(), text: String(r.getContentText()) };
}
/** AI 대답에서 선수 목록 꺼내기 (```json 감싸기 등도 처리) */
function aiRows_(txt) {
  var j = JSON.parse(txt), t = '';
  ((j.candidates && j.candidates[0] && j.candidates[0].content && j.candidates[0].content.parts) || []).forEach(function (p) { if (p.text) t += p.text; });
  var m = t.match(/\[[\s\S]*\]/); if (!m) return [];
  var rows = JSON.parse(m[0]);
  return (rows || []).filter(function (x) { return x && x.name && /[가-힣]/.test(x.name); }).slice(0, 30);
}

/**
 * 편집기에서 위쪽 함수 고르는 칸을 aiTest 로 바꾸고 ▷실행 → 아래 '실행 로그'에 결과가 나와요.
 * (처음 실행하면 외부 연결 권한을 허용하라고 나와요 → 허용)
 */
function aiTest() {
  var key = String(PropertiesService.getScriptProperties().getProperty('GEMINI_KEY') || '').trim();
  Logger.log('키: ' + (key ? key.slice(0, 6) + '… (' + key.length + '글자)' : '없음'));
  if (!key) return;
  aiModels_.full = true; var ms = aiModels_(key); aiModels_.full = false;
  Logger.log('시도할 모델 순서: ' + ms.slice(0, 6).join(', '));
  var body = { contents: [{ parts: [{ text: '다음을 JSON 배열로만 답하라: [{"order":1,"pos":"P","no":"","name":"김민준"}]' }] }] };
  for (var i = 0; i < ms.length && i < 6; i++) {
    var r = aiCall_(key, ms[i], body);
    Logger.log(ms[i] + ' → ' + r.code + (r.code === 200 ? ' 성공! 읽은 선수: ' + JSON.stringify(aiRows_(r.text)) : ' ' + r.text.replace(/\s+/g, ' ').slice(0, 200)));
    if (r.code === 200) { PropertiesService.getScriptProperties().setProperty('GEMINI_OK_MODEL', ms[i]); Logger.log('✅ 이 모델로 명단 사진을 읽어요: ' + ms[i]); return; }
    var sug = r.text.match(/use (?:models\/)?(gemini-[\w.\-]+)/); if (sug && ms.indexOf(sug[1]) < 0) ms.splice(i + 1, 0, sug[1]);
  }
  Logger.log('❌ 쓸 수 있는 모델을 찾지 못했어요');
}

/** 명단 사진 → 글자. 먼저 구글 AI(키가 있으면), 안 되면 구글 드라이브 글자 인식. 사진은 인식 후 바로 지움 */
function ocr_(q) {
  try {
    var b0 = Utilities.base64Decode(String(q.data || ''));
    if (b0.length && b0.length <= 4000000) { var nm = (q.names || []).filter(function (x) { return /^[가-힣]{2,4}$/.test(String(x)); }).slice(0, 60);
      var ai = aiRoster_(b0, String(q.mime || 'image/jpeg'), nm); if (ai) return { ok: true, rows: ai, via: 'ai' }; }
  } catch (e0) { AI_ERR = String(e0).slice(0, 160); }
  var aiErr = AI_ERR;
  if (typeof Drive === 'undefined') return { ok: false, error: '서버에 Drive API 서비스가 켜져 있지 않아요 (관리자에게 문의)' };
  var bytes = Utilities.base64Decode(String(q.data || ''));
  if (!bytes.length) return { ok: false, error: '사진이 없어요' };
  if (bytes.length > 4000000) return { ok: false, error: '사진이 너무 커요' };
  var blob = Utilities.newBlob(bytes, String(q.mime || 'image/jpeg'), 'roster.jpg');
  var id = '';
  try {
    if (Drive.Files.create) {      // Drive API v3
      id = Drive.Files.create({ name: '명단사진_' + Date.now(), mimeType: 'application/vnd.google-apps.document' }, blob, { ocrLanguage: 'ko' }).id;
    } else {                       // Drive API v2
      id = Drive.Files.insert({ title: '명단사진_' + Date.now(), mimeType: 'application/vnd.google-apps.document' }, blob, { ocr: true, ocrLanguage: 'ko' }).id;
    }
    var text = DocumentApp.openById(id).getBody().getText();
    return { ok: true, text: text, aiErr: aiErr };
  } catch (e) {
    return { ok: false, error: '글자 인식 실패: ' + String(e).slice(0, 80) };
  } finally {
    if (id) { try { DriveApp.getFileById(id).setTrashed(true); } catch (e2) { } }
  }
}

/** 로고 목록: [{name, id, t}]  (이름은 파일 이름에서 확장자를 뺀 것) */
function logoList_() {
  var out = [], files = logoFolder_().getFiles();
  while (files.hasNext()) {
    var f = files.next(), mt = f.getMimeType();
    if (!/^image\//.test(mt)) continue;
    out.push({ name: f.getName().replace(/\.[^.]+$/, ''), id: f.getId(), t: f.getLastUpdated().getTime() });
  }
  out.sort(function (a, b) { return a.name < b.name ? -1 : 1; });
  return out;
}

function sheet_() {
  var ss = book_();
  var sh = ss.getSheetByName(SHEET);
  if (!sh) {
    sh = ss.insertSheet(SHEET);
    sh.appendRow(HEAD);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, HEAD.length).setFontWeight('bold').setBackground('#e8f0fe');
    sh.setColumnWidth(1, 140); sh.setColumnWidth(6, 140); sh.setColumnWidth(9, 60);
    // 상태 칸은 목록에서 고르게
    var rule = SpreadsheetApp.newDataValidation().requireValueInList(['대기', '승인', '차단'], true).build();
    sh.getRange('E2:E').setDataValidation(rule);
  }
  var P = PropertiesService.getScriptProperties();
  if (!P.getProperty('GMAIL_COL')) {           // 예전에 만든 시트엔 '지메일' 칸 제목 추가 (한 번만)
    if (!sh.getRange(1, 10).getValue()) sh.getRange(1, 10).setValue('지메일').setFontWeight('bold').setBackground('#e8f0fe');
    P.setProperty('GMAIL_COL', '1');
  }
  return sh;
}

function gmail_(g) { g = String(g || '').trim().toLowerCase(); return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(g) ? g.slice(0, 80) : ''; }
function phone_(p) { return String(p || '').replace(/[^0-9]/g, ''); }
function now_() { return Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm'); }
function out_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

function find_(sh, phone) {
  var last = sh.getLastRow();
  if (last < 2) return 0;
  var vals = sh.getRange(2, 4, last - 1, 1).getValues();
  for (var i = 0; i < vals.length; i++) if (phone_(vals[i][0]) === phone) return i + 2;
  return 0;
}


// ------------------------------------------------------------------ 설정 보내기 · 원격 지원 (v17.8)
function cfgFolder_() {
  var props = PropertiesService.getScriptProperties(), id = props.getProperty('CFG_FOLDER_ID');
  if (id) { try { var f = DriveApp.getFolderById(id); if (!f.isTrashed()) return f; } catch (e) { } }
  var it = DriveApp.getFoldersByName(CFG_FOLDER);
  var folder = it.hasNext() ? it.next() : DriveApp.createFolder(CFG_FOLDER);
  props.setProperty('CFG_FOLDER_ID', folder.getId());
  return folder;
}
function pushMeta_(phone) { try { return JSON.parse(PropertiesService.getScriptProperties().getProperty('push_' + phone) || 'null'); } catch (e) { return null; } }
function pushMetaSet_(phone, m) { var P = PropertiesService.getScriptProperties(); if (m) P.setProperty('push_' + phone, JSON.stringify(m)); else P.deleteProperty('push_' + phone); }

// 캐시(빠른 임시 저장): 값 하나 100KB 한도라 잘게 나눠 저장
var CCH = 25000, RS_TTL = 21600;
function cPut_(k, s, ttl) {
  var c = CacheService.getScriptCache(), n = Math.max(1, Math.ceil(s.length / CCH)), m = {};
  m[k] = String(n);
  for (var i = 0; i < n; i++) m[k + '#' + i] = s.substr(i * CCH, CCH);
  c.putAll(m, ttl || RS_TTL);
}
function cGet_(k) {
  var c = CacheService.getScriptCache(), n = +c.get(k) || 0;
  if (!n) return null;
  var keys = []; for (var i = 0; i < n; i++) keys.push(k + '#' + i);
  var got = c.getAll(keys), s = '';
  for (var j = 0; j < n; j++) { if (got[keys[j]] == null) return null; s += got[keys[j]]; }
  return s;
}
function cJson_(k) { var s = cGet_(k); try { return s ? JSON.parse(s) : null; } catch (e) { return null; } }
function cDel_(k) { var c = CacheService.getScriptCache(), n = +c.get(k) || 0, keys = [k]; for (var i = 0; i < n; i++) keys.push(k + '#' + i); c.removeAll(keys); }

/** 관리자 · 원격 지원 · 설정 전달 요청 처리. 해당 없으면 null */
function support_(q, sh, row, r, phone, isAdmin) {
  var ok = r[4] === '승인';
  var a = q.action;
  if (a === 'users') {
    if (!isAdmin) return { ok: false, error: '관리자만 볼 수 있어요' };
    var last = sh.getLastRow(), rows = last < 2 ? [] : sh.getRange(2, 1, last - 1, 10).getValues(), list = [];
    rows.forEach(function (x) {
      var p = phone_(x[3]); if (!p || p === phone || x[4] !== '승인') return;
      var m = pushMeta_(p);
      list.push({ phone: p, school: String(x[1]), name: String(x[2]), last: String(x[5] || ''), ver: String(x[6] || ''), gmail: String(x[9] || ''), push: m ? { at: m.at, st: m.st } : null });
    });
    return { ok: true, users: list, me: String(r[9] || '') };
  }
  if (a === 'setGmail') {
    var g = gmail_(q.gmail);
    if (!g) return { ok: false, error: '이메일 주소를 정확히 넣어 주세요' };
    sh.getRange(row, 10).setValue(g);
    return { ok: true, gmail: g };
  }
  if (a === 'cfgPush') {
    if (!isAdmin) return { ok: false, error: '관리자만 보낼 수 있어요' };
    var to = phone_(q.to), data = String(q.data || '');
    if (!to || !find_(sh, to)) return { ok: false, error: '받는 사람을 찾지 못했어요' };
    if (!data || data.length > 8000000) return { ok: false, error: '보낼 설정이 비었거나 너무 커요' };
    var old = pushMeta_(to);
    if (old && old.fid) { try { DriveApp.getFileById(old.fid).setTrashed(true); } catch (e) { } }
    var f = cfgFolder_().createFile('설정_' + to + '.json', data, 'application/json');
    var meta = { fid: f.getId(), at: now_(), t: Date.now(), from: String(r[2] || '관리자'), memo: String(q.memo || '').slice(0, 80), cats: (q.cats || []).slice(0, 8), st: '대기' };
    pushMetaSet_(to, meta);
    return { ok: true, at: meta.at };
  }
  if (a === 'cfgGet') {
    var m2 = pushMeta_(phone);
    if (!m2) return { ok: false, error: '받을 설정이 없어요' };
    var txt = ''; try { txt = DriveApp.getFileById(m2.fid).getBlob().getDataAsString('UTF-8'); } catch (e) { return { ok: false, error: '설정 파일을 읽지 못했어요' }; }
    return { ok: true, meta: { at: m2.at, t: m2.t, from: m2.from, memo: m2.memo, cats: m2.cats }, data: txt };
  }
  if (a === 'cfgAck') {
    var m3 = pushMeta_(phone);
    if (m3 && (!q.t || +q.t === +m3.t)) { m3.st = q.res === 'later' ? '대기' : '적용'; m3.done = now_(); pushMetaSet_(phone, m3); }
    return { ok: true };
  }
  if (a === 'rsOpen') {
    if (!ok) return { ok: false, error: '승인된 사용자만 쓸 수 있어요' };
    var c = CacheService.getScriptCache(), code = '';
    for (var k = 0; k < 20 && !code; k++) { var t = String(100000 + Math.floor(Math.random() * 900000)); if (!c.get('rc:' + t)) code = t; }
    var oldS = cJson_('rs:' + phone); if (oldS && oldS.code) c.remove('rc:' + oldS.code);
    c.put('rc:' + code, phone, 600);
    cPut_('rs:' + phone, JSON.stringify({ code: code, st: 'wait', exp: Date.now() + 600000, who: String(r[2] || ''), school: String(r[1] || '') }));
    cPut_('rsnap:' + phone, String(q.snap || '{}'));
    cPut_('rgame:' + phone, String(q.game || '{}'));
    cDel_('rpat:' + phone);
    return { ok: true, code: code, exp: 600 };
  }
  if (a === 'rsJoin') {
    if (!isAdmin) return { ok: false, error: '관리자만 연결할 수 있어요' };
    var code2 = String(q.code || '').replace(/[^0-9]/g, '');
    var tp = CacheService.getScriptCache().get('rc:' + code2);
    if (!tp) return { ok: false, error: '코드가 맞지 않거나 시간이 지났어요' };
    var s2 = cJson_('rs:' + tp);
    if (!s2 || s2.st === 'end') return { ok: false, error: '상대방이 지원 요청을 끝냈어요' };
    s2.st = 'on'; s2.admin = phone; s2.adminName = String(r[2] || '관리자'); s2.seen = Date.now();
    cPut_('rs:' + tp, JSON.stringify(s2));
    CacheService.getScriptCache().remove('rc:' + code2);
    return { ok: true, to: tp, who: s2.who, school: s2.school, snap: cGet_('rsnap:' + tp) || '{}', game: cGet_('rgame:' + tp) || '{}' };
  }
  if (a === 'rsAdm') {      // 관리자: 바뀐 설정 보내기 + 상대 경기 상황 받기 (주기적으로)
    var to2 = phone_(q.to), s3 = cJson_('rs:' + to2);
    if (!isAdmin || !s3 || s3.admin !== phone) return { ok: false, error: '연결된 지원이 없어요', code: 'nosess' };
    if (s3.st === 'end') return { ok: true, st: 'end' };
    var pat = cJson_('rpat:' + to2) || { seq: 0, ls: {}, stf: null };
    var ls = q.ls || {}, any = false;
    Object.keys(ls).forEach(function (key) { pat.seq++; pat.ls[key] = { s: pat.seq, v: ls[key] }; any = true; });
    if (q.stf) { pat.seq++; pat.stf = { s: pat.seq, v: q.stf }; any = true; }
    if (any) cPut_('rpat:' + to2, JSON.stringify(pat));
    s3.seen = Date.now(); cPut_('rs:' + to2, JSON.stringify(s3));
    return { ok: true, st: s3.st, seq: pat.seq, got: s3.got || 0, meAgo: s3.lastMe ? Date.now() - s3.lastMe : 0, game: cGet_('rgame:' + to2) || '{}' };
  }
  if (a === 'rsMe') {       // 지원 받는 사람: 경기 상황 올리기 + 관리자가 바꾼 설정 받기 (주기적으로)
    var s4 = cJson_('rs:' + phone);
    if (!s4) return { ok: true, st: 'end' };
    if (s4.st === 'wait' && Date.now() > s4.exp) return { ok: true, st: 'expired' };
    if (q.game) cPut_('rgame:' + phone, String(q.game));
    var pat2 = cJson_('rpat:' + phone) || { seq: 0, ls: {}, stf: null }, since = +q.since || 0, outLs = {};
    Object.keys(pat2.ls).forEach(function (key) { if (pat2.ls[key].s > since) outLs[key] = pat2.ls[key].v; });
    if (since > (s4.got || 0)) s4.got = since;
    s4.lastMe = Date.now(); cPut_('rs:' + phone, JSON.stringify(s4));
    return { ok: true, st: s4.st, admin: s4.adminName || '', seq: pat2.seq, ls: outLs, stf: pat2.stf && pat2.stf.s > since ? pat2.stf.v : null };
  }
  if (a === 'rsEnd') {
    var tgt = q.to ? phone_(q.to) : phone, s5 = cJson_('rs:' + tgt);
    if (s5 && (tgt === phone || s5.admin === phone)) {
      s5.st = 'end'; cPut_('rs:' + tgt, JSON.stringify(s5), 600);
      if (s5.code) CacheService.getScriptCache().remove('rc:' + s5.code);
    }
    return { ok: true, st: 'end' };
  }
  return null;
}

function doGet() { return out_({ ok: true, app: 'baseball-scoreboard-auth' }); }

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var q = JSON.parse(e.postData.contents || '{}');
    var sh = sheet_();
    var phone = phone_(q.phone);
    if (phone.length < 9) return out_({ ok: false, error: '전화번호를 확인해 주세요' });
    if (!q.pw || String(q.pw).length < 20) return out_({ ok: false, error: '비밀번호를 확인해 주세요' });
    var row = find_(sh, phone);

    if (q.action === 'register') {
      var school = String(q.school || '').trim().slice(0, 40), name = String(q.name || '').trim().slice(0, 20);
      if (!school || !name) return out_({ ok: false, error: '학교와 이름을 넣어 주세요' });
      if (row) {
        var cur = sh.getRange(row, 1, 1, 9).getValues()[0];
        if (cur[8] && cur[8] !== q.pw) return out_({ ok: false, error: '이미 가입된 전화번호예요. 로그인해 주세요' });
        sh.getRange(row, 2, 1, 2).setValues([[school, name]]);
        sh.getRange(row, 9).setValue(q.pw);
        if (gmail_(q.gmail)) sh.getRange(row, 10).setValue(gmail_(q.gmail));
        return out_({ ok: true, status: cur[4] || '대기' });
      }
      var st = AUTO_APPROVE ? '승인' : '대기';
      sh.appendRow([now_(), school, name, "'" + phone, st, now_(), String(q.ver || ''), '', q.pw, gmail_(q.gmail)]);
      if (MAIL_ON_JOIN) {
        try {
          MailApp.sendEmail(Session.getEffectiveUser().getEmail(), '[스코어보드] 새 가입 요청: ' + school + ' ' + name,
            school + ' ' + name + ' (' + phone + (gmail_(q.gmail) ? ', ' + gmail_(q.gmail) : '') + ') 님이 가입했어요.\n\n시트에서 상태를 "승인"으로 바꾸면 앱을 쓸 수 있어요.\n' +
            book_().getUrl());
        } catch (err) { }
      }
      return out_({ ok: true, status: st });
    }

    // login / status
    if (!row) return out_({ ok: false, error: '가입되지 않은 전화번호예요', code: 'nouser' });
    var r = sh.getRange(row, 1, 1, 10).getValues()[0];
    if (!r[8]) sh.getRange(row, 9).setValue(q.pw);            // 비번 초기화된 경우 새 비밀번호로 정함
    else if (r[8] !== q.pw) return out_({ ok: false, error: '비밀번호가 맞지 않아요', code: 'badpw' });
    var isAdmin = String(r[7] || '').indexOf('관리자') >= 0 && r[4] === '승인';
    if (q.action === 'ocr') {
      if (r[4] !== '승인') return out_({ ok: false, error: '승인된 사용자만 쓸 수 있어요' });
      return out_(ocr_(q));
    }
    if (q.action === 'logoSave') {
      if (!isAdmin) return out_({ ok: false, error: '관리자만 로고를 저장할 수 있어요 (시트 메모 칸에 "관리자")' });
      return out_({ ok: true, results: (q.items || []).slice(0, 15).map(saveLogoFromUrl_) });
    }
    if (q.action === 'logos' || q.action === 'logo') {
      if (r[4] !== '승인') return out_({ ok: false, error: '승인된 사용자만 쓸 수 있어요' });
      if (q.action === 'logos') return out_({ ok: true, logos: logoList_() });
      var file = DriveApp.getFileById(String(q.id || ''));
      var parents = file.getParents(), inFolder = false, fid = logoFolder_().getId();
      while (parents.hasNext()) if (parents.next().getId() === fid) inFolder = true;
      if (!inFolder) return out_({ ok: false, error: '로고 폴더의 파일이 아니에요' });
      var blob = file.getBlob();
      if (blob.getBytes().length > 1500000) return out_({ ok: false, error: '이미지가 너무 커요 (1.5MB 이하로 올려 주세요)' });
      return out_({ ok: true, name: file.getName().replace(/\.[^.]+$/, ''), mime: blob.getContentType(), data: Utilities.base64Encode(blob.getBytes()) });
    }
    var sup = support_(q, sh, row, r, phone, isAdmin);
    if (sup) return out_(sup);
    sh.getRange(row, 6, 1, 2).setValues([[now_(), String(q.ver || r[6] || '')]]);
    var pm = r[4] === '승인' ? pushMeta_(phone) : null;
    return out_({ ok: true, status: r[4] || '대기', school: r[1], name: r[2], admin: isAdmin, gmail: String(r[9] || ''),
      push: pm && pm.st === '대기' ? { at: pm.at, t: pm.t, from: pm.from, memo: pm.memo, cats: pm.cats } : null });
  } catch (err) {
    return out_({ ok: false, error: '서버 오류: ' + err });
  } finally {
    lock.releaseLock();
  }
}
