/**
 * 야구 스코어보드 - 사용자 가입 · 승인 서버 (구글 시트)
 *
 * 시트 열: 가입일시 | 학교 | 이름 | 전화번호 | 상태 | 마지막 접속 | 앱 버전 | 메모 | (비번 확인값)
 * 메모 칸에 "관리자" 라고 적은 사람은 앱에서 '주소로 로고 가져오기' 를 쓸 수 있어요.
 * 상태 칸에 "승인" 이라고 바꾸면 그 사람이 앱을 쓸 수 있어요. "차단" 이면 막혀요.
 * 비밀번호를 잊은 사람은 맨 끝 칸(비번 확인값)을 지우면, 다음 로그인 때 새 비밀번호로 정해져요.
 *
 * 시트는 처음 가입이 들어올 때 내 구글 드라이브에 "스코어보드 사용자" 라는 이름으로 저절로 만들어져요.
 * (구글 시트의 확장 프로그램 → Apps Script 에 붙여 넣어도 그 시트를 그대로 써요)
 */

var SHEET = '사용자';
var HEAD = ['가입일시', '학교', '이름', '전화번호', '상태', '마지막 접속', '앱 버전', '메모', '비번 확인값'];
var AUTO_APPROVE = false;   // true 로 바꾸면 가입하자마자 바로 승인
var MAIL_ON_JOIN = true;    // 새 가입이 들어오면 메일로 알림
var LOGO_FOLDER = '스코어보드 로고';   // 이 폴더에 "학교이름.png" 로 올리면 앱의 로고 찾기에 나와요

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
  return sh;
}

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
        return out_({ ok: true, status: cur[4] || '대기' });
      }
      var st = AUTO_APPROVE ? '승인' : '대기';
      sh.appendRow([now_(), school, name, "'" + phone, st, now_(), String(q.ver || ''), '', q.pw]);
      if (MAIL_ON_JOIN) {
        try {
          MailApp.sendEmail(Session.getEffectiveUser().getEmail(), '[스코어보드] 새 가입 요청: ' + school + ' ' + name,
            school + ' ' + name + ' (' + phone + ') 님이 가입했어요.\n\n시트에서 상태를 "승인"으로 바꾸면 앱을 쓸 수 있어요.\n' +
            book_().getUrl());
        } catch (err) { }
      }
      return out_({ ok: true, status: st });
    }

    // login / status
    if (!row) return out_({ ok: false, error: '가입되지 않은 전화번호예요', code: 'nouser' });
    var r = sh.getRange(row, 1, 1, 9).getValues()[0];
    if (!r[8]) sh.getRange(row, 9).setValue(q.pw);            // 비번 초기화된 경우 새 비밀번호로 정함
    else if (r[8] !== q.pw) return out_({ ok: false, error: '비밀번호가 맞지 않아요', code: 'badpw' });
    var isAdmin = String(r[7] || '').indexOf('관리자') >= 0 && r[4] === '승인';
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
    sh.getRange(row, 6, 1, 2).setValues([[now_(), String(q.ver || r[6] || '')]]);
    return out_({ ok: true, status: r[4] || '대기', school: r[1], name: r[2], admin: isAdmin });
  } catch (err) {
    return out_({ ok: false, error: '서버 오류: ' + err });
  } finally {
    lock.releaseLock();
  }
}
