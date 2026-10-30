/**
 * 姫の地図 — 기록은 구글 시트에, 사진은 구글 드라이브에 저장하는 Apps Script
 *
 * 1) 구글 시트를 새로 만들고 → 확장 프로그램 → Apps Script → 이 코드를 통째로 붙여넣기
 * 2) 위쪽 함수 선택에서 setup 을 고르고 ▶ 실행 → 권한 허용 (처음 한 번)
 * 3) 배포 → 새 배포 → 유형: 웹 앱 / 실행 사용자: 나 / 액세스 권한: 모든 사용자 → 배포
 * 4) 나온 웹 앱 주소(…/exec)를 사이트 ⚙︎ 설정에 붙여넣기
 *
 * 코드를 고친 뒤에는 배포 → 배포 관리 → ✏️ → 버전: 새 버전 으로 다시 배포해야 반영돼요.
 */

const SHEET_NAME = '기록';
const FOLDER_NAME = '姫の地図 사진';
const WRITE_KEY = ''; // 비워 두면 주소를 아는 누구나 저장 가능. 정하면 사이트 설정에 같은 비밀번호 입력.
const HEADERS = ['id', 'type', 'name', 'lat', 'lng', 'address', 'country', 'date', 'rating', 'review', 'photos', 'tags', 'createdAt', 'updatedAt'];

function setup() {
  sheet_();
  folder_();
}

function doGet() {
  try {
    return json_({ ok: true, records: list_() });
  } catch (err) {
    return json_({ ok: false, error: String(err.message || err) });
  }
}

function doPost(e) {
  try {
    const body = JSON.parse((e.postData && e.postData.contents) || '{}');
    if (WRITE_KEY && body.key !== WRITE_KEY) return json_({ ok: false, error: '비밀번호가 맞지 않아요' });
    if (body.action === 'upload') return json_({ ok: true, id: upload_(body) });

    const lock = LockService.getScriptLock();
    lock.waitLock(20000);
    try {
      if (body.action === 'save') return json_({ ok: true, record: save_(body.record, body.removed || []) });
      if (body.action === 'delete') { delete_(body.id); return json_({ ok: true }); }
    } finally {
      lock.releaseLock();
    }
    return json_({ ok: false, error: '알 수 없는 요청이에요' });
  } catch (err) {
    return json_({ ok: false, error: String(err.message || err) });
  }
}

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold');
    sh.setFrozenRows(1);
    // 날짜·시각 칸은 글자 그대로 두기
    ['date', 'createdAt', 'updatedAt'].forEach(h => {
      const col = HEADERS.indexOf(h) + 1;
      sh.getRange(1, col, sh.getMaxRows(), 1).setNumberFormat('@');
    });
  }
  return sh;
}

function list_() {
  const sh = sheet_();
  const last = sh.getLastRow();
  if (last < 2) return [];
  const tz = Session.getScriptTimeZone();
  const vals = sh.getRange(2, 1, last - 1, HEADERS.length).getValues();
  return vals.filter(r => r[0] !== '').map(r => {
    const o = {};
    HEADERS.forEach((h, i) => {
      let v = r[i];
      if (v instanceof Date) v = Utilities.formatDate(v, tz, h === 'date' ? 'yyyy-MM-dd' : "yyyy-MM-dd'T'HH:mm:ss");
      o[h] = v;
    });
    o.id = String(o.id);
    o.lat = Number(o.lat);
    o.lng = Number(o.lng);
    o.rating = Number(o.rating) || 0;
    o.photos = String(o.photos || '').split(',').map(s => s.trim()).filter(Boolean);
    o.tags = String(o.tags || '').split(',').map(s => s.trim()).filter(Boolean);
    return o;
  });
}

function findRow_(sh, id) {
  const last = sh.getLastRow();
  if (last < 2) return -1;
  const ids = sh.getRange(2, 1, last - 1, 1).getValues().map(r => String(r[0]));
  const i = ids.indexOf(String(id));
  return i < 0 ? -1 : i + 2;
}

function save_(rec, removed) {
  if (!rec || !rec.id || !rec.name) throw new Error('이름이 비어 있어요');
  const sh = sheet_();
  const row = HEADERS.map(h => {
    const v = rec[h];
    if (Array.isArray(v)) return v.join(',');
    return v == null ? '' : v;
  });
  const r = findRow_(sh, rec.id);
  if (r > 0) sh.getRange(r, 1, 1, row.length).setValues([row]);
  else sh.appendRow(row);
  trash_(removed);
  return rec;
}

function delete_(id) {
  const sh = sheet_();
  const r = findRow_(sh, id);
  if (r < 0) return;
  const photos = String(sh.getRange(r, HEADERS.indexOf('photos') + 1).getValue() || '').split(',').filter(Boolean);
  sh.deleteRow(r);
  trash_(photos);
}

function trash_(ids) {
  (ids || []).forEach(id => {
    if (/^[\w-]{20,}$/.test(id)) {
      try { DriveApp.getFileById(id).setTrashed(true); } catch (e) {}
    }
  });
}

function folder_() {
  const it = DriveApp.getFoldersByName(FOLDER_NAME);
  return it.hasNext() ? it.next() : DriveApp.createFolder(FOLDER_NAME);
}

function upload_(b) {
  if (!b.data) throw new Error('사진 데이터가 없어요');
  const blob = Utilities.newBlob(Utilities.base64Decode(b.data), b.mimeType || 'image/jpeg', b.filename || ('photo_' + Date.now() + '.jpg'));
  const f = folder_().createFile(blob);
  f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return f.getId();
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
