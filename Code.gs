/**
 * 姫の手帳 — 스프레드시트 백엔드 (Google Apps Script)
 *
 * 하는 일
 *  1) 스프레드시트의 일정·할일·메모·루틴·살것 탭을 사이트와 주고받기
 *  2) 구글 캘린더 일정을 읽어서 사이트 캘린더에 보여주기 (읽기 전용)
 *  3) 구글 드라이브 '허브_인박스' 폴더에 Claude가 넣은 JSON 파일을 스프레드시트로 옮기기
 *
 * 사이트의 모든 내용(일정·할일·메모·루틴·살것·영양제·사이트 문구)은 이 스프레드시트에서 고칠 수 있어요.
 *
 * 처음 한 번: 위쪽 함수 목록에서 setup 선택 → 실행 → 권한 허용
 * 그다음: 배포 → 새 배포 → 웹 앱 (실행: 나 / 액세스: 모든 사용자)
 */

const INBOX_FOLDER = '허브_인박스';
const TZ = 'Asia/Seoul';

const SHEETS = {
  events: { name: '일정', cols: [['id','id'],['date','날짜'],['time','시간'],['title','제목'],['cat','분류'],['note','메모'],['sticker','스티커']] },
  todos:  { name: '할일', cols: [['id','id'],['title','제목'],['cat','분류'],['due','마감일'],['important','중요'],['done','완료'],['doneAt','완료시각'],['created','생성']] },
  notes:  { name: '메모', cols: [['id','id'],['title','제목'],['body','내용'],['cat','분류'],['pinned','고정'],['updated','수정시각']] },
  habits: { name: '루틴', cols: [['id','id'],['name','이름'],['cat','분류'],['log','기록']] },
  shop:   { name: '살것', cols: [['id','id'],['title','품목'],['done','완료'],['created','생성']] },
  supps:  { name: '영양제', cols: [['id','id'],['time','시간'],['name','영양제'],['bold','강조'],['note','메모']],
            seed: [['아침 공복','유산균',''],['점심 식후','오메가3',''],['점심 식후','비타민 D3',''],['점심 식후','비타민 C',''],['저녁~취침 전','마그네슘','']] },
};
const SITE_SHEET = '사이트';
const ARCHIVE_SHEET = '보관';
const ARCHIVE_DAYS = 90;   // 완료한 지 이 일수가 지난 할일은 '보관' 탭으로 옮겨요
const SITE_ROWS = [['title','사이트 제목','姫の手帳'],['footer','메뉴 아래 문구','업무 · 개인 · 취미'],['suppTitle','영양제 팝업 제목','영양제']];
const CAT_KO = { work: '업무', life: '개인', play: '취미' };
const CAT_EN = { '업무': 'work', '개인': 'life', '취미': 'play', work: 'work', life: 'life', play: 'play' };
// 스티커: 시트에는 한글 이름으로 적어요. 비우면 제목을 보고 자동으로 골라요.
const STICKER_KO = { k_heart:'키티하트', k_wink:'키티윙크', k_bear:'키티곰', k_hug:'키티포옹', k_shy:'키티쑥스', k_sleep:'키티잠', k_lie:'키티', k_doc:'키티의사',
  laptop:'노트북', books:'공부', openbook:'기록', cup:'카페', cosmetics:'화장품', bag:'쇼핑', suitcase:'짐', mirror:'거울',
  hospital:'병원', syringe:'주사', tooth:'치과', pills:'약', dumbbell:'운동', forkknife:'식사', bath:'목욕', mask:'피부',
  plane:'여행', crown:'생일', alarm:'마감', bangbang:'시험', star:'촬영', heart:'하트', bow:'리본', sakura:'벚꽃', clover:'행운',
  note:'음악', sun:'맑음', cloud:'구름', drop:'물', calendar:'일정', clipboard:'업무', none:'없음' };
const STICKER_EN = Object.keys(STICKER_KO).reduce((m, k) => (m[STICKER_KO[k]] = k, m[k] = k, m), {});
const BOOL_FIELDS = ['important', 'done', 'pinned', 'bold'];
const TIME_FIELDS = ['doneAt', 'created', 'updated'];

/* ---------------- 처음 한 번 실행 ---------------- */
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Object.keys(SHEETS).forEach(sheet_);
  siteSheet_();
  const folder = inboxFolder_(true);
  const props = PropertiesService.getScriptProperties();
  let token = props.getProperty('TOKEN');
  if (!token) {
    token = Utilities.getUuid().replace(/-/g, '').slice(0, 24);
    props.setProperty('TOKEN', token);
  }
  const guide = ss.getSheetByName('안내') || ss.insertSheet('안내', 0);
  guide.clear();
  guide.getRange(1, 1, 6, 2).setValues([
    ['사이트 토큰', token],
    ['인박스 폴더', folder.getUrl()],
    ['분류 값', '업무 / 개인 / 취미'],
    ['예/아니오 값', 'O 이면 예, 비우면 아니오'],
    ['날짜 형식', '2026-09-30'],
    ['루틴 기록', '체크한 날짜를 쉼표로 구분 (예: 2026-09-21, 2026-09-22)'],
  ]);
  guide.getRange('A1:A6').setFontWeight('bold');
  guide.setColumnWidth(1, 120);
  guide.setColumnWidth(2, 420);
  const s1 = ss.getSheetByName('시트1') || ss.getSheetByName('Sheet1');
  if (s1 && s1.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(s1);
  Logger.log('토큰: ' + token);
}

/* ---------------- 웹 앱 입구 ---------------- */
function doGet(e) { return handle_((e && e.parameter) || {}); }
function doPost(e) {
  let body = {};
  try { body = JSON.parse(e.postData.contents); } catch (err) { return json_({ error: 'bad_request' }); }
  return handle_(body);
}

function handle_(p) {
  if (!auth_(p.token)) return json_({ error: 'unauthorized' });
  try {
    if (p.action === 'ping') return json_({ ok: true });
    if (p.action === 'gcal') return json_({ ok: true, events: gcal_(p.from, p.to) });

    const lock = LockService.getScriptLock();
    lock.waitLock(20000);
    try {
      if (p.action === 'load') {
        archiveIfDue_();
        const imported = processInbox_();
        return json_(Object.assign({ ok: true, imported: imported }, loadFast_()));
      }
      if (p.action === 'ops') {
        (p.ops || []).forEach(o => {
          if (!SHEETS[o.key]) return;
          if (o.op === 'upsert' && o.item && o.item.id) upsert_(o.key, o.item);
          else if (o.op === 'delete' && o.id) delete_(o.key, o.id);
        });
        return json_({ ok: true });
      }
    } finally {
      lock.releaseLock();
    }
    return json_({ error: 'unknown_action' });
  } catch (err) {
    return json_({ error: 'server', message: String((err && err.message) || err) });
  }
}

function auth_(t) {
  const tok = PropertiesService.getScriptProperties().getProperty('TOKEN');
  return !!tok && String(t || '') === tok;
}
function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

/* ---------------- 시트 읽기/쓰기 ---------------- */
function sheet_(key) {
  const def = SHEETS[key];
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let s = ss.getSheetByName(def.name);
  if (!s) {
    s = ss.insertSheet(def.name);
    s.getRange(1, 1, s.getMaxRows(), def.cols.length).setNumberFormat('@');
    s.getRange(1, 1, 1, def.cols.length).setValues([def.cols.map(c => c[1])]).setFontWeight('bold');
    s.setFrozenRows(1);
    s.setColumnWidth(1, 90);
    if (def.seed) {
      const rows = def.seed.map(r => [newId_()].concat(r, Array(def.cols.length - 1 - r.length).fill('')));
      s.getRange(2, 1, rows.length, def.cols.length).setValues(rows);
    }
  }
  if (s.getLastColumn() < def.cols.length) {       // 새 칸이 추가된 경우 제목 줄 보충
    const from = s.getLastColumn() + 1, cnt = def.cols.length - from + 1;
    s.getRange(1, from, s.getMaxRows(), cnt).setNumberFormat('@');
    s.getRange(1, from, 1, cnt).setValues([def.cols.slice(from - 1).map(c => c[1])]).setFontWeight('bold');
  }
  return s;
}

function siteSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let s = ss.getSheetByName(SITE_SHEET);
  if (!s) {
    s = ss.insertSheet(SITE_SHEET);
    s.getRange(1, 1, 50, 2).setNumberFormat('@');
    s.getRange(1, 1, 1, 2).setValues([['항목', '값']]).setFontWeight('bold');
    s.getRange(2, 1, SITE_ROWS.length, 2).setValues(SITE_ROWS.map(r => [r[1], r[2]]));
    s.setFrozenRows(1);
    s.setColumnWidth(1, 160);
    s.setColumnWidth(2, 320);
  }
  return s;
}
function readSite_() {
  const s = siteSheet_();
  const n = s.getLastRow();
  const vals = n > 1 ? s.getRange(2, 1, n - 1, 2).getDisplayValues() : [];
  const out = {};
  SITE_ROWS.forEach(([key, label, def]) => {
    const hit = vals.find(r => String(r[0]).trim() === label);
    out[key] = hit && String(hit[1]).trim() ? String(hit[1]).trim() : def;
  });
  return out;
}

// 시트 목록을 한 번만 가져오고, 탭마다 한 번에 통째로 읽어서 왕복 횟수를 줄여요
function loadFast_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const byName = {};
  ss.getSheets().forEach(s => byName[s.getName()] = s);
  const data = {};
  Object.keys(SHEETS).forEach(k => {
    const def = SHEETS[k], s = byName[def.name];
    if (!s) { data[k] = readAll_(k); return; }             // 탭이 없으면 기존 방식으로 만들면서 읽기
    const vals = s.getDataRange().getDisplayValues();
    if ((vals[0] || []).length < def.cols.length) { data[k] = readAll_(k); return; }
    const out = [];
    for (let i = 1; i < vals.length; i++) {
      const r = vals[i].slice(0, def.cols.length);
      if (r.every(v => String(v).trim() === '')) continue;
      if (!r[0]) { r[0] = newId_(); s.getRange(i + 1, 1).setNumberFormat('@').setValue(r[0]); }
      out.push(fromRow_(k, r));
    }
    data[k] = out;
  });
  let site;
  const ssite = byName[SITE_SHEET];
  if (ssite) {
    const vals = ssite.getDataRange().getDisplayValues().slice(1);
    site = {};
    SITE_ROWS.forEach(([key, label, def]) => {
      const hit = vals.find(r => String(r[0]).trim() === label);
      site[key] = hit && String(hit[1] || '').trim() ? String(hit[1]).trim() : def;
    });
  } else site = readSite_();
  return { data: data, site: site };
}

function readAll_(key) {
  const def = SHEETS[key], s = sheet_(key);
  const n = s.getLastRow();
  if (n < 2) return [];
  const vals = s.getRange(2, 1, n - 1, def.cols.length).getDisplayValues();
  const out = [];
  vals.forEach((r, i) => {
    if (r.every(v => String(v).trim() === '')) return;
    if (!r[0]) {                       // 시트에서 직접 추가한 줄 → id 붙이기
      r[0] = newId_();
      s.getRange(i + 2, 1).setNumberFormat('@').setValue(r[0]);
    }
    out.push(fromRow_(key, r));
  });
  return out;
}

function findRow_(s, id) {
  const n = s.getLastRow();
  if (n < 2) return 0;
  const hit = s.getRange(2, 1, n - 1, 1).createTextFinder(String(id)).matchEntireCell(true).findNext();
  return hit ? hit.getRow() : 0;
}
function upsert_(key, item) {
  const def = SHEETS[key], s = sheet_(key);
  const row = findRow_(s, item.id) || s.getLastRow() + 1;
  s.getRange(row, 1, 1, def.cols.length).setNumberFormat('@').setValues([toRow_(key, item)]);
}
function delete_(key, id) {
  const s = sheet_(key);
  const row = findRow_(s, id);
  if (row) s.deleteRow(row);
}

function fromRow_(key, r) {
  const o = {};
  SHEETS[key].cols.forEach(([f], i) => {
    const v = String(r[i] == null ? '' : r[i]).trim();
    if (f === 'cat') o[f] = CAT_EN[v] || 'life';
    else if (f === 'sticker') o[f] = STICKER_EN[v] || '';
    else if (BOOL_FIELDS.indexOf(f) >= 0) o[f] = /^(o|y|yes|true|1|✓|v|예)$/i.test(v);
    else if (TIME_FIELDS.indexOf(f) >= 0) o[f] = parseStamp_(v);
    else if (f === 'log') { o.log = {}; v.split(/[,\s]+/).map(normDate_).filter(Boolean).forEach(d => o.log[d] = true); }
    else if (f === 'date' || f === 'due') o[f] = normDate_(v);
    else if (f === 'time') o[f] = normTime_(v);
    else o[f] = v;
  });
  return o;
}
function toRow_(key, o) {
  return SHEETS[key].cols.map(([f]) => {
    const v = o[f];
    if (f === 'cat') return CAT_KO[v] || '개인';
    if (f === 'sticker') return STICKER_KO[v] || '';
    if (BOOL_FIELDS.indexOf(f) >= 0) return v ? 'O' : '';
    if (TIME_FIELDS.indexOf(f) >= 0) return v ? Utilities.formatDate(new Date(Number(v)), TZ, 'yyyy-MM-dd HH:mm') : '';
    if (f === 'log') return Object.keys(v || {}).filter(d => v[d]).sort().join(', ');
    return v == null ? '' : String(v);
  });
}

function parseStamp_(v) {
  if (!v) return null;
  if (/^\d{10,}$/.test(v)) return Number(v);
  const m = v.match(/^(\d{4})-(\d{1,2})-(\d{1,2})(?:[ T](\d{1,2}):(\d{2}))?/);
  if (!m) return null;
  const iso = m[1] + '-' + pad_(m[2]) + '-' + pad_(m[3]) + 'T' + pad_(m[4] || 0) + ':' + (m[5] || '00') + ':00+09:00';
  const t = new Date(iso).getTime();
  return isNaN(t) ? null : t;
}
function normDate_(v) {
  v = String(v || '').trim();
  if (!v) return '';
  let m = v.match(/^(\d{4})[-./](\d{1,2})[-./](\d{1,2})/);
  if (m) return m[1] + '-' + pad_(m[2]) + '-' + pad_(m[3]);
  m = v.match(/^(\d{1,2})[/.](\d{1,2})$/);          // 9/30 → 올해
  if (m) return Utilities.formatDate(new Date(), TZ, 'yyyy') + '-' + pad_(m[1]) + '-' + pad_(m[2]);
  return '';
}
function normTime_(v) {   // 08:00 같은 시각은 맞춰 쓰고, '아침' 같은 글자는 그대로 둬요
  const s = String(v || '').trim();
  let m = s.match(/^(\d{1,2}):?(\d{2})(?!\d)/);
  if (m) return pad_(m[1]) + ':' + m[2];
  m = s.match(/ (\d{2}):(\d{2}):\d{2} GMT/);   // 시트가 시각을 날짜로 바꿔 둔 경우
  if (m) return m[1] + ':' + m[2];
  return s;
}
function pad_(n) { return ('0' + n).slice(-2); }
function newId_() { return Date.now().toString(36) + Math.random().toString(36).slice(2, 7); }

/* ---------------- 오래된 완료 할일 보관 ---------------- */
// 사이트를 열 때 하루 한 번만 확인해요. 편집기에서 archiveNow를 실행하면 바로 옮겨요.
function archiveIfDue_() {
  const props = PropertiesService.getScriptProperties();
  const today = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
  if (props.getProperty('ARCHIVED_ON') === today) return 0;
  const n = archiveOld_();
  props.setProperty('ARCHIVED_ON', today);
  return n;
}
function archiveNow() {
  const n = archiveOld_();
  Logger.log('보관 탭으로 옮긴 할일: ' + n + '개');
}
function archiveSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let a = ss.getSheetByName(ARCHIVE_SHEET);
  if (!a) {
    const cols = SHEETS.todos.cols;
    a = ss.insertSheet(ARCHIVE_SHEET);
    a.getRange(1, 1, a.getMaxRows(), cols.length).setNumberFormat('@');
    a.getRange(1, 1, 1, cols.length).setValues([cols.map(c => c[1])]).setFontWeight('bold');
    a.setFrozenRows(1);
  }
  return a;
}
function archiveOld_() {
  const cols = SHEETS.todos.cols, s = sheet_('todos');
  const n = s.getLastRow();
  if (n < 2) return 0;
  const vals = s.getRange(2, 1, n - 1, cols.length).getDisplayValues();
  const cutoff = Date.now() - ARCHIVE_DAYS * 86400000;
  const move = [], rows = [];
  vals.forEach((r, i) => {
    const it = fromRow_('todos', r);
    if (it.done && it.doneAt && it.doneAt < cutoff) { move.push(r); rows.push(i + 2); }
  });
  if (!move.length) return 0;
  const a = archiveSheet_();
  a.getRange(a.getLastRow() + 1, 1, move.length, cols.length).setNumberFormat('@').setValues(move);
  // 아래쪽부터, 붙어 있는 행은 한 번에 삭제
  for (let i = rows.length - 1; i >= 0;) {
    let j = i;
    while (j > 0 && rows[j - 1] === rows[j] - 1) j--;
    s.deleteRows(rows[j], i - j + 1);
    i = j - 1;
  }
  return move.length;
}

/* ---------------- 구글 캘린더 (읽기 전용) ---------------- */
function gcal_(from, to) {
  const f = normDate_(from), t = normDate_(to);
  if (!f || !t) return [];
  const start = new Date(f + 'T00:00:00+09:00'), end = new Date(t + 'T00:00:00+09:00');
  const out = [];
  CalendarApp.getAllCalendars().forEach(cal => {
    try { if (cal.isHidden()) return; } catch (e) {}
    const hol = /holiday/i.test(cal.getId());
    let evs = [];
    try { evs = cal.getEvents(start, end); } catch (e) { return; }
    evs.forEach(ev => {
      const base = { id: ev.getId(), title: ev.getTitle() || '(제목 없음)', hol: hol, note: ev.getLocation() || '' };
      if (ev.isAllDayEvent()) {
        let cur = Utilities.formatDate(ev.getAllDayStartDate(), TZ, 'yyyy-MM-dd');
        const last = Utilities.formatDate(ev.getAllDayEndDate(), TZ, 'yyyy-MM-dd');
        let guard = 0;
        do {
          out.push(Object.assign({}, base, { id: base.id + '_' + cur, date: cur, time: '' }));
          cur = nextDay_(cur);
        } while (cur < last && guard++ < 62);
      } else {
        const s = ev.getStartTime();
        const d = Utilities.formatDate(s, TZ, 'yyyy-MM-dd');
        out.push(Object.assign({}, base, { id: base.id + '_' + d, date: d, time: Utilities.formatDate(s, TZ, 'HH:mm') }));
      }
    });
  });
  return out;
}
function nextDay_(d) {
  return Utilities.formatDate(new Date(new Date(d + 'T12:00:00+09:00').getTime() + 86400000), TZ, 'yyyy-MM-dd');
}

/* ---------------- Claude 인박스 ---------------- */
function inboxFolder_(create) {
  // 폴더를 이름으로 찾는 게 느려서, 한 번 찾으면 id를 기억해 두고 바로 열어요
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('INBOX_ID');
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) { props.deleteProperty('INBOX_ID'); } }
  const it = DriveApp.getFoldersByName(INBOX_FOLDER);
  const f = it.hasNext() ? it.next() : (create ? DriveApp.createFolder(INBOX_FOLDER) : null);
  if (f) props.setProperty('INBOX_ID', f.getId());
  return f;
}

function processInbox_() {
  const folder = inboxFolder_(false);
  if (!folder) return 0;
  const files = folder.getFiles();
  let n = 0;
  while (files.hasNext()) {
    const f = files.next(), name = f.getName();
    if (name.indexOf('[오류]') === 0) continue;
    let text = '';
    try {
      text = f.getMimeType() === MimeType.GOOGLE_DOCS
        ? DocumentApp.openById(f.getId()).getBody().getText()
        : f.getBlob().getDataAsString('UTF-8');
    } catch (e) { continue; }
    try {
      text = text.trim().replace(/^```(?:json)?\s*/i, '').replace(/\s*```$/, '');
      const j = JSON.parse(text);
      const items = Array.isArray(j) ? j : (j.items || [j]);
      items.forEach(x => { if (addInbox_(x)) n++; });
      f.setTrashed(true);
    } catch (e) {
      f.setName('[오류] ' + name);   // 형식이 틀린 파일은 이름만 바꿔두고 건너뜀
    }
  }
  return n;
}

function addInbox_(x) {
  if (!x || typeof x !== 'object') return false;
  const type = String(x.type || '').trim().toLowerCase();
  const KEY = { event: 'events', '일정': 'events', todo: 'todos', '할일': 'todos', someday: 'todos', '언젠가': 'todos',
                note: 'notes', memo: 'notes', '메모': 'notes', shop: 'shop', '살것': 'shop', '살 것': 'shop',
                habit: 'habits', '루틴': 'habits', supp: 'supps', '영양제': 'supps' }[type];
  if (!KEY) return false;
  const cat = CAT_EN[String(x.cat || '').trim()] || 'life';
  const now = Date.now();
  let item;
  if (KEY === 'events') {
    const d = normDate_(x.date);
    if (!x.title || !d) return false;
    item = { id: newId_(), title: String(x.title), date: d, time: normTime_(x.time), cat: cat, note: String(x.note || ''),
             sticker: STICKER_EN[String(x.sticker || '').trim()] || '' };
  } else if (KEY === 'todos') {
    if (!x.title) return false;
    const someday = type === 'someday' || type === '언젠가';
    item = { id: newId_(), title: String(x.title), cat: cat, due: someday ? '' : normDate_(x.due),
             important: !!x.important, done: false, doneAt: null, created: now };
  } else if (KEY === 'notes') {
    if (!x.title && !x.body) return false;
    item = { id: newId_(), title: String(x.title || ''), body: String(x.body || ''), cat: cat, pinned: !!x.pinned, updated: now };
  } else if (KEY === 'shop') {
    if (!x.title) return false;
    item = { id: newId_(), title: String(x.title), done: false, created: now };
  } else if (KEY === 'supps') {
    const nm = x.name || x.title;
    if (!nm) return false;
    item = { id: newId_(), time: String(x.time || ''), name: String(nm), bold: !!x.bold, note: String(x.note || '') };
  } else {
    const nm = x.name || x.title;
    if (!nm) return false;
    item = { id: newId_(), name: String(nm), cat: cat, log: {} };
  }
  upsert_(KEY, item);
  return true;
}
