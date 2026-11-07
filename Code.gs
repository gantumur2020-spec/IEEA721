/* =====================================================================
   I.EEA721 «Цахилгаан станцын орчин үеийн технологи» — оноо хадгалах сервер
   Google Apps Script (Google Sheets-тэй холбоотой). Үнэгүй.
   Тохиргоо: Project Settings → Script Properties → PW = багшийн нууц үг
   Хуудсууд: Log (бүх үйлдэл), Students (оюутны жагсаалт: код, нэр, анги),
             BD (бие даалтын илгээлт ба дүн). Хавсаргасан файл Google Drive-ийн
             «EEA721_Бие_даалт» хавтаст хадгалагдана (багшид л харагдана).
   ===================================================================== */
const SHEET = 'Log';
const HEAD = ['Огноо','SID','Код','Нэр','Анги','Төрөл','Түлхүүр','Оноо','Нийт','Секунд','Өдөр','Лекц','Асуулт'];
const ROSTER_SHEET = 'Students';
const BD_SHEET = 'BD';
const BD_HEAD = ['Огноо','Код','Нэр','Анги','БД','Бичвэр','Файл','Файлын нэр','Дүн','Тайлбар','Дүгнэсэн'];
const BD_N = 4;

function ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }
function sh_() {
  let s = ss_().getSheetByName(SHEET);
  if (!s) { s = ss_().insertSheet(SHEET); s.appendRow(HEAD); s.setFrozenRows(1); }
  return s;
}
function rosterSheet_() {
  let s = ss_().getSheetByName(ROSTER_SHEET);
  if (!s) { s = ss_().insertSheet(ROSTER_SHEET); s.appendRow(['Код', 'Овог нэр', 'Сургууль/анги']); s.setFrozenRows(1); }
  return s;
}
function bdSheet_() {
  let s = ss_().getSheetByName(BD_SHEET);
  if (!s) { s = ss_().insertSheet(BD_SHEET); s.appendRow(BD_HEAD); s.setFrozenRows(1); }
  return s;
}
function bdFolder_() {
  const id = props_().getProperty('BD_FOLDER');
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) {} }
  const f = DriveApp.createFolder('EEA721_Бие_даалт');
  props_().setProperty('BD_FOLDER', f.getId());
  return f;
}
/* код + БД-ийн хамгийн сүүлийн мөр (мөрийн дугаар, утга) */
function bdLatest_(code, n) {
  const v = bdSheet_().getDataRange().getValues();
  for (let i = v.length - 1; i >= 1; i--) if (norm_(v[i][1]) === norm_(code) && Number(v[i][4]) === n) return { row: i + 1, v: v[i] };
  return null;
}
function bdItem_(r) {
  return { at: r[0] instanceof Date ? r[0].toISOString() : String(r[0]), text: String(r[5] || ''), url: String(r[6] || ''),
    fname: String(r[7] || ''), grade: r[8] === '' || r[8] == null ? '' : Number(r[8]), comment: String(r[9] || '') };
}
function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
function clip_(v, n) { return String(v == null ? '' : v).slice(0, n); }
function norm_(s) { return String(s || '').toUpperCase().replace(/\s+/g, '').replace(/Ё/g, 'Е'); }
function props_() { return PropertiesService.getScriptProperties(); }
function pwOk_(pw) { const p = props_().getProperty('PW'); return !!p && pw === p; }

function roster_() {
  const v = rosterSheet_().getDataRange().getValues();
  const out = [];
  for (let i = 1; i < v.length; i++) if (v[i][0]) out.push({ code: String(v[i][0]).trim(), name: String(v[i][1]).trim(), group: String(v[i][2] || '').trim() });
  return out;
}
function findStudent_(code) {
  const c = norm_(code);
  return roster_().find(r => norm_(r.code) === c) || null;
}
function theme_() {
  try { return JSON.parse(props_().getProperty('THEME') || 'null'); } catch (e) { return null; }
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    const p = JSON.parse(e.postData.contents);
    const act = String(p.action || '');
    /* --- багшийн үйлдэл --- */
    if (act === 'grade') {
      if (!pwOk_(p.pw)) return out_({ ok: false, error: 'bad_pw' });
      const n = Number(p.bd), L = bdLatest_(p.code, n);
      if (!L) return out_({ ok: false, error: 'no_submission' });
      const g = String(p.grade == null ? '' : p.grade).trim();
      if (g !== '' && !(Number(g) >= 0 && Number(g) <= 5)) return out_({ ok: false, error: 'bad_grade' });
      bdSheet_().getRange(L.row, 9, 1, 3).setValues([[g === '' ? '' : Number(g), clip_(p.comment, 500), g === '' ? '' : new Date()]]);
      return out_({ ok: true });
    }
    if (act === 'setRoster' || act === 'setTheme') {
      if (!pwOk_(p.pw)) return out_({ ok: false, error: 'bad_pw' });
      if (act === 'setTheme') {
        const t = p.theme || {};
        const clean = { id: clip_(t.id, 20) };
        if (/^#[0-9a-fA-F]{6}$/.test(t.accent || '')) clean.accent = t.accent;
        props_().setProperty('THEME', JSON.stringify(clean));
        return out_({ ok: true });
      }
      const list = (p.roster || []).slice(0, 2000).map(r => [clip_(r.code, 20).trim(), clip_(r.name, 80).trim(), clip_(r.group, 60).trim()]).filter(r => r[0] && r[1]);
      const s = rosterSheet_();
      if (s.getLastRow() > 1) s.getRange(2, 1, s.getLastRow() - 1, 3).clearContent();
      if (list.length) s.getRange(2, 1, list.length, 3).setValues(list);
      return out_({ ok: true, count: list.length });
    }
    /* --- оюутны үйлдэл: зөвхөн жагсаалтад байгаа код --- */
    if (['register', 'study', 'result', 'bd'].indexOf(act) < 0) return out_({ ok: false, error: 'bad_action' });
    const st = findStudent_(p.code);
    if (!st) return out_({ ok: false, error: 'not_enrolled' });
    if (act === 'bd') {
      const n = Number(p.bd);
      if (!(n >= 1 && n <= BD_N)) return out_({ ok: false, error: 'bad_bd' });
      const L = bdLatest_(st.code, n);
      if (L && L.v[8] !== '' && L.v[8] != null) return out_({ ok: false, error: 'graded' });
      let url = '', fname = '';
      if (p.file && p.file.data) {
        const bytes = Utilities.base64Decode(p.file.data);
        if (bytes.length > 8 * 1024 * 1024) return out_({ ok: false, error: 'too_big' });
        fname = clip_(p.file.name, 120);
        const blob = Utilities.newBlob(bytes, clip_(p.file.type, 100) || 'application/octet-stream', 'BD' + n + '_' + st.code + '_' + fname);
        url = bdFolder_().createFile(blob).getUrl();
      }
      bdSheet_().appendRow([new Date(), st.code, st.name, st.group, n, clip_(p.text, 45000), url, fname, '', '', '']);
      return out_({ ok: true, file: !!url });
    }
    const s = sh_();
    const who = st.code;
    if (act === 'result' && p.key === 'final') {
      const v = s.getDataRange().getValues();
      for (let i = 1; i < v.length; i++) {
        if (norm_(v[i][2]) === norm_(who) && v[i][5] === 'result' && v[i][6] === 'final')
          return out_({ ok: false, error: 'final_exists' });
      }
    }
    s.appendRow([new Date(), clip_(p.sid, 40), who, st.name, st.group, act,
      clip_(p.key, 10), p.score == null ? '' : Number(p.score), p.total == null ? '' : Number(p.total),
      p.sec == null ? '' : Number(p.sec), clip_(p.date, 10),
      (p.lec || []).map(Number).join(','), (p.items || []).join(',')]);
    return out_({ ok: true });
  } catch (err) {
    return out_({ ok: false, error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
}

function doGet(e) {
  const prm = (e && e.parameter) || {};
  if (prm.action === 'settings') return out_({ ok: true, theme: theme_() });
  if (prm.action === 'login') {
    const st = findStudent_(prm.code);
    if (!st || norm_(st.name) !== norm_(prm.name)) return out_({ ok: false, error: 'not_found' });
    return out_({ ok: true, code: st.code, name: st.name, group: st.group });
  }
  if (prm.action === 'mybd') {
    const st = findStudent_(prm.code);
    if (!st || norm_(st.name) !== norm_(prm.name)) return out_({ ok: false, error: 'not_found' });
    const items = {};
    for (let n = 1; n <= BD_N; n++) {
      const L = bdLatest_(st.code, n);
      if (L) { const it = bdItem_(L.v); items[n] = { at: it.at, file: !!it.url, grade: it.grade, comment: it.comment }; }
    }
    return out_({ ok: true, items: items });
  }
  if (prm.action === 'teacher') {
    if (!pwOk_(prm.pw)) return out_({ ok: false, error: 'bad_pw' });
    return out_({ ok: true, rows: aggregate_(), roster: roster_(), theme: theme_() });
  }
  return out_({ ok: true, msg: 'I.EEA721 server ажиллаж байна' });
}

function aggregate_() {
  const m = {};
  roster_().forEach(r => { m[norm_(r.code)] = { _id: r.code, name: r.name, group: r.group, code: r.code, study: {} }; });
  const v = sh_().getDataRange().getValues();
  for (let i = 1; i < v.length; i++) {
    const r = v[i];
    const key = norm_(r[2]);
    if (!key || !m[key]) continue;           // жагсаалтаас хасагдсан оюутныг тооцохгүй
    const o = m[key];
    const at = r[0] instanceof Date ? r[0].toISOString() : String(r[0]);
    if (r[5] === 'study' && r[10]) {
      const d = o.study[r[10]] || (o.study[r[10]] = { sec: 0, lec: [] });
      d.sec += Number(r[9]) || 0;
      String(r[11] || '').split(',').filter(String).forEach(x => { x = Number(x); if (d.lec.indexOf(x) < 0) d.lec.push(x); });
      d.lec.sort((a, b) => a - b);
    } else if (r[5] === 'result') {
      const k = String(r[6]), sc = Number(r[7]) || 0;
      if (k === 'final') {
        if (!o.final) o.final = { score: sc, at: at, sec: Number(r[9]) || 0 };
      } else if (/^t\d\d$/.test(k)) {
        const t = o[k] || (o[k] = { best: 0, attempts: 0, at: '' });
        t.best = Math.max(t.best, sc); t.attempts++; t.at = at; t.last = sc;
      }
    }
  }
  const b = bdSheet_().getDataRange().getValues();
  for (let i = 1; i < b.length; i++) {          // доош нь уншиж сүүлийн илгээлтийг үлдээнэ
    const key = norm_(b[i][1]);
    if (!key || !m[key]) continue;
    const o = m[key]; o.bd = o.bd || {};
    o.bd[Number(b[i][4])] = bdItem_(b[i]);
  }
  return Object.keys(m).map(k => m[k]);
}

/* Эхний жагсаалт: Apps Script засварлагч дээр энэ функцийг нэг удаа ▶ Run хийнэ (Drive-ийн зөвшөөрлийг давхар асууна). */
function seedRoster() {
  // Жишээ мөр — өөрийн ангийн жагсаалтаар солино (эсвэл вэбийн «Тохиргоо» хэсгээс Excel-ээс хуулж оруулна).
  const list = [
    ['B000000001', 'Т.Туршилт', 'Туршилт']
  ];
  const s = rosterSheet_();
  if (s.getLastRow() > 1) s.getRange(2, 1, s.getLastRow() - 1, 3).clearContent();
  s.getRange(2, 1, list.length, 3).setValues(list);
}
