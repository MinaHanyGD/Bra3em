/**
 * خدمة براعم — الحضور والملاحظات على Google Sheet
 * -------------------------------------------------
 * الملف ده بيتحط في Apps Script الخاص بشيت الحضور (Extensions ← Apps Script)
 * وبيتنشر كـ Web App. الموقع بيقرا منه وبيكتب عليه من خلاله.
 *
 * التابات اللي بيعملها لوحده:
 *   «حضور <اسم الفصل>» : صف لكل طفل (الكود، الاسم) وعمود لكل تاريخ اجتماع (ح / غ / ع)
 *   «الملاحظات»        : ملاحظة في كل صف (الكود، الفصل، الاسم، التاريخ، الملاحظة…)
 */

// ✏️ غيّر الكلمة دي لأي كلمة طويلة من عندك، وحط نفس الكلمة في CONFIG.key جوه app.js
const KEY = 'غيّر-الكلمة-دي';

const TZ = 'Africa/Cairo';
const PREFIX = 'حضور ';
const NOTES = 'الملاحظات';
const NOTE_HDR = ['المعرّف', 'الكود', 'الفصل', 'الاسم', 'التاريخ', 'الملاحظة', 'وقت التسجيل', 'ts'];
const VALID = ['ح', 'غ', 'ع', ''];

/* ---------------- نقاط الدخول ---------------- */

function doGet(e) {
  if (!e || !e.parameter || e.parameter.key !== KEY) return out_({ ok: false, err: 'auth' });
  try { return out_(readAll_()); } catch (err) { return out_({ ok: false, err: String(err && err.message || err) }); }
}

function doPost(e) {
  let b;
  try { b = JSON.parse(e.postData.contents); } catch (err) { return out_({ ok: false, err: 'bad' }); }
  if (!b || b.key !== KEY) return out_({ ok: false, err: 'auth' });
  const lock = LockService.getScriptLock();
  try { lock.waitLock(25000); } catch (err) { return out_({ ok: false, err: 'busy' }); }
  try {
    const results = [];
    if (b.sync) syncKids_(b.sync);
    if (b.ops) applyOps_(b.ops, results);
    return out_({ ok: true, results: results });
  } catch (err) {
    return out_({ ok: false, err: String(err && err.message || err) });
  } finally {
    lock.releaseLock();
  }
}

/** شغّلها مرة واحدة بإيدك (Run) عشان تدّي الصلاحيات وتجهّز تاب الملاحظات. */
function setup() { notesSheet_(); }

/* ---------------- أدوات ---------------- */

function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
function today_() { return day_(Date.now()); }
function day_(ts) { return Utilities.formatDate(new Date(ts), TZ, 'yyyy-MM-dd'); }
function dk_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, TZ, 'yyyy-MM-dd');
  return String(v == null ? '' : v).trim();
}
function ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }

/* ---------------- القراءة ---------------- */

function readAll_() {
  const tabs = {};
  ss_().getSheets().forEach(function (sh) {
    const n = sh.getName();
    if (n.indexOf(PREFIX) !== 0) return;
    const cls = n.slice(PREFIX.length).trim();
    const v = sh.getDataRange().getValues();
    const dates = [];
    const rows = [];
    if (v.length) for (let c = 2; c < v[0].length; c++) dates.push(dk_(v[0][c]));
    for (let r = 1; r < v.length; r++) {
      if (!String(v[r][0]).trim()) continue;
      const row = [String(v[r][0]).trim(), String(v[r][1] == null ? '' : v[r][1])];
      for (let i = 0; i < dates.length; i++) row.push(String(v[r][2 + i] == null ? '' : v[r][2 + i]).trim());
      rows.push(row);
    }
    tabs[cls] = { dates: dates, rows: rows };
  });
  const notes = [];
  const ns = ss_().getSheetByName(NOTES);
  if (ns && ns.getLastRow() > 1) {
    const v = ns.getRange(1, 1, ns.getLastRow(), NOTE_HDR.length).getValues();
    for (let r = 1; r < v.length; r++) {
      if (!String(v[r][0]).trim()) continue;
      notes.push({
        nid: String(v[r][0]).trim(), id: String(v[r][1]).trim(), cls: String(v[r][2]).trim(),
        name: String(v[r][3]), date: dk_(v[r][4]), text: String(v[r][5]), ts: Number(v[r][7]) || 0
      });
    }
  }
  return { ok: true, today: today_(), tabs: tabs, notes: notes };
}

/* ---------------- تابات الحضور ---------------- */

function tabSheet_(cls) {
  const ss = ss_(), name = PREFIX + cls;
  let sh = ss.getSheetByName(name);
  if (sh) return sh;
  sh = ss.insertSheet(name);
  sh.setRightToLeft(true);
  if (sh.getMaxColumns() < 80) sh.insertColumnsAfter(sh.getMaxColumns(), 80 - sh.getMaxColumns());
  sh.getRange(1, 1, 1, 2).setNumberFormat('@').setValues([['الكود', 'الاسم']]).setFontWeight('bold');
  sh.setFrozenRows(1);
  sh.setFrozenColumns(2);
  sh.setColumnWidth(2, 230);
  const rng = sh.getRange('C2:BZ1000');
  sh.setConditionalFormatRules([['ح', '#d9f2e3'], ['غ', '#fbdcdc'], ['ع', '#fdecc8']].map(function (p) {
    return SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(p[0]).setBackground(p[1]).setRanges([rng]).build();
  }));
  return sh;
}

/** بيحمّل التاب في الذاكرة عشان نعدّل فيه ونكتب الخانات اللي اتغيّرت بس */
function loadTab_(cls) {
  const sh = tabSheet_(cls);
  const v = sh.getDataRange().getValues();
  const hdr = (v[0] || ['الكود', 'الاسم']).map(dk_);
  while (hdr.length < 2) hdr.push('');
  const rows = v.slice(1).map(function (r) {
    const x = r.slice();
    while (x.length < hdr.length) x.push('');
    return x;
  });
  return { sh: sh, hdr: hdr, rows: rows, dirty: {}, newCols: [], firstNew: -1, nameChanges: [] };
}
function col_(t, date) {
  let c = t.hdr.indexOf(date);
  if (c >= 2) return c;
  c = t.hdr.length;
  t.hdr.push(date);
  t.rows.forEach(function (r) { r.push(''); });
  t.newCols.push(c);
  return c;
}
function row_(t, id, name) {
  for (let i = 0; i < t.rows.length; i++) if (String(t.rows[i][0]).trim() === id) return i;
  const r = [id, name || ''];
  while (r.length < t.hdr.length) r.push('');
  t.rows.push(r);
  if (t.firstNew < 0) t.firstNew = t.rows.length - 1;
  return t.rows.length - 1;
}
function mark_(t, r, c) {
  const d = t.dirty[c] || (t.dirty[c] = [r, r]);
  if (r < d[0]) d[0] = r;
  if (r > d[1]) d[1] = r;
}
/** بيكتب الأعمدة اللي اتغيّرت بس (مش بيلمس أي أعمدة تانية ضفتها بإيدك) */
function saveTab_(t) {
  const sh = t.sh;
  const nr = t.rows.length + 1, nc = t.hdr.length;
  if (sh.getMaxRows() < nr) sh.insertRowsAfter(sh.getMaxRows(), nr - sh.getMaxRows());
  if (sh.getMaxColumns() < nc) sh.insertColumnsAfter(sh.getMaxColumns(), nc - sh.getMaxColumns());
  t.newCols.forEach(function (c) { sh.getRange(1, c + 1).setNumberFormat('@').setValue(t.hdr[c]); });
  if (t.firstNew >= 0) {
    mark_(t, t.firstNew, 0); mark_(t, t.rows.length - 1, 0);
    mark_(t, t.firstNew, 1); mark_(t, t.rows.length - 1, 1);
  }
  Object.keys(t.dirty).forEach(function (c) {
    const d = t.dirty[c], c0 = Number(c);
    const vals = [];
    for (let r = d[0]; r <= d[1]; r++) vals.push([t.rows[r][c0]]);
    sh.getRange(d[0] + 2, c0 + 1, vals.length, 1).setNumberFormat('@').setValues(vals);
  });
}

/* ---------------- مزامنة الأطفال (اللي اتضافوا في الإكسيل) ---------------- */

function syncKids_(kids) {
  const by = {};
  kids.forEach(function (k) {
    if (!k || !k.id || !k.cls) return;
    (by[k.cls] = by[k.cls] || []).push(k);
  });
  Object.keys(by).forEach(function (cls) {
    const t = loadTab_(cls);
    by[cls].forEach(function (k) {
      const before = t.rows.length;
      const r = row_(t, String(k.id).trim(), String(k.name || '').trim());
      if (r < before && String(t.rows[r][1]).trim() !== String(k.name || '').trim()) {
        t.rows[r][1] = String(k.name || '').trim();
        mark_(t, r, 1);
      }
    });
    saveTab_(t);
  });
}

/* ---------------- تنفيذ التسجيلات ---------------- */

function applyOps_(ops, results) {
  const tabs = {};
  const now = Date.now();
  let notes = null;
  ops.forEach(function (o) {
    try {
      if (o.t === 'att') {
        const date = String(o.date);
        if (!/^\d{4}-\d{2}-\d{2}$/.test(date) || VALID.indexOf(o.v) < 0 || !o.id || !o.cls) throw new Error('bad');
        const parsedDate = new Date(date + 'T00:00:00Z');
        if (isNaN(parsedDate.getTime()) || parsedDate.toISOString().slice(0, 10) !== date) throw new Error('bad');
        const t = tabs[o.cls] || (tabs[o.cls] = loadTab_(o.cls));
        const c = col_(t, o.date);
        const r = row_(t, String(o.id).trim(), String(o.name || '').trim());
        t.rows[r][c] = o.v;
        mark_(t, r, c);
        results.push({ ok: true });
      } else if (o.t === 'note') {
        if (!o.nid || !o.id || !String(o.text || '').trim()) throw new Error('bad');
        if (!notes) notes = loadNotes_();
        if (!notes.ids[o.nid]) {
          notes.sh.appendRow([o.nid, o.id, o.cls || '', o.name || '', /^\d{4}-\d{2}-\d{2}$/.test(String(o.date)) ? o.date : today_(),
            String(o.text), Utilities.formatDate(new Date(Number(o.ts) || now), TZ, 'yyyy-MM-dd HH:mm'), Number(o.ts) || now]);
          notes.ids[o.nid] = true;
        }
        results.push({ ok: true });
      } else if (o.t === 'del') {
        if (!o.nid) throw new Error('bad');
        if (!notes) notes = loadNotes_();
        const last = notes.sh.getLastRow();
        if (last > 1) {
          const ids = notes.sh.getRange(1, 1, last, 1).getValues();
          for (let i = ids.length - 1; i >= 1; i--) {
            if (String(ids[i][0]).trim() === o.nid) { notes.sh.deleteRow(i + 1); break; }
          }
        }
        delete notes.ids[o.nid];
        results.push({ ok: true });
      } else {
        throw new Error('bad');
      }
    } catch (err) {
      results.push({ ok: false, err: String(err && err.message || err) });
    }
  });
  Object.keys(tabs).forEach(function (cls) { saveTab_(tabs[cls]); });
}

/* ---------------- تاب الملاحظات ---------------- */

function notesSheet_() {
  const ss = ss_();
  let sh = ss.getSheetByName(NOTES);
  if (sh) return sh;
  sh = ss.insertSheet(NOTES);
  sh.setRightToLeft(true);
  sh.getRange('A:E').setNumberFormat('@');
  sh.getRange(1, 1, 1, NOTE_HDR.length).setValues([NOTE_HDR]).setFontWeight('bold');
  sh.setFrozenRows(1);
  sh.setColumnWidth(4, 200);
  sh.setColumnWidth(6, 420);
  sh.getRange('F:F').setWrap(true);
  return sh;
}
function loadNotes_() {
  const sh = notesSheet_();
  const ids = {};
  const last = sh.getLastRow();
  if (last > 1) sh.getRange(2, 1, last - 1, 1).getValues().forEach(function (r) { if (r[0]) ids[String(r[0]).trim()] = true; });
  return { sh: sh, ids: ids };
}
