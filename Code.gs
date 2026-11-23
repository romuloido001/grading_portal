// Grading Portal backend — paste into Extensions > Apps Script of a Google Sheet.
// Deploy > New deployment > Web app > Execute as: Me, Access: Anyone.
const ADMIN_KEY = '123';   // can add/edit AND delete learners
const TEACHER_KEY = '123'; // can add/edit grades only
const role = k => k === ADMIN_KEY ? 'admin' : k === TEACHER_KEY ? 'teacher' : null;
const LH = ['LRN','Name','Age','Sex','Section','Grade','SchoolYear','Code','Adviser','Principal','C1','C2','C3','Days'];
const GH = ['LRN','Subject','T1','T2','T3'];

function sh(n, h) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let s = ss.getSheetByName(n);
  if (!s) { s = ss.insertSheet(n); s.appendRow(h); s.getRange('A:A').setNumberFormat('@'); }
  return s;
}
function rows(s) {
  const v = s.getDataRange().getValues(), h = v.shift();
  return v.map(r => Object.fromEntries(h.map((k, i) => [k, r[i]])));
}
function out(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
function load(lrn) {
  const l = rows(sh('Learners', LH)).find(r => String(r.LRN) === String(lrn));
  if (!l) return null;
  const g = rows(sh('Grades', GH)).filter(r => String(r.LRN) === String(lrn));
  return { learner: l, grades: g };
}

function doGet(e) {
  const p = e.parameter;
  if (p.action === 'report') {
    const d = load(p.lrn);
    if (!d || String(d.learner.Code) !== String(p.code)) return out({ ok: false, error: 'LRN or access code is incorrect.' });
    delete d.learner.Code;
    return out({ ok: true, ...d });
  }
  if (!role(p.key)) return out({ ok: false, error: 'Invalid teacher key.' });
  if (p.action === 'list') {
    return out({ ok: true, learners: rows(sh('Learners', LH)).map(r => ({ LRN: String(r.LRN), Name: r.Name, Section: r.Section })) });
  }
  if (p.action === 'load') {
    const d = load(p.lrn);
    return d ? out({ ok: true, ...d }) : out({ ok: false, error: 'Not found.' });
  }
  return out({ ok: false, error: 'Unknown action.' });
}

function doPost(e) {
  const b = JSON.parse(e.postData.contents);
  if (!role(b.key)) return out({ ok: false, error: 'Invalid teacher key.' });
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const ls = sh('Learners', LH), gs = sh('Grades', GH);
    if (b.action === 'save') {
      const l = b.learner, lrn = String(l.LRN);
      const row = LH.map(k => (k === 'LRN' ? lrn : l[k] === undefined ? '' : l[k]));
      const ids = ls.getRange(1, 1, ls.getLastRow(), 1).getValues().map(r => String(r[0]));
      const i = ids.indexOf(lrn);
      if (i > 0) ls.getRange(i + 1, 1, 1, LH.length).setValues([row]); else ls.appendRow(row);
      for (let r = gs.getLastRow(); r > 1; r--) if (String(gs.getRange(r, 1).getValue()) === lrn) gs.deleteRow(r);
      b.grades.forEach(g => gs.appendRow([lrn, g.Subject, g.T1, g.T2, g.T3]));
      return out({ ok: true });
    }
    if (b.action === 'delete') {
      if (role(b.key) !== 'admin') return out({ ok: false, error: 'Only the admin key can delete.' });
      const ids = ls.getRange(1, 1, ls.getLastRow(), 1).getValues().map(r => String(r[0]));
      const i = ids.indexOf(String(b.lrn));
      if (i > 0) ls.deleteRow(i + 1);
      for (let r = gs.getLastRow(); r > 1; r--) if (String(gs.getRange(r, 1).getValue()) === String(b.lrn)) gs.deleteRow(r);
      return out({ ok: true });
    }
    return out({ ok: false, error: 'Unknown action.' });
  } finally { lock.releaseLock(); }
}
