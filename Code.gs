// Grading Portal backend — paste into Extensions > Apps Script of a Google Sheet.
// Deploy > Manage deployments > Edit > New version. Access: Anyone.
// Accounts live in the "Users" sheet (created automatically with 3 sample logins, password 123):
//   Username | Password | Role (admin / adviser / teacher) | Name | Section (adviser) | Subjects (teacher, comma-separated)
const LH = ['LRN','Name','Age','Sex','Section','Grade','SchoolYear','Code','Adviser','Principal','C1','C2','C3','Days'];
const GH = ['LRN','Subject','T1','T2','T3'];
const AH = ['Section','Subject','Teacher'];
const SUBJECTS = ['Filipino','English','Mathematics','Science','Araling Panlipunan','Values Education','Technology and Livelihood Education (TLE)','Information and Communication Technology (ICT)','Music and Arts','Physical Education and Health'];
const UH = ['Username','Password','Role','Name','Section','Subjects','Status'];
const REG_CODE = '123'; // teachers must enter this code to request an account (admin still has to approve)
const hash = p => Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, 'cnhs:' + p));

function sh(n, h) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let s = ss.getSheetByName(n);
  if (!s) { s = ss.insertSheet(n); s.appendRow(h); s.getRange('A:B').setNumberFormat('@'); }
  return s;
}
function rows(s) {
  const v = s.getDataRange().getValues(), h = v.shift();
  return v.map(r => Object.fromEntries(h.map((k, i) => [k, r[i]])));
}
function out(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
const fail = m => out({ ok: false, error: m });

function users() {
  const s = sh('Users', UH);
  if (s.getRange(1, 7).getValue() !== 'Status') s.getRange(1, 7).setValue('Status');
  if (s.getLastRow() < 2) {
    s.getRange('A:B').setNumberFormat('@');
    [['admin','123','admin','Administrator','','','active'],
     ['adviser','123','adviser','Sample Adviser','ATIS','','active'],
     ['teacher','123','teacher','Sample Subject Teacher','','Science','active']].forEach(r => s.appendRow(r));
  }
  return rows(s);
}
function auth(u, p) {
  return users().find(r => String(r.Username).toLowerCase() === String(u).toLowerCase() &&
    (String(r.Password) === String(p) || String(r.Password) === hash(p))) || null;
}
const pending = u => String(u.Status) === 'pending';
const pub = x => ({ Username: x.Username, Role: x.Role, Name: x.Name, Section: x.Section, Subjects: x.Subjects });
const same = (a, b) => String(a).trim().toLowerCase() === String(b).trim().toLowerCase();
let _A = null;
const assigns = () => _A || (_A = rows(sh('Assignments', AH)));
const mine = u => assigns().filter(a => same(a.Teacher, u.Username));
const canSee = (u, l) => u.Role === 'admin' ||
  (u.Role === 'adviser' && same(l.Section, u.Section)) ||
  (u.Role === 'teacher' && mine(u).some(a => same(a.Section, l.Section)));
// teachers may grade only the subjects an adviser assigned to them for that learner's section
const canGrade = (u, subj, section) => u.Role === 'admin' ||
  (u.Role === 'teacher' && mine(u).some(a => same(a.Section, section) && a.Subject === subj));

function load(lrn) {
  const l = rows(sh('Learners', LH)).find(r => String(r.LRN) === String(lrn));
  if (!l) return null;
  return { learner: l, grades: rows(sh('Grades', GH)).filter(r => String(r.LRN) === String(lrn)) };
}
function findRow(s, lrn) {
  const ids = s.getRange(1, 1, s.getLastRow(), 1).getValues().map(r => String(r[0]));
  const i = ids.indexOf(String(lrn));
  return i > 0 ? i + 1 : 0;
}

function doGet(e) {
  const p = e.parameter;
  if (!p.action) return out({ ok: true, version: 'v7-assign-teachers' }); // open the web app URL in a browser to check which code is live
  if (p.action === 'report') {            // parents: LRN + access code
    const d = load(p.lrn);
    if (!d || String(d.learner.Code) !== String(p.code)) return fail('LRN or access code is incorrect.');
    delete d.learner.Code;
    return out({ ok: true, ...d });
  }
  const u = auth(p.u, p.p);
  if (!u) return fail('Invalid username or password.');
  if (pending(u)) return fail('Your account is waiting for admin approval.');
  if (p.action === 'pending') {
    if (u.Role !== 'admin') return fail('Admin only.');
    return out({ ok: true, users: users().filter(pending).map(pub) });
  }
  if (p.action === 'assignments') {       // adviser (own section) or admin
    const sec = u.Role === 'adviser' ? u.Section : u.Role === 'admin' ? p.section : null;
    if (!sec) return fail('Not allowed.');
    const us = users();
    return out({ ok: true, section: sec,
      assignments: assigns().filter(a => same(a.Section, sec)).map(a => ({ Subject: a.Subject, Teacher: a.Teacher, Name: (us.find(x => same(x.Username, a.Teacher)) || {}).Name || a.Teacher })),
      teachers: us.filter(x => x.Role === 'teacher' && String(x.Status) !== 'pending').map(x => ({ Username: x.Username, Name: x.Name, Subjects: x.Subjects })) });
  }
  if (p.action === 'users' || p.action === 'learners') {   // admin page
    if (u.Role !== 'admin') return fail('Admin only.');
    if (p.action === 'users') return out({ ok: true, users: users().map(x => ({ ...pub(x), Status: x.Status || 'active' })) });
    return out({ ok: true, learners: rows(sh('Learners', LH)).map(l => ({ LRN: String(l.LRN), Name: l.Name, Sex: l.Sex, Section: l.Section, Code: l.Code, Adviser: l.Adviser })) });
  }
  if (p.action === 'list') {
    const list = rows(sh('Learners', LH)).filter(l => canSee(u, l))
      .map(l => ({ LRN: String(l.LRN), Name: l.Name, Section: l.Section }));
    return out({ ok: true, user: pub(u), learners: list, mine: u.Role === 'teacher' ? mine(u).map(a => ({ Section: a.Section, Subject: a.Subject })) : [] });
  }
  if (p.action === 'load') {
    const d = load(p.lrn);
    if (!d || !canSee(u, d.learner)) return fail('Not found.');
    if (u.Role === 'teacher') delete d.learner.Code;
    return out({ ok: true, ...d });
  }
  return fail('Unknown action.');
}

function doPost(e) {
  const b = JSON.parse(e.postData.contents);
  if (b.action === 'register') return register(b);
  const u = auth(b.u, b.p);
  if (!u) return fail('Invalid username or password.');
  if (pending(u)) return fail('Your account is waiting for admin approval.');
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const ls = sh('Learners', LH), gs = sh('Grades', GH);
    if (b.action === 'save') {            // admin: learner profile + all grades
      if (u.Role !== 'admin') return fail('Admin only.');
      const l = b.learner, lrn = String(l.LRN);
      const row = LH.map(k => (k === 'LRN' ? lrn : l[k] === undefined ? '' : l[k]));
      const r = findRow(ls, lrn);
      if (r) ls.getRange(r, 1, 1, LH.length).setValues([row]); else ls.appendRow(row);
      saveGrades(gs, lrn, b.grades, u, l.Section);
      return out({ ok: true });
    }
    if (b.action === 'saveGrades') {      // subject teacher (own subjects) or admin
      const d = load(b.lrn);
      if (!d || !canSee(u, d.learner)) return fail('Not found.');
      const n = saveGrades(gs, b.lrn, b.grades, u, d.learner.Section);
      return out({ ok: true, saved: n });
    }
    if (b.action === 'saveInfo') {        // adviser (own section) or admin: attendance + comments
      const d = load(b.lrn);
      if (!d || !(u.Role === 'admin' || u.Role === 'adviser') || !canSee(u, d.learner)) return fail('Not allowed.');
      const r = findRow(ls, b.lrn);
      ['C1','C2','C3','Days'].forEach(k => ls.getRange(r, LH.indexOf(k) + 1).setValue(b[k] === undefined ? '' : b[k]));
      return out({ ok: true });
    }
    if (b.action === 'userUpdate' || b.action === 'userDelete') {
      if (u.Role !== 'admin') return fail('Admin only.');
      const us = sh('Users', UH), v = us.getDataRange().getValues();
      const i = v.findIndex((r, k) => k > 0 && String(r[0]).toLowerCase() === String(b.username).toLowerCase());
      if (i < 1) return fail('Account not found.');
      const self = String(v[i][0]).toLowerCase() === String(u.Username).toLowerCase();
      if (b.action === 'userDelete') {
        if (self) return fail("You can't delete your own account.");
        us.deleteRow(i + 1); return out({ ok: true });
      }
      if (b.role) {
        if (!['admin', 'adviser', 'teacher'].includes(b.role)) return fail('Invalid role.');
        if (self && b.role !== 'admin') return fail("You can't change your own role.");
        us.getRange(i + 1, 3).setValue(b.role);
      }
      if (b.name !== undefined) us.getRange(i + 1, 4).setValue(b.name);
      if (b.section !== undefined) us.getRange(i + 1, 5).setValue(b.section);
      if (b.subjects !== undefined) us.getRange(i + 1, 6).setValue(b.subjects);
      if (b.status) us.getRange(i + 1, 7).setValue(b.status);
      if (b.newPassword) {
        if (String(b.newPassword).length < 6) return fail('Password must be at least 6 characters.');
        us.getRange(i + 1, 2).setValue(hash(b.newPassword));
      }
      return out({ ok: true });
    }
    if (b.action === 'importLearners') {  // admin: add many learners at once
      if (u.Role !== 'admin') return fail('Admin only.');
      const have = new Set(ls.getRange(1, 1, ls.getLastRow(), 1).getValues().map(r => String(r[0])));
      const d = b.defaults || {};
      let added = 0, skipped = [];
      (b.rows || []).forEach(r => {
        const lrn = String(r.LRN || '').trim();
        if (!/^\d{12}$/.test(lrn) || !String(r.Name || '').trim()) { skipped.push(lrn || '(blank)'); return; }
        if (have.has(lrn)) { skipped.push(lrn + ' (exists)'); return; }
        const rec = { LRN: lrn, Name: String(r.Name).trim(), Age: r.Age || '', Sex: r.Sex || '', Section: String(r.Section || '').trim(),
          Grade: d.Grade || 7, SchoolYear: d.SchoolYear || '2026–2027', Code: String(Math.floor(100000 + Math.random() * 900000)),
          Adviser: d.Adviser || '', Principal: d.Principal || '', C1: '', C2: '', C3: '', Days: '' };
        ls.appendRow(LH.map(k => rec[k])); have.add(lrn); added++;
      });
      return out({ ok: true, added, skipped });
    }
    if (b.action === 'assign' || b.action === 'unassign') {   // adviser assigns subject teachers to own section
      const sec = u.Role === 'adviser' ? u.Section : u.Role === 'admin' ? b.section : null;
      if (!sec) return fail('Only the adviser of this section or the admin can assign teachers.');
      if (!SUBJECTS.includes(b.subject)) return fail('Unknown subject.');
      const as = sh('Assignments', AH), v = as.getDataRange().getValues();
      const i = v.findIndex((r, k) => k > 0 && same(r[0], sec) && r[1] === b.subject);
      if (b.action === 'unassign') { if (i > 0) as.deleteRow(i + 1); return out({ ok: true }); }
      const t = users().find(x => same(x.Username, b.teacher) && x.Role === 'teacher' && String(x.Status) !== 'pending');
      if (!t) return fail('Teacher not found or not approved yet.');
      const row = [String(sec).trim(), b.subject, t.Username];
      if (i > 0) as.getRange(i + 1, 1, 1, 3).setValues([row]); else as.appendRow(row);
      return out({ ok: true });
    }
    if (b.action === 'decide') {          // admin approves / rejects a new account
      if (u.Role !== 'admin') return fail('Admin only.');
      const us = sh('Users', UH), v = us.getDataRange().getValues();
      const i = v.findIndex((r, k) => k > 0 && String(r[0]).toLowerCase() === String(b.username).toLowerCase());
      if (i < 1) return fail('Account not found.');
      if (b.approve) us.getRange(i + 1, 7).setValue('active'); else us.deleteRow(i + 1);
      return out({ ok: true });
    }
    if (b.action === 'delete') {
      if (u.Role !== 'admin') return fail('Only the admin can delete.');
      const r = findRow(ls, b.lrn);
      if (r) ls.deleteRow(r);
      for (let i = gs.getLastRow(); i > 1; i--) if (String(gs.getRange(i, 1).getValue()) === String(b.lrn)) gs.deleteRow(i);
      return out({ ok: true });
    }
    return fail('Unknown action.');
  } finally { lock.releaseLock(); }
}

function saveGrades(gs, lrn, grades, u, section) {
  const v = gs.getDataRange().getValues();
  let n = 0;
  (grades || []).forEach(g => {
    if (!canGrade(u, g.Subject, section)) return;  // server-side permission check
    const i = v.findIndex((r, k) => k > 0 && String(r[0]) === String(lrn) && r[1] === g.Subject);
    const row = [String(lrn), g.Subject, g.T1, g.T2, g.T3];
    if (i > 0) gs.getRange(i + 1, 1, 1, 5).setValues([row]); else gs.appendRow(row);
    n++;
  });
  return n;
}

function register(b) {
  if (String(b.regCode) !== REG_CODE) return fail('Invalid registration code.');
  const name = String(b.name || '').trim(), un = String(b.username || '').trim().toLowerCase(), role = b.role;
  if (!name) return fail('Enter your full name.');
  if (!/^[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}$/.test(un) || un.length > 80) return fail('Username must be a valid email address (e.g. name@gmail.com).');
  if (String(b.password || '').length < 6) return fail('Password must be at least 6 characters.');
  if (role !== 'teacher' && role !== 'adviser') return fail('Choose a role.');
  if (role === 'adviser' && !String(b.section || '').trim()) return fail('Enter your advisory section.');
  if (role === 'teacher' && !String(b.subjects || '').trim()) return fail('Enter the subject(s) you teach.');
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    if (users().some(r => String(r.Username).toLowerCase() === un.toLowerCase())) return fail('Username is already taken.');
    sh('Users', UH).appendRow([un, hash(b.password), role, name, role === 'adviser' ? String(b.section).trim() : '', role === 'teacher' ? String(b.subjects).trim() : '', 'pending']);
    return out({ ok: true });
  } finally { lock.releaseLock(); }
}
