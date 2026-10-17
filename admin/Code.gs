// PDF Unterschreiben – Anmeldung, Freischaltcode und Nutzungsübersicht.
// Läuft als Apps Script in der Google-Tabelle „PDF Unterschreiben – Nutzer“ und ist als Web-App
// veröffentlicht (Ausführen als: Ich, Zugriff: Jeder). In KEY steht derselbe Schlüssel wie im
// GitHub-Secret REPORT_KEY – im Repository steht hier nur ein Platzhalter.
const KEY = 'HIER_SCHLUESSEL_EINTRAGEN';
const DOMAIN = 'schlueter-gabelstapler.de';
const ADMIN_MAILS = ['b.stockey@hotmail.de']; // Admin darf sich mit dieser Adresse anmelden
const CODE_MINUTEN = 30;

function doPost(e) {
  let d;
  try { d = JSON.parse(e.postData.contents); } catch (x) { return out({ ok: false, error: 'Ungültige Anfrage.' }); }
  if (!d || d.key !== KEY) return out({ ok: false, error: 'Nicht berechtigt.' });
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { return out(handle(d)); }
  catch (x) { return out({ ok: false, error: 'Serverfehler: ' + x.message }); }
  finally { lock.releaseLock(); }
}

function handle(d) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const users = sheet(ss, 'Nutzer', ['Name', 'E-Mail', 'Rechner-ID', 'Version', 'Erste Anmeldung', 'Zuletzt aktiv',
    'Freigeschaltet', 'Gesperrt (ja)', 'Starts', 'Gespeichert', 'Gedruckt', 'Per E-Mail']);
  const log = sheet(ss, 'Protokoll', ['Zeit', 'Name', 'E-Mail', 'Ereignis', 'Version', 'Anzahl']);

  const id = clean(d.id, 64), name = clean(d.name, 100), mail = clean(d.email, 150).toLowerCase();
  const ev = clean(d.event, 20), ver = clean(d.version, 20);
  const n = Math.max(1, Math.min(1000, Number(d.count) || 1));
  if (!id) return { ok: false, error: 'Rechner-ID fehlt.' };
  if (ADMIN_MAILS.indexOf(mail) < 0 && (!/^[a-z0-9.\-]+@/.test(mail) || !mail.endsWith('@' + DOMAIN))) return { ok: false, error: 'Nur E-Mail-Adressen von ' + DOMAIN + ' sind erlaubt.' };

  const now = new Date();
  if (ev === 'feedbacklist') return feedbackList(ss, users, id, mail);
  if (ev === 'feedback') return feedback(ss, name, mail, ver, d.text);
  if (ev === 'feedbackreply') return feedbackReply(ss, users, id, mail, d.code, d.text);
  if (ev !== 'check') log.appendRow([now, name, mail, ev, "'" + ver, n]);

  let row = findRow(users, id);
  if (row < 0) { users.appendRow([name, mail, id, "'" + ver, now, now, 'nein', '', 0, 0, 0, 0]); row = users.getLastRow(); }
  users.getRange(row, 1, 1, 4).setValues([[name, mail, id, "'" + ver]]);
  users.getRange(row, 6).setValue(now);
  const r = users.getRange(row, 1, 1, 12).getValues()[0];
  const blocked = String(r[7]).trim().toLowerCase() === 'ja';
  const props = PropertiesService.getScriptProperties();

  if (ev === 'register' || ev === 'update' || ev === 'resend') {
    if (ev === 'update' && String(r[6]).toLowerCase() === 'ja' && r[1] === mail) return { ok: true, verified: true, blocked: blocked };
    const code = String(Math.floor(100000 + Math.random() * 900000));
    props.setProperty('code_' + id, JSON.stringify({ code: code, mail: mail, exp: Date.now() + CODE_MINUTEN * 60000, tries: 0 }));
    users.getRange(row, 7).setValue('nein');
    MailApp.sendEmail({
      to: mail,
      subject: 'Dein Freischaltcode für PDF Unterschreiben: ' + code,
      body: 'Hallo ' + name + ',\n\ndein Freischaltcode für PDF Unterschreiben lautet:\n\n    ' + code +
        '\n\nTrag ihn im Programm ein. Der Code ist ' + CODE_MINUTEN + ' Minuten gültig.\n\n' +
        'Falls du dich nicht angemeldet hast, kannst du diese E-Mail ignorieren.\n\nViele Grüße\nPDF Unterschreiben',
      name: 'PDF Unterschreiben'
    });
    return { ok: true, sent: true };
  }

  if (ev === 'verify') {
    const raw = props.getProperty('code_' + id);
    if (!raw) return { ok: false, error: 'Bitte zuerst einen Code anfordern.' };
    const c = JSON.parse(raw);
    if (Date.now() > c.exp) { props.deleteProperty('code_' + id); return { ok: false, error: 'Der Code ist abgelaufen. Bitte „Code neu senden“.' }; }
    if (c.mail !== mail) return { ok: false, error: 'Der Code gehört zu einer anderen E-Mail-Adresse.' };
    if (String(d.code || '').trim() !== c.code) {
      c.tries++;
      if (c.tries >= 5) { props.deleteProperty('code_' + id); return { ok: false, error: 'Zu viele Fehlversuche. Bitte „Code neu senden“.' }; }
      props.setProperty('code_' + id, JSON.stringify(c));
      return { ok: false, error: 'Der Code stimmt nicht.' };
    }
    props.deleteProperty('code_' + id);
    users.getRange(row, 7).setValue('ja');
    return { ok: true, verified: true, blocked: blocked };
  }

  const col = { start: 9, save: 10, print: 11, mail: 12 }[ev];
  if (col) { const cell = users.getRange(row, col); cell.setValue((Number(cell.getValue()) || 0) + n); }
  return { ok: true, verified: String(r[6]).toLowerCase() === 'ja', blocked: blocked };
}

// Verbesserungsvorschläge: landen im Blatt „Vorschläge“; der Admin bekommt zusätzlich eine E-Mail.
function feedback(ss, name, mail, ver, text) {
  const t = clean(text, 3000);
  if (t.length < 3) return { ok: false, error: 'Bitte einen Vorschlag eintragen.' };
  const s = ideasSheet(ss);
  s.appendRow([new Date(), name, mail, "'" + ver, t, 'neu']);
  try {
    MailApp.sendEmail({ to: ADMIN_MAILS[0], subject: 'Neuer Verbesserungsvorschlag von ' + name,
      body: name + ' (' + mail + ', Version ' + ver + ') schreibt:\n\n' + t, name: 'PDF Unterschreiben' });
  } catch (x) {}
  return { ok: true };
}

const IDEA_HEAD = ['Zeit', 'Name', 'E-Mail', 'Version', 'Vorschlag', 'Status / Notiz', 'Antwort', 'Beantwortet am'];
function ideasSheet(ss) {
  const s = sheet(ss, 'Vorschläge', IDEA_HEAD);
  if (s.getLastColumn() < IDEA_HEAD.length) s.getRange(1, 1, 1, IDEA_HEAD.length).setValues([IDEA_HEAD]).setFontWeight('bold');
  return s;
}

// Nur der freigeschaltete Admin darf die Vorschläge lesen und beantworten
function isAdmin(users, id, mail) {
  const row = findRow(users, id);
  return row > 0 && ADMIN_MAILS.indexOf(mail) >= 0 &&
    String(users.getRange(row, 7).getValue()).toLowerCase() === 'ja' && users.getRange(row, 2).getValue() === mail;
}

function feedbackList(ss, users, id, mail) {
  if (!isAdmin(users, id, mail)) return { ok: false, error: 'Nur für den Administrator.' };
  const s = ss.getSheetByName('Vorschläge');
  if (!s || s.getLastRow() < 2) return { ok: true, items: [] };
  const last = s.getLastRow(), n = Math.min(100, last - 1), first = last - n + 1;
  const v = s.getRange(first, 1, n, 8).getValues();
  const iso = x => x instanceof Date ? x.toISOString() : String(x || '');
  return { ok: true, items: v.map((r, i) => ({ row: first + i, zeit: iso(r[0]),
    name: String(r[1]), email: String(r[2]), version: String(r[3]), text: String(r[4]), status: String(r[5]),
    antwort: String(r[6] || ''), beantwortet: iso(r[7]) })).reverse() };
}

// Antwort des Admins: geht per E-Mail an die Person und wird in der Tabelle vermerkt
function feedbackReply(ss, users, id, mail, rowNo, text) {
  if (!isAdmin(users, id, mail)) return { ok: false, error: 'Nur für den Administrator.' };
  const t = clean(text, 3000), row = Number(rowNo);
  if (t.length < 2) return { ok: false, error: 'Bitte eine Antwort eintragen.' };
  const s = ideasSheet(ss);
  if (!(row >= 2 && row <= s.getLastRow())) return { ok: false, error: 'Vorschlag nicht gefunden.' };
  const r = s.getRange(row, 1, 1, 8).getValues()[0];
  const to = String(r[2]).trim();
  if (!/^[^@\s]+@[^@\s]+$/.test(to)) return { ok: false, error: 'Zu diesem Vorschlag gibt es keine E-Mail-Adresse.' };
  const when = r[0] instanceof Date ? Utilities.formatDate(r[0], 'Europe/Berlin', 'dd.MM.yyyy') : '';
  MailApp.sendEmail({
    to: to,
    subject: 'Antwort auf deinen Verbesserungsvorschlag für PDF Unterschreiben',
    body: 'Hallo ' + r[1] + ',\n\ndanke für deinen Vorschlag' + (when ? ' vom ' + when : '') + ':\n\n    „' +
      String(r[4]).replace(/\n/g, '\n    ') + '“\n\n' + t + '\n\nViele Grüße\nPDF Unterschreiben',
    name: 'PDF Unterschreiben'
  });
  const now = new Date();
  s.getRange(row, 6, 1, 3).setValues([['beantwortet', t, now]]);
  return { ok: true, antwort: t, beantwortet: now.toISOString() };
}

function sheet(ss, name, head) {
  let s = ss.getSheetByName(name);
  if (!s) {
    s = ss.insertSheet(name);
    s.appendRow(head);
    s.setFrozenRows(1);
    s.getRange(1, 1, 1, head.length).setFontWeight('bold');
  }
  return s;
}

function findRow(s, id) {
  const last = s.getLastRow();
  if (last < 2) return -1;
  const v = s.getRange(2, 3, last - 1, 1).getValues();
  for (let i = 0; i < v.length; i++) if (v[i][0] === id) return i + 2;
  return -1;
}

function clean(v, max) {
  let s = String(v == null ? '' : v).slice(0, max).trim();
  if (/^[=+\-@]/.test(s)) s = "'" + s; // keine Formeln einschleusen
  return s;
}

function out(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

// Einmal im Editor ausführen, damit Google die Berechtigungen (Tabelle, E-Mail) abfragt.
function einrichten() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  sheet(ss, 'Nutzer', ['Name', 'E-Mail', 'Rechner-ID', 'Version', 'Erste Anmeldung', 'Zuletzt aktiv',
    'Freigeschaltet', 'Gesperrt (ja)', 'Starts', 'Gespeichert', 'Gedruckt', 'Per E-Mail']);
  sheet(ss, 'Protokoll', ['Zeit', 'Name', 'E-Mail', 'Ereignis', 'Version', 'Anzahl']);
  const s1 = ss.getSheetByName('Tabellenblatt1') || ss.getSheetByName('Sheet1');
  if (s1 && ss.getSheets().length > 2) ss.deleteSheet(s1);
  MailApp.getRemainingDailyQuota();
}
