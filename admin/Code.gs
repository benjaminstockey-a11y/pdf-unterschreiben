// PDF Unterschreiben – Anmeldung, Freischaltcode und Nutzungsübersicht.
// Läuft als Apps Script in der Google-Tabelle „PDF Unterschreiben – Nutzer“ und ist als Web-App
// veröffentlicht (Ausführen als: Ich, Zugriff: Jeder). In KEY steht derselbe Schlüssel wie im
// GitHub-Secret REPORT_KEY – im Repository steht hier nur ein Platzhalter.
const KEY = 'HIER_SCHLUESSEL_EINTRAGEN';
const DOMAIN = 'schlueter-gabelstapler.de';
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
  if (!/^[a-z0-9.\-]+@/.test(mail) || !mail.endsWith('@' + DOMAIN)) return { ok: false, error: 'Nur E-Mail-Adressen von ' + DOMAIN + ' sind erlaubt.' };

  const now = new Date();
  log.appendRow([now, name, mail, ev, ver, n]);

  let row = findRow(users, id);
  if (row < 0) { users.appendRow([name, mail, id, ver, now, now, 'nein', '', 0, 0, 0, 0]); row = users.getLastRow(); }
  users.getRange(row, 1, 1, 4).setValues([[name, mail, id, ver]]);
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
        'Falls du dich nicht angemeldet hast, kannst du diese E-Mail ignorieren.\n\nViele Grüße\nBenjamin Stockey',
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
