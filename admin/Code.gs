// PDF Unterschreiben – Anmeldung, Freischaltcode und Nutzungsübersicht.
// Läuft als Apps Script in der Google-Tabelle „PDF Unterschreiben – Nutzer“ und ist als Web-App
// veröffentlicht (Ausführen als: Ich, Zugriff: Jeder). In KEY steht derselbe Schlüssel wie im
// GitHub-Secret REPORT_KEY – im Repository steht hier nur ein Platzhalter.
// Beim Einspielen wird admin/Typenschild.js unten an diese Datei angehängt (ohne die letzte module.exports-Zeile).
const KEY = 'HIER_SCHLUESSEL_EINTRAGEN';
const DOMAIN = 'schlueter-gabelstapler.de';
const ADMIN_MAILS = ['b.stockey@hotmail.de']; // Admin darf sich mit dieser Adresse anmelden
const CODE_MINUTEN = 30;
// Gerätelisten (Excel) melden sich mit diesem eigenen Schlüssel. Er erlaubt NUR die Nutzungsmeldung
// der Listen – keine Anmeldung, keine Admin-Funktionen. Muss mit MELDE_KEY in den Listen übereinstimmen.
const LISTEN_KEY = 'geraeteliste-v1';
// Die Excel-Vorlage „Allgemeine Vorlagen“ darf damit NUR Typenschilder lesen lassen.
const VORLAGEN_KEY = 'allgemeine-vorlagen-v1';

function doPost(e) {
  let d;
  try { d = JSON.parse(e.postData.contents); } catch (x) { return out({ ok: false, error: 'Ungültige Anfrage.' }); }
  if (d && d.key === LISTEN_KEY && /^liste_/.test(String(d.event || ''))) {
    const l = LockService.getScriptLock();
    l.waitLock(20000);
    try { return out(listeMelden(d)); }
    catch (x) { return out({ ok: false, error: 'Serverfehler: ' + x.message }); }
    finally { l.releaseLock(); }
  }
  // Typenschild direkt aus der Excel-Vorlage: Antwort als Text „OK“ + Ergebnis bzw. „FEHLER“ + Meldung
  if (d && d.key === VORLAGEN_KEY && d.event === 'typenschild') {
    let t;
    try {
      const b64 = String(d.b64 || '').replace(/\s/g, '');
      if (b64.length < 100 || b64.length > 15000000) throw new Error('Kein gültiges Foto.');
      const mime = /^image\/(jpeg|png|gif|bmp|webp)$/.test(String(d.mime || '')) ? d.mime : 'image/jpeg';
      t = 'OK\n' + typenschildAuswerten(googleOcr(b64, mime));
    } catch (x) { t = 'FEHLER\n' + x.message; }
    return ContentService.createTextOutput(t).setMimeType(ContentService.MimeType.TEXT);
  }
  // Texterkennung (Typenschilder) – ohne Sperre, damit andere Anfragen nicht warten müssen
  if (d && d.key === KEY && d.event === 'ocr') {
    try { return out(ocrAnfrage(d)); } catch (x) { return out({ ok: false, error: 'Texterkennung: ' + x.message }); }
  }
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
  if (ev === 'adminlist') return adminList(ss, users, log, id, mail);
  if (ev === 'feedbackdelete') return feedbackDelete(ss, users, id, mail, d.code, d.text);
  if (ev === 'adminblock') return adminBlock(users, id, mail, d.code, d.text);
  if (ev === 'admindelete') return adminDelete(ss, users, id, mail, d.code, d.text);
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
      subject: 'Dein Freischaltcode für PDF Unterschreiben & Dokumentation: ' + code,
      body: 'Hallo ' + name + ',\n\ndein Freischaltcode für PDF Unterschreiben & Dokumentation lautet:\n\n    ' + code +
        '\n\nTrag ihn im Programm ein. Der Code ist ' + CODE_MINUTEN + ' Minuten gültig.\n\n' +
        'Falls du dich nicht angemeldet hast, kannst du diese E-Mail ignorieren.\n\nViele Grüße\nPDF Unterschreiben & Dokumentation',
      name: 'PDF Unterschreiben & Dokumentation'
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
      body: name + ' (' + mail + ', Version ' + ver + ') schreibt:\n\n' + t, name: 'PDF Unterschreiben & Dokumentation' });
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
    subject: 'Antwort auf deinen Verbesserungsvorschlag für PDF Unterschreiben & Dokumentation',
    body: 'Hallo ' + r[1] + ',\n\ndanke für deinen Vorschlag' + (when ? ' vom ' + when : '') + ':\n\n    „' +
      String(r[4]).replace(/\n/g, '\n    ') + '“\n\n' + t + '\n\nViele Grüße\nPDF Unterschreiben & Dokumentation',
    name: 'PDF Unterschreiben & Dokumentation'
  });
  const now = new Date();
  s.getRange(row, 6, 1, 3).setValues([['beantwortet', t, now]]);
  return { ok: true, antwort: t, beantwortet: now.toISOString() };
}

// Admin-App: Vorschlag löschen. Zur Sicherheit muss der Zeitpunkt passen – falls sich die Zeilen
// inzwischen verschoben haben, wird nichts Falsches gelöscht.
function feedbackDelete(ss, users, id, mail, rowNo, zeit) {
  if (!isAdmin(users, id, mail)) return { ok: false, error: 'Nur für den Administrator.' };
  const s = ss.getSheetByName('Vorschläge'), row = Number(rowNo);
  if (!s || !(row >= 2 && row <= s.getLastRow())) return { ok: false, error: 'Vorschlag nicht gefunden.' };
  const v = s.getRange(row, 1).getValue();
  const iso = v instanceof Date ? v.toISOString() : String(v || '');
  if (iso !== String(zeit || '')) return { ok: false, error: 'Die Liste hat sich geändert – bitte aktualisieren und nochmal versuchen.' };
  s.deleteRow(row);
  return { ok: true };
}

// Admin-App: einzelnen Eintrag löschen (Protokoll oder Vorlage-Nutzung).
// code = "Blatt|Zeile", text = Zeitpunkt des Eintrags – passt er nicht mehr, wird nichts gelöscht.
function adminDelete(ss, users, id, mail, code, zeit) {
  if (!isAdmin(users, id, mail)) return { ok: false, error: 'Nur für den Administrator.' };
  const p = String(code || '').split('|'), name = p[0], row = Number(p[1]);
  if (['Protokoll', 'Vorlage-Nutzung'].indexOf(name) < 0) return { ok: false, error: 'Unbekannte Liste.' };
  const s = ss.getSheetByName(name);
  if (!s || !(row >= 2 && row <= s.getLastRow())) return { ok: false, error: 'Eintrag nicht gefunden.' };
  const v = s.getRange(row, 1).getValue();
  const iso = v instanceof Date ? v.toISOString() : String(v || '');
  if (iso !== String(zeit || '')) return { ok: false, error: 'Die Liste hat sich geändert – bitte aktualisieren und nochmal versuchen.' };
  s.deleteRow(row);
  return { ok: true };
}

// Gerätelisten (Excel): eine Zeile je Liste + Windows-Benutzer + Rechner
const LISTEN_HEAD = ['Listen-ID', 'Liste', 'Herkunft', 'Benutzer', 'Name', 'Rechner', 'Datei', 'Version',
  'Erstes Öffnen', 'Zuletzt', 'Geöffnet', 'Synchronisiert'];
const VORLAGE_HEAD = ['Zeit', 'Listen-ID', 'Liste', 'Benutzer', 'Name', 'Rechner'];
function listeMelden(d) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const s = sheet(ss, 'Gerätelisten', LISTEN_HEAD);
  const ev = clean(d.event, 20), lid = clean(d.listId, 64), liste = clean(d.liste, 120), datei = clean(d.datei, 120);
  const user = clean(d.user, 60), name = clean(d.name, 100), pc = clean(d.pc, 60), ver = clean(d.version, 20);
  const vorlage = d.vorlage === true;
  if (!lid) return { ok: false, error: 'Listen-ID fehlt.' };
  if (['liste_start', 'liste_sync', 'liste_vorlage'].indexOf(ev) < 0) return { ok: false, error: 'Unbekanntes Ereignis.' };
  const now = new Date();
  if (ev === 'liste_vorlage') {
    const v = sheet(ss, 'Vorlage-Nutzung', VORLAGE_HEAD);
    const ids = v.getLastRow() < 2 ? [] : v.getRange(2, 2, v.getLastRow() - 1, 1).getValues().map(r => String(r[0]));
    if (ids.indexOf(lid) < 0) v.appendRow([now, lid, liste, user, name, pc]);
    return { ok: true };
  }
  let row = -1;
  const last = s.getLastRow();
  if (last >= 2) {
    const v = s.getRange(2, 1, last - 1, 6).getValues();
    for (let i = 0; i < v.length; i++) if (String(v[i][0]) === lid && String(v[i][3]) === user && String(v[i][5]) === pc) { row = i + 2; break; }
  }
  if (row < 0) { s.appendRow([lid, liste, vorlage ? 'Vorlage' : '', user, name, pc, datei, "'" + ver, now, now, 0, 0]); row = s.getLastRow(); }
  s.getRange(row, 2).setValue(liste);
  if (vorlage) s.getRange(row, 3).setValue('Vorlage');
  s.getRange(row, 5, 1, 6).setValues([[name, pc, datei, "'" + ver, s.getRange(row, 9).getValue() || now, now]]);
  const col = ev === 'liste_sync' ? 12 : 11, c = s.getRange(row, col);
  c.setValue((Number(c.getValue()) || 0) + 1);
  return { ok: true };
}

// Admin-App: Gerätelisten zusammenfassen (je Liste alle Nutzer)
function listenUebersicht(ss) {
  const iso = x => x instanceof Date ? x.toISOString() : String(x || '');
  const s = ss.getSheetByName('Gerätelisten'), v = ss.getSheetByName('Vorlage-Nutzung');
  const rows = s && s.getLastRow() >= 2 ? s.getRange(2, 1, s.getLastRow() - 1, 12).getValues() : [];
  const map = {};
  rows.forEach(r => {
    const id = String(r[0]);
    const l = map[id] || (map[id] = { id: id, liste: '', vorlage: false, last: '', nutzer: [] });
    l.liste = String(r[1]) || l.liste;
    if (String(r[2]) === 'Vorlage') l.vorlage = true;
    const u = { benutzer: String(r[3]), name: String(r[4]), rechner: String(r[5]), datei: String(r[6]),
      version: String(r[7]).replace(/^'/, ''), first: iso(r[8]), last: iso(r[9]), opens: Number(r[10]) || 0, syncs: Number(r[11]) || 0 };
    l.nutzer.push(u);
    if (!l.last || new Date(u.last) > new Date(l.last)) l.last = u.last;
  });
  const vor = v && v.getLastRow() >= 2 ? v.getRange(2, 1, v.getLastRow() - 1, 6).getValues() : [];
  return {
    listen: Object.keys(map).map(k => map[k]),
    vorlage: vor.map((r, i) => ({ row: i + 2, zeit: iso(r[0]), id: String(r[1]), liste: String(r[2]), benutzer: String(r[3]), name: String(r[4]), rechner: String(r[5]) })).reverse()
  };
}

// Admin-App: Nutzerliste und Protokoll lesen
function adminList(ss, users, log, id, mail) {
  if (!isAdmin(users, id, mail)) return { ok: false, error: 'Nur für den Administrator.' };
  const iso = x => x instanceof Date ? x.toISOString() : String(x || '');
  const u = users.getLastRow() < 2 ? [] : users.getRange(2, 1, users.getLastRow() - 1, 12).getValues();
  const nLog = Math.min(150, Math.max(0, log.getLastRow() - 1));
  const first = log.getLastRow() - nLog + 1;
  const l = nLog ? log.getRange(first, 1, nLog, 6).getValues().map((r, i) => r.concat([first + i])).reverse() : [];
  const lu = listenUebersicht(ss);
  return { ok: true, listen: lu.listen, vorlage: lu.vorlage,
    users: u.map(r => ({ name: String(r[0]), email: String(r[1]), id: String(r[2]), version: String(r[3]).replace(/^'/, ''),
      first: iso(r[4]), last: iso(r[5]), verified: String(r[6]).toLowerCase() === 'ja', blocked: String(r[7]).trim().toLowerCase() === 'ja',
      starts: Number(r[8]) || 0 })),
    log: l.map(r => ({ row: r[6], zeit: iso(r[0]), name: String(r[1]), email: String(r[2]), event: String(r[3]), version: String(r[4]).replace(/^'/, '') })) };
}

// Admin-App: Nutzer sperren („ja“) oder entsperren – der eigene Zugang lässt sich nicht sperren
function adminBlock(users, id, mail, targetId, value) {
  if (!isAdmin(users, id, mail)) return { ok: false, error: 'Nur für den Administrator.' };
  const row = findRow(users, String(targetId || ''));
  if (row < 0) return { ok: false, error: 'Nutzer nicht gefunden.' };
  if (ADMIN_MAILS.indexOf(String(users.getRange(row, 2).getValue()).toLowerCase()) >= 0) return { ok: false, error: 'Der Administrator kann nicht gesperrt werden.' };
  const block = String(value) === 'ja';
  users.getRange(row, 8).setValue(block ? 'ja' : '');
  return { ok: true, blocked: block };
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

// ---------------- Texterkennung mit Google ----------------
// Das Foto wird kurz als Google-Dokument mit Texterkennung in Drive abgelegt, der Text gelesen und
// die Datei sofort wieder endgültig gelöscht. (DriveApp wird nur für die Drive-Berechtigung genannt.)
function ocrAnfrage(d) {
  const mail = clean(d.email, 150).toLowerCase();
  if (ADMIN_MAILS.indexOf(mail) < 0 && !mail.endsWith('@' + DOMAIN)) return { ok: false, error: 'Nicht berechtigt.' };
  const b64 = String(d.b64 || '');
  if (b64.length < 100 || b64.length > 15000000) return { ok: false, error: 'Kein gültiges Foto.' };
  const mime = /^image\/(jpeg|png|gif|bmp|webp)$/.test(String(d.mime || '')) ? d.mime : 'image/jpeg';
  return { ok: true, text: googleOcr(b64, mime) };
}

function googleOcr(b64, mime) {
  // Braucht vollen Drive-Zugriff (Datei anlegen/löschen), nicht nur Lesen: DriveApp.createFile
  DriveApp.getRootFolder();
  const tok = ScriptApp.getOAuthToken();
  const auth = { Authorization: 'Bearer ' + tok };
  const bd = 'puocr' + Date.now();
  const meta = { name: 'Typenschild-OCR', mimeType: 'application/vnd.google-apps.document' };
  const body = Utilities.newBlob('--' + bd + '\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n' + JSON.stringify(meta) +
    '\r\n--' + bd + '\r\nContent-Type: ' + (mime || 'image/jpeg') + '\r\nContent-Transfer-Encoding: base64\r\n\r\n' + b64 + '\r\n--' + bd + '--').getBytes();
  const r = UrlFetchApp.fetch('https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&ocrLanguage=de',
    { method: 'post', contentType: 'multipart/related; boundary=' + bd, payload: body, headers: auth, muteHttpExceptions: true });
  if (r.getResponseCode() >= 300) throw new Error('Hochladen fehlgeschlagen (' + r.getResponseCode() + ') ' + r.getContentText().slice(0, 400));
  const id = JSON.parse(r.getContentText()).id;
  try {
    return UrlFetchApp.fetch('https://www.googleapis.com/drive/v3/files/' + id + '/export?mimeType=text/plain', { headers: auth }).getContentText();
  } finally {
    UrlFetchApp.fetch('https://www.googleapis.com/drive/v3/files/' + id, { method: 'delete', headers: auth, muteHttpExceptions: true });
  }
}


// Einmal im Editor ausführen, damit Google die Drive-Berechtigung für die Texterkennung abfragt.
function berechtigungenPruefen() {
  DriveApp.getRootFolder();
  UrlFetchApp.fetch('https://www.googleapis.com/drive/v3/about?fields=user', { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() } });
}
