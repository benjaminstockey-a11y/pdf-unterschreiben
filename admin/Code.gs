// PDF Unterschreiben – Anmeldung, Freischaltcode und Nutzungsübersicht.
// Läuft als Apps Script in der Google-Tabelle „PDF Unterschreiben – Nutzer“ und ist als Web-App
// veröffentlicht (Ausführen als: Ich, Zugriff: Jeder). In KEY steht derselbe Schlüssel wie im
// GitHub-Secret REPORT_KEY – im Repository steht hier nur ein Platzhalter.
// Beim Einspielen werden admin/Typenschild.js, admin/QrCode.js und admin/Vorlagen.js unten angehängt (jeweils ohne die letzte module.exports-Zeile).
const KEY = 'HIER_SCHLUESSEL_EINTRAGEN';
const DOMAIN = 'schlueter-gabelstapler.de';
const ADMIN_MAILS = ['b.stockey@hotmail.de']; // Admin darf sich mit dieser Adresse anmelden
const CODE_MINUTEN = 30;
// Gerätelisten (Excel) melden sich mit diesem eigenen Schlüssel. Er erlaubt NUR die Nutzungsmeldung
// der Listen – keine Anmeldung, keine Admin-Funktionen. Muss mit MELDE_KEY in den Listen übereinstimmen.
const LISTEN_KEY = 'geraeteliste-v1';
// Die Excel-Vorlage „Allgemeine Vorlagen“ darf damit NUR Typenschilder lesen lassen.
const VORLAGEN_KEY = 'allgemeine-vorlagen-v1';
// true = die Excel-Vorlage „Allgemeine Vorlagen“ bekommt keine Server-Funktionen mehr (abgelöst durch die Service-Mappe)
const VORLAGEN_AUS = true;
const VORLAGEN_AUS_TEXT = 'Die Allgemeinen Vorlagen werden nicht mehr verwendet. Bitte die Service-Mappe nutzen (Formulare) und dort auf die neueste Version aktualisieren.';

function doPost(e) {
  let d;
  try { d = JSON.parse(e.postData.contents); } catch (x) { return out({ ok: false, error: 'Ungültige Anfrage.' }); }
  if (VORLAGEN_AUS && d && d.key === VORLAGEN_KEY)
    return ContentService.createTextOutput('FEHLER\n' + VORLAGEN_AUS_TEXT).setMimeType(ContentService.MimeType.TEXT);
  if (d && d.key === LISTEN_KEY && /^liste_/.test(String(d.event || ''))) {
    const l = LockService.getScriptLock();
    l.waitLock(20000);
    try { return out(listeMelden(d)); }
    catch (x) { return out({ ok: false, error: 'Serverfehler: ' + x.message }); }
    finally { l.releaseLock(); }
  }
  // Typenschild direkt aus der Excel-Vorlage: Antwort als Text „OK“ + Ergebnis bzw. „FEHLER“ + Meldung
  if (d && (d.key === VORLAGEN_KEY || d.key === KEY) && d.event === 'typenschild') {
    let t;
    try {
      const b64 = String(d.b64 || '').replace(/\s/g, '');
      if (b64.length < 100 || b64.length > 15000000) throw new Error('Kein gültiges Foto.');
      const mime = /^image\/(jpeg|png|gif|bmp|webp)$/.test(String(d.mime || '')) ? d.mime : 'image/jpeg';
      t = 'OK\n' + typenschildAuswerten(googleOcr(b64, mime));
    } catch (x) { t = 'FEHLER\n' + x.message; }
    return ContentService.createTextOutput(t).setMimeType(ContentService.MimeType.TEXT);
  }
  // Allgemeine Vorlagen: Anmeldung, gemeinsame Geräteliste, Typenschild per Handy (Vorlagen.js)
  if (d && d.key === VORLAGEN_KEY && /^(v_anmelden|v_code|g_holen|g_speichern|ts_start|ts_ergebnis)$/.test(String(d.event || ''))) {
    let t;
    try { t = 'OK\n' + vorlagenAnfrage(d); } catch (x) { t = 'FEHLER\n' + x.message; }
    return ContentService.createTextOutput(t).setMimeType(ContentService.MimeType.TEXT);
  }
  // Neue Register für die Allgemeinen Vorlagen (Excel-Dateien aus einem Drive-Ordner)
  if (d && (d.key === VORLAGEN_KEY || d.key === KEY) && /^register_(liste|holen)$/.test(String(d.event || ''))) {
    let t;
    try { t = 'OK\n' + registerAnfrage(d); } catch (x) { t = 'FEHLER\n' + x.message; }
    return ContentService.createTextOutput(t).setMimeType(ContentService.MimeType.TEXT);
  }
  // Handy-Seite (GitHub Pages, ohne Google-Anmeldung): nur mit gültigem QR-Code-Token
  if (d && /^h_(check|foto|ts|vorschau|vorschaubild)$/.test(String(d.event || ''))) {
    let t;
    try { t = 'OK\n' + handyAnfrage(d); } catch (x) { t = 'FEHLER\n' + x.message; }
    return ContentService.createTextOutput(t).setMimeType(ContentService.MimeType.TEXT);
  }
  // Service-Mappe: gemeinsame Geräteliste (Formulare)
  if (d && d.key === KEY && /^[gbma]_(holen|speichern)$/.test(String(d.event || ''))) {
    let t;
    try { t = 'OK\n' + vorlagenAnfrage(d, true); } catch (x) { t = 'FEHLER\n' + x.message; }
    return ContentService.createTextOutput(t).setMimeType(ContentService.MimeType.TEXT);
  }
  // Fotos vom Handy (QR-Code) – ohne Sperre
  if (d && d.key === KEY && /^foto(start|liste|holen|ende|vorschau)$/.test(String(d.event || ''))) {
    try { return out(fotoAnfrage(d)); } catch (x) { return out({ ok: false, error: 'Fotos: ' + x.message }); }
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
  // „PDF zu PowerPoint“ meldet sich mit app: 'pptx' – eigene Nutzerliste, sonst gleicher Ablauf.
  const pptx = d.app === 'pptx';
  const APP_NAME = pptx ? 'PDF zu PowerPoint' : 'Service-Mappe';
  const users = sheet(ss, pptx ? 'Nutzer PDF zu PowerPoint' : 'Nutzer', ['Name', 'E-Mail', 'Rechner-ID', 'Version', 'Erste Anmeldung', 'Zuletzt aktiv',
    'Freigeschaltet', 'Gesperrt (ja)', 'Starts', 'Gespeichert', 'Gedruckt', 'Per E-Mail']);
  const log = sheet(ss, 'Protokoll', ['Zeit', 'Name', 'E-Mail', 'Ereignis', 'Version', 'Anzahl']);

  const id = clean(d.id, 64), name = clean(d.name, 100), mail = clean(d.email, 150).toLowerCase();
  const ev = clean(d.event, 20), ver = clean(d.version, 20);
  const n = Math.max(1, Math.min(1000, Number(d.count) || 1));
  if (!id) return { ok: false, error: 'Rechner-ID fehlt.' };
  if (ADMIN_MAILS.indexOf(mail) < 0 && (!/^[a-z0-9.\-]+@/.test(mail) || !mail.endsWith('@' + DOMAIN))) return { ok: false, error: 'Nur E-Mail-Adressen von ' + DOMAIN + ' sind erlaubt.' };

  const now = new Date();
  if (pptx && ['register', 'update', 'resend', 'verify', 'start', 'check'].indexOf(ev) < 0) return { ok: false, error: 'Nicht erlaubt.' };
  if (ev === 'feedbacklist') return feedbackList(ss, users, id, mail);
  if (ev === 'feedback') return feedback(ss, name, mail, ver, d.text);
  if (ev === 'feedbackreply') return feedbackReply(ss, users, id, mail, d.code, d.text);
  if (ev === 'adminlist') return adminList(ss, users, log, id, mail);
  if (ev === 'feedbackdelete') return feedbackDelete(ss, users, id, mail, d.code, d.text);
  if (ev === 'adminblock') return adminBlock(users, id, mail, d.code, d.text);
  if (ev === 'adminuserdelete') return adminUserDelete(users, id, mail, d.code);
  if (ev === 'admindelete') return adminDelete(ss, users, id, mail, d.code, d.text);
  if (ev !== 'check') log.appendRow([now, name, mail, ev, "'" + (pptx ? 'PPTX ' : '') + ver, n]);
  protokollKuerzen(log, now);

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
      subject: 'Dein Freischaltcode für ' + APP_NAME + ': ' + code,
      body: 'Hallo ' + name + ',\n\ndein Freischaltcode für ' + APP_NAME + ' lautet:\n\n    ' + code +
        '\n\nTrag ihn im Programm ein. Der Code ist ' + CODE_MINUTEN + ' Minuten gültig.\n\n' +
        'Falls du dich nicht angemeldet hast, kannst du diese E-Mail ignorieren.\n\nViele Grüße\n' + APP_NAME,
      name: APP_NAME
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
      body: name + ' (' + mail + ', Version ' + ver + ') schreibt:\n\n' + t, name: 'Service-Mappe' });
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
    subject: 'Antwort auf deinen Verbesserungsvorschlag für die Service-Mappe',
    body: 'Hallo ' + r[1] + ',\n\ndanke für deinen Vorschlag' + (when ? ' vom ' + when : '') + ':\n\n    „' +
      String(r[4]).replace(/\n/g, '\n    ') + '“\n\n' + t + '\n\nViele Grüße\nService-Mappe',
    name: 'Service-Mappe'
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

// Anmeldung (Zeile in „Nutzer“) löschen – auf dem Rechner muss man sich danach neu anmelden
function adminUserDelete(users, id, mail, targetId) {
  if (!isAdmin(users, id, mail)) return { ok: false, error: 'Nur für den Administrator.' };
  targetId = String(targetId || '');
  if (!targetId || targetId === id) return { ok: false, error: 'Die eigene Anmeldung dieser App kann nicht gelöscht werden.' };
  const row = findRow(users, targetId);
  if (row < 0) return { ok: false, error: 'Anmeldung nicht gefunden.' };
  users.deleteRow(row);
  return { ok: true, deleted: true };
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

// ---------------- Fotos vom Handy (QR-Code) ----------------
// Das Programm holt sich einen Einmal-Link (gilt FOTO_MIN Minuten) und zeigt ihn als QR-Code. Das Handy öffnet
// die Upload-Seite (doGet), wählt die Original-Fotos aus der Galerie (mit Aufnahmezeit) und lädt sie in einen
// eigenen Ordner in Drive. Das Programm holt sie dort ab und bestätigt den Empfang, dann wird jedes Foto endgültig gelöscht.
const FOTO_MIN = 30;

function fotoOrdner() {
  const name = 'PDF Unterschreiben – Handy-Fotos (temporär)';
  const it = DriveApp.getRootFolder().getFoldersByName(name);
  return it.hasNext() ? it.next() : DriveApp.getRootFolder().createFolder(name);
}

function fotoSitzung(token, nachlauf) {
  if (!/^[a-f0-9]{40}$/.test(String(token || ''))) return null;
  const v = PropertiesService.getScriptProperties().getProperty('foto_' + token);
  if (!v) return null;
  const s = JSON.parse(v);
  return s.exp + (nachlauf ? 3600000 : 0) > Date.now() ? s : null;
}

function driveLoeschen(id) {
  UrlFetchApp.fetch('https://www.googleapis.com/drive/v3/files/' + id,
    { method: 'delete', headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true });
}

// abgelaufene Sitzungen (älter als 1 Stunde nach Ablauf) samt Ordner entfernen
function fotoAufraeumen() {
  const p = PropertiesService.getScriptProperties(), alle = p.getProperties();
  Object.keys(alle).forEach(function (k) {
    if (k.indexOf('foto_') !== 0) return;
    try {
      const s = JSON.parse(alle[k]);
      if (s.exp + 3600000 < Date.now()) { driveLoeschen(s.folder); p.deleteProperty(k); }
    } catch (x) { p.deleteProperty(k); }
  });
}

function fotoAnfrage(d) {
  const mail = clean(d.email, 150).toLowerCase();
  if (ADMIN_MAILS.indexOf(mail) < 0 && !mail.endsWith('@' + DOMAIN)) return { ok: false, error: 'Nicht berechtigt.' };
  if (d.event === 'fotostart') {
    fotoAufraeumen();
    const token = (Utilities.getUuid() + Utilities.getUuid()).replace(/-/g, '').slice(0, 40);
    const folder = fotoOrdner().createFolder('upload-' + token.slice(0, 8));
    PropertiesService.getScriptProperties().setProperty('foto_' + token,
      JSON.stringify({ exp: Date.now() + FOTO_MIN * 60000, folder: folder.getId() }));
    return { ok: true, token: token, url: (neueVersion(d.version, '1.20.1') ? HANDY_URL_NEU : neueVersion(d.version, '1.17.0') ? HANDY_URL : ScriptApp.getService().getUrl()) + '?foto=' + token, minuten: FOTO_MIN };
  }
  const s = fotoSitzung(d.token, true);
  if (!s) return { ok: false, abgelaufen: true, error: 'Der QR-Code ist abgelaufen.' };
  if (d.event === 'fotoliste') {
    // vom Programm bestätigte Fotos jetzt endgültig löschen
    const ack = Array.isArray(d.ack) ? d.ack.slice(0, 100) : [];
    ack.forEach(function (id) {
      try { const f = DriveApp.getFileById(String(id)); if (f.getParents().next().getId() === s.folder) driveLoeschen(f.getId()); } catch (x) { }
    });
    const it = DriveApp.getFolderById(s.folder).getFiles(), a = [];
    while (it.hasNext()) { const f = it.next(); if (/^vorschau-/.test(f.getName())) continue; a.push({ id: f.getId(), name: f.getName(), size: f.getSize() }); }
    return { ok: true, files: a, exp: s.exp };
  }
  if (d.event === 'fotoholen') {
    const f = DriveApp.getFileById(String(d.id || ''));
    if (f.getParents().next().getId() !== s.folder) return { ok: false, error: 'Nicht berechtigt.' };
    const b = f.getBlob();
    let info = {};
    try { info = JSON.parse(f.getDescription() || '{}') || {}; } catch (x) { info = { lm: Number(f.getDescription()) || 0 }; }
    const r = { ok: true, name: f.getName(), mime: b.getContentType(), lm: info.lm || 0, exif: info.exif ? 1 : 0, b64: Utilities.base64Encode(b.getBytes()) };
    return r;   // gelöscht wird erst nach der Empfangsbestätigung (ack bei fotoliste) bzw. bei fotoende
  }
  // Vorschau für die Kunden-Unterschrift: das Programm legt die Seiten der Berichte als Bilder ab
  if (d.event === 'fotovorschau') {
    const b64 = String(d.b64 || '');
    if (b64.length < 100 || b64.length > 4000000) return { ok: false, error: 'Vorschau zu groß.' };
    const nr = ('000' + (Number(d.nr) || 0)).slice(-3);
    const f = DriveApp.getFolderById(s.folder).createFile(Utilities.newBlob(Utilities.base64Decode(b64), 'image/jpeg', 'vorschau-' + nr + '.jpg'));
    f.setDescription(JSON.stringify({ titel: clean(d.titel, 120) }));
    return { ok: true };
  }
  if (d.event === 'fotoende') {
    driveLoeschen(s.folder);
    PropertiesService.getScriptProperties().deleteProperty('foto_' + d.token);
    return { ok: true };
  }
  return { ok: false, error: 'Unbekannte Anfrage.' };
}

// vom Handy aufgerufen (google.script.run)
function fotoHochladen(token, name, lm, mime, b64, exif) {
  const s = fotoSitzung(token, false);
  if (!s) throw new Error('Der QR-Code ist abgelaufen. Bitte im Programm einen neuen anzeigen lassen.');
  mime = String(mime || 'image/jpeg');
  if (!/^image\//.test(mime)) throw new Error('Nur Fotos.');
  b64 = String(b64 || '');
  if (b64.length < 100 || b64.length > 40000000) throw new Error('Foto zu groß.');
  const ordner = DriveApp.getFolderById(s.folder);
  name = String(name || 'foto.jpg').replace(/[\\\/:*?"<>|]/g, '_').slice(0, 120);
  const file = ordner.createFile(Utilities.newBlob(Utilities.base64Decode(b64), mime, name));
  file.setDescription(JSON.stringify({ lm: Number(lm) || 0, exif: exif ? 1 : 0 }));
  return true;
}

// Die Handy-Seite liegt auf GitHub Pages (docs/handy im Repository) und schickt die Fotos per doPost.
// Über Google (doGet) gab es Fehler, wenn im Handy-Browser mehrere Google-Konten angemeldet sind.
const HANDY_URL = 'https://pdfunterschreiben-admin.github.io/service-mappe/handy/';
// ab Service-Mappe 1.20.1 unter eigener Adresse (ältere Programme kennen nur die alte)
const HANDY_URL_NEU = 'https://service-mappe.github.io/handy/';
// Programme vor 1.17.0 zeigen nur QR-Codes für script.google.com an
function neueVersion(v, ab) {
  const a = String(v || '0').split('.').map(x => parseInt(x, 10) || 0), b = ab.split('.').map(Number);   // „1.19.7-test.5“ -> 1, 19, 7
  for (let i = 0; i < 3; i++) { if ((a[i] || 0) !== b[i]) return (a[i] || 0) > b[i]; }
  return true;
}
function handyAnfrage(d) {
  const token = String(d.token || '');
  if (d.event === 'h_check') {
    if (d.modus === 'ts') return CacheService.getScriptCache().get('ts_' + token.replace(/[^a-f0-9]/g, '')) === 'WARTEN' ? '1' : '0';
    return fotoSitzung(token, false) ? '1' : '0';
  }
  if (d.event === 'h_foto') { fotoHochladen(token, d.name, d.lm, d.mime, d.b64, d.exif); return '1'; }
  // Vorschau der Berichte für die Kunden-Unterschrift
  if (d.event === 'h_vorschau' || d.event === 'h_vorschaubild') {
    const s = fotoSitzung(token, false);
    if (!s) throw new Error('Der QR-Code ist abgelaufen.');
    const it = DriveApp.getFolderById(s.folder).getFiles(), a = [];
    while (it.hasNext()) { const f = it.next(); if (/^vorschau-/.test(f.getName())) a.push(f); }
    a.sort(function (x, y) { return x.getName() < y.getName() ? -1 : 1; });
    if (d.event === 'h_vorschau') return JSON.stringify(a.map(function (f) { let t = ''; try { t = JSON.parse(f.getDescription() || '{}').titel || ''; } catch (x) { } return { id: f.getId(), titel: t }; }));
    const f = a.filter(function (x) { return x.getId() === String(d.id || ''); })[0];
    if (!f) throw new Error('Nicht gefunden.');
    return Utilities.base64Encode(f.getBlob().getBytes());
  }
  if (d.event === 'h_ts') return tsFoto(token, d.b64, d.mime) ? '1' : '0';
  throw new Error('Unbekannte Anfrage.');
}

// alte Handy-Seite über Google (für QR-Codes, die noch auf script.google.com zeigen)
function doGet(e) {
  // Typenschild per Handy (Allgemeine Vorlagen)
  const ts = String((e && e.parameter && e.parameter.ts) || '').replace(/[^a-f0-9]/g, '');
  if (ts) {
    const v = CacheService.getScriptCache().get('ts_' + ts);
    const html = v === 'WARTEN' ? TS_SEITE.replace('%TOKEN%', ts) : FOTO_ABGELAUFEN.replace('„Fotos vom Handy“', '„Typenschild einlesen“');
    return HtmlService.createHtmlOutput(html).setTitle('Typenschild fotografieren')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }
  const token = String((e && e.parameter && e.parameter.foto) || '');
  const ok = !!fotoSitzung(token, false);
  const html = ok ? FOTO_SEITE.replace('%TOKEN%', token) : FOTO_ABGELAUFEN;
  return HtmlService.createHtmlOutput(html).setTitle('Fotos senden – Service-Mappe')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

const FOTO_ABGELAUFEN = '<div style="font:16px system-ui,sans-serif;padding:24px"><h2>Link abgelaufen</h2>' +
  '<p>Bitte im Programm „Service-Mappe“ bei der Dokumentation erneut auf „Fotos vom Handy“ tippen und den neuen QR-Code scannen.</p></div>';

const FOTO_SEITE = '<style>body{font:16px system-ui,sans-serif;margin:0;padding:20px;color:#1a1a1a;background:#f6f6f4}' +
  'h2{margin:0 0 6px;font-size:20px}.hint{color:#666;font-size:14px;margin:0 0 18px}' +
  'label.btn{display:block;text-align:center;background:#c8102e;color:#fff;border-radius:10px;padding:16px;font-weight:600;font-size:17px}' +
  'input{display:none}#liste{margin-top:18px;display:flex;flex-direction:column;gap:6px}' +
  '.z{display:flex;justify-content:space-between;background:#fff;border-radius:8px;padding:10px 12px;font-size:14px}' +
  '.ok{color:#1d7a35}.err{color:#c8102e}#fertig{margin-top:18px;font-weight:600;color:#1d7a35}</style>' +
  '<h2>Fotos an die Service-Mappe senden</h2>' +
  '<p class="hint">Fotos aus der Galerie auswählen (Typenschild zuerst). Sie werden für die Übertragung verkleinert und erscheinen gleich im Programm am PC.</p>' +
  '<label class="btn">Fotos auswählen<input type="file" id="f" accept="image/*" multiple></label>' +
  '<div id="liste"></div><div id="fertig"></div>' +
  '<script>' + "var T=\"%TOKEN%\",q=[],aktiv=0,n=0,MAX=3;\nfunction exifZeit(buf){try{var v=new DataView(buf);if(v.getUint16(0)!==0xFFD8)return 0;var o=2;\nwhile(o<v.byteLength-10){var m=v.getUint16(o),len=v.getUint16(o+2);\nif(m===0xFFE1&&v.getUint32(o+4)===0x45786966){var t=o+10,le=v.getUint16(t)===0x4949;\nvar u16=function(p){return v.getUint16(p,le)},u32=function(p){return v.getUint32(p,le)};\nvar ex=0,dt=0,scan=function(p){var c=u16(p);for(var i=0;i<c;i++){var e=p+2+i*12,tag=u16(e);\nif(tag===0x8769)ex=t+u32(e+8);if(tag===0x9003||(tag===0x0132&&!dt))dt=t+u32(e+8);}};\nscan(t+u32(t+4));if(ex)scan(ex);if(!dt)return 0;var s=\"\";for(var k=0;k<19;k++)s+=String.fromCharCode(v.getUint8(dt+k));\nvar a=/(\\d{4}):(\\d{2}):(\\d{2}) (\\d{2}):(\\d{2}):(\\d{2})/.exec(s);return a?new Date(+a[1],a[2]-1,+a[3],+a[4],+a[5],+a[6]).getTime():0;}\nif((m&0xFF00)!==0xFF00)break;o+=2+len;}}catch(e){}return 0;}\nfunction lesen(f,wie){return new Promise(function(ok,no){var r=new FileReader();r.onload=function(){ok(r.result)};r.onerror=no;wie===\"buf\"?r.readAsArrayBuffer(f):r.readAsDataURL(f);});}\nfunction verkleinern(f){return createImageBitmap(f).then(function(b){var s=Math.min(1,2400/Math.max(b.width,b.height)),c=document.createElement(\"canvas\");\nc.width=Math.round(b.width*s);c.height=Math.round(b.height*s);c.getContext(\"2d\").drawImage(b,0,0,c.width,c.height);return c.toDataURL(\"image/jpeg\",0.88).split(\",\")[1];});}\ndocument.getElementById(\"f\").onchange=function(e){for(var i=0;i<e.target.files.length;i++){var f=e.target.files[i],z=document.createElement(\"div\");\nz.className=\"z\";z.innerHTML=\"<span></span><span>wartet …</span>\";z.firstChild.textContent=f.name;document.getElementById(\"liste\").appendChild(z);q.push({f:f,z:z,nr:q.length});}\ne.target.value=\"\";for(var k=0;k<MAX;k++)weiter();};\nfunction weiter(){if(aktiv>=MAX)return;var j=q.shift();if(!j){if(!aktiv&&n)document.getElementById(\"fertig\").textContent=\"Fertig – \"+n+\" Foto(s) gesendet. Du kannst weitere auswählen oder die Seite schließen.\";return;}\naktiv++;j.z.lastChild.textContent=\"wird vorbereitet …\";var zeit=0;\nlesen(j.f.slice(0,262144),\"buf\").then(function(b){zeit=exifZeit(b);return verkleinern(j.f).catch(function(){return lesen(j.f).then(function(d){return String(d).split(\",\")[1];});});})\n.then(function(b64){var versuch=0;var senden=function(){versuch++;j.z.lastChild.textContent=versuch>1?\"neuer Versuch (\"+versuch+\") …\":\"wird gesendet …\";\ngoogle.script.run.withSuccessHandler(function(){n++;j.z.lastChild.textContent=\"✓ gesendet\";j.z.lastChild.className=\"ok\";aktiv--;weiter();})\n.withFailureHandler(function(err){var t=String(err&&err.message||err);if(versuch<3&&!/abgelaufen/.test(t)){setTimeout(senden,2000*versuch);return;}\nj.z.lastChild.textContent=\"Fehler: \"+t;j.z.lastChild.className=\"err\";aktiv--;weiter();})\n.fotoHochladen(T,j.f.name.replace(/\\.[^.]+$/,\"\")+\".jpg\",zeit||j.f.lastModified,\"image/jpeg\",b64,zeit?1:0);};senden();})\n.catch(function(err){j.z.lastChild.textContent=\"Fehler: \"+err;j.z.lastChild.className=\"err\";aktiv--;weiter();});}\n" + '</script>';

// ---------------- Protokoll klein halten ----------------
// Einmal am Tag: Einträge, die älter als PROTOKOLL_MONATE sind, löschen (das Protokoll ist nach Zeit sortiert).
// Die Nutzerliste mit den Zählern (Starts, Gespeichert …) bleibt vollständig.
const PROTOKOLL_MONATE = 12;
function protokollKuerzen(log, now) {
  const p = PropertiesService.getScriptProperties();
  const heute = Utilities.formatDate(now, 'Europe/Berlin', 'yyyy-MM-dd');
  if (p.getProperty('protokoll_gekuerzt') === heute) return;
  p.setProperty('protokoll_gekuerzt', heute);
  try {
    const last = log.getLastRow();
    if (last < 3) return;
    const grenze = new Date(now.getTime());
    grenze.setMonth(grenze.getMonth() - PROTOKOLL_MONATE);
    const zeiten = log.getRange(2, 1, Math.min(last - 1, 50000), 1).getValues();
    let k = 0;
    while (k < zeiten.length && zeiten[k][0] instanceof Date && zeiten[k][0] < grenze) k++;
    if (k > 0 && k < last - 1) log.deleteRows(2, k);
  } catch (x) { }
}

// ---------------- Neue Register für die Allgemeinen Vorlagen ----------------
// Excel-Dateien (oder Google-Tabellen) in diesem Drive-Ordner werden von den Allgemeinen Vorlagen beim Öffnen
// automatisch als neues Register eingefügt (das erste Blatt der Datei, Registername = Dateiname).
// Eine geänderte Datei ersetzt das Register beim nächsten Öffnen.
const REGISTER_ORDNER = 'Allgemeine Vorlagen – neue Register';
function registerOrdner() {
  const it = DriveApp.getRootFolder().getFoldersByName(REGISTER_ORDNER);
  return it.hasNext() ? it.next() : DriveApp.getRootFolder().createFolder(REGISTER_ORDNER);
}
function registerAnfrage(d) {
  const ordner = registerOrdner();
  const excel = /spreadsheet|excel|officedocument\.spreadsheetml/i;
  if (d.event === 'register_liste') {
    const it = ordner.getFiles(), z = [];
    while (it.hasNext()) {
      const f = it.next();
      if (f.isTrashed() || !excel.test(f.getMimeType())) continue;
      const name = f.getName().replace(/\.(xlsx|xlsm|xls)$/i, '').replace(/[|\r\n]/g, ' ').trim();
      z.push(f.getId() + '|' + f.getLastUpdated().getTime() + '|' + name);
    }
    return z.join('\n');
  }
  // register_holen
  const f = DriveApp.getFileById(String(d.id || ''));
  if (f.getParents().next().getId() !== ordner.getId()) throw new Error('Nicht berechtigt.');
  let bytes, ext;
  if (f.getMimeType() === 'application/vnd.google-apps.spreadsheet') {
    bytes = UrlFetchApp.fetch('https://www.googleapis.com/drive/v3/files/' + f.getId() +
      '/export?mimeType=application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() } }).getBlob().getBytes();
    ext = 'xlsx';
  } else {
    bytes = f.getBlob().getBytes();
    const m = /\.(xlsx|xlsm|xls)$/i.exec(f.getName());
    ext = m ? m[1].toLowerCase() : 'xlsx';
  }
  if (bytes.length > 20000000) throw new Error('Datei zu groß.');
  return ext + '\n' + Utilities.base64Encode(bytes);
}
