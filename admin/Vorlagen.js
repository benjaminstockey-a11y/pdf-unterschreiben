// ---------------- Allgemeine Vorlagen: Anmeldung, gemeinsame Geräteliste, Typenschild per Handy ----------------
// Die Excel-Vorlage meldet sich mit VORLAGEN_KEY. Gemeinsame Geräteliste und Handy-Foto gibt es nur für
// Vorlagen-Nutzer, die sich einmal mit ihrer Firmen-E-Mail und einem Code angemeldet haben (Blatt „Nutzer Vorlagen“).

var V_NUTZER = 'Nutzer Vorlagen', V_GERAETE = 'Geräteliste (Vorlagen)';
var B_PRUEF = 'Batterieprüfungen';
var M_DATEN = 'Meine Daten';
var K_LISTE = 'Kunden';
var K_KOPF = ['Kunde', 'Straße / Hausnummer', 'PLZ / Ort', 'Ansprechpartner', 'Telefon', 'E-Mail', 'Stand (ms)', 'Zuletzt gespeichert', 'Von'];
var M_KOPF = ['E-Mail', 'Daten', 'Stand (ms)', 'Zuletzt gespeichert', 'Name'];
// Firmenweite Einstellungen (Zeile „*alle*“ im Blatt „Meine Daten“): wer sie ändern darf
var A_SCHLUESSEL = { '§batGrenz': ['j.niemann@schlueter-gabelstapler.de', 'b.stockey@schlueter-gabelstapler.de', 'b.stockey@hotmail.de'] };
var B_KOPF = ['Batterie-Nr.', 'Prüfdatum', 'Kunde', 'Fahrzeug', 'Hersteller', 'Zellen', 'Zellentype', 'Mittel Volt', 'Mittel Säure', 'Auffällig', 'Zellwerte', 'Stand (ms)', 'Zuletzt gespeichert', 'Von'];
var V_GERAETE_KOPF = ['Fabr. Nr.', 'Hersteller', 'Typenbezeichnung', 'Baujahr', 'Schutzklasse', 'Nennleistung', 'Einheit',
  'Nennspannung (V)', 'Nennstrom (A)', 'Geräteart', 'Prüflingsnr.', 'Stand (ms)', 'Zuletzt gespeichert', 'Von'];

function vNorm(v) {
  return String(v || '').toUpperCase().trim().replace(/[\s\-\/.]/g, '').replace(/O/g, '0').replace(/[IL]/g, '1');
}
function vNutzerBlatt() {
  return sheet(SpreadsheetApp.getActiveSpreadsheet(), V_NUTZER,
    ['Name', 'E-Mail', 'Vorlagen-ID', 'Erste Anmeldung', 'Zuletzt aktiv', 'Freigeschaltet', 'Code', 'Code-Zeit', 'Gesperrt (ja)']);
}
function vZeile(s, vid) {
  if (!vid || s.getLastRow() < 2) return -1;
  const ids = s.getRange(2, 3, s.getLastRow() - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) if (String(ids[i][0]) === vid) return i + 2;
  return -1;
}
// freigeschaltet und nicht gesperrt? (aktualisiert „Zuletzt aktiv“)
function vFrei(vid) {
  const s = vNutzerBlatt(), r = vZeile(s, vid);
  if (r < 0) return false;
  const z = s.getRange(r, 1, 1, 9).getValues()[0];
  if (String(z[5]).toLowerCase() !== 'ja' || String(z[8]).toLowerCase() === 'ja') return false;
  s.getRange(r, 5).setValue(new Date());
  return true;
}

// programm: Anfrage kommt von der Service-Mappe (mit dem Programm-Schlüssel, Nutzer schon angemeldet)
function vorlagenAnfrage(d, programm) {
  const ev = String(d.event || ''), vid = clean(d.vid, 64);
  if (!programm && !/^[A-Za-z0-9\-]{16,64}$/.test(vid)) throw new Error('Ungültige Vorlagen-ID.');
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    if (ev === 'v_anmelden') {
      const name = clean(d.name, 100), mail = clean(d.email, 150).toLowerCase();
      if (ADMIN_MAILS.indexOf(mail) < 0 && (!/^[a-z0-9.\-]+@/.test(mail) || !mail.endsWith('@' + DOMAIN)))
        throw new Error('Nur E-Mail-Adressen von ' + DOMAIN + ' sind erlaubt.');
      const s = vNutzerBlatt(), now = new Date();
      let r = vZeile(s, vid);
      if (r < 0) { s.appendRow([name, mail, vid, now, now, 'nein', '', '', '']); r = s.getLastRow(); }
      const code = String(Math.floor(100000 + Math.random() * 900000));
      s.getRange(r, 1, 1, 2).setValues([[name, mail]]);
      s.getRange(r, 6, 1, 3).setValues([['nein', code, now]]);
      MailApp.sendEmail({ to: mail, name: 'Allgemeine Vorlagen',
        subject: 'Dein Freischaltcode für die Allgemeinen Vorlagen: ' + code,
        body: 'Hallo ' + name + ',\n\ndein Freischaltcode für die Allgemeinen Vorlagen lautet:\n\n    ' + code +
          '\n\nTrag ihn in Excel ein. Der Code ist ' + CODE_MINUTEN + ' Minuten gültig.\n\n' +
          'Falls du dich nicht angemeldet hast, kannst du diese E-Mail ignorieren.\n\nViele Grüße\nAllgemeine Vorlagen' });
      return 'GESENDET';
    }
    if (ev === 'v_code') {
      const s = vNutzerBlatt(), r = vZeile(s, vid);
      if (r < 0) throw new Error('Bitte zuerst anmelden.');
      const z = s.getRange(r, 1, 1, 9).getValues()[0];
      if (String(z[8]).toLowerCase() === 'ja') throw new Error('Dieser Zugang ist gesperrt.');
      const alt = z[7] instanceof Date ? (Date.now() - z[7].getTime()) / 60000 : 1e9;
      if (String(d.code || '').trim() !== String(z[6]) || !z[6] || alt > CODE_MINUTEN) throw new Error('Der Code stimmt nicht oder ist abgelaufen.');
      s.getRange(r, 6, 1, 3).setValues([['ja', '', '']]);
      s.getRange(r, 5).setValue(new Date());
      return 'FREI';
    }
    if (!programm && !vFrei(vid)) return 'NICHT_FREI';

    if (ev === 'g_holen') {
      const g = sheet(SpreadsheetApp.getActiveSpreadsheet(), V_GERAETE, V_GERAETE_KOPF);
      const seit = Number(d.seit) || 0, jetzt = Date.now(), z = [String(jetzt)];
      if (g.getLastRow() >= 2) {
        g.getRange(2, 1, g.getLastRow() - 1, 12).getValues().forEach(function (row) {
          if (Number(row[11]) > seit) z.push(row.slice(0, 11).map(function (v) {
            if (v instanceof Date) v = Utilities.formatDate(v, 'Europe/Berlin', 'dd.MM.yyyy');
            return String(v).replace(/[\t\r\n]/g, ' ');
          }).join('\t'));
        });
      }
      return z.join('\n');
    }
    if (ev === 'g_speichern') {
      const g = sheet(SpreadsheetApp.getActiveSpreadsheet(), V_GERAETE, V_GERAETE_KOPF);
      const zeilen = Array.isArray(d.geraete) ? d.geraete.slice(0, 500) : [];
      const s = vNutzerBlatt(), wer = programm ? clean(d.name, 100) : s.getRange(vZeile(s, vid), 1).getValue();
      const last = g.getLastRow(), keys = last >= 2 ? g.getRange(2, 1, last - 1, 1).getValues().map(function (r) { return vNorm(r[0]); }) : [];
      let n = 0;
      zeilen.forEach(function (f) {
        if (!Array.isArray(f)) return;
        f = f.slice(0, 11).map(function (v) { return clean(v, 120); });
        while (f.length < 11) f.push('');
        const k = vNorm(f[0]);
        if (k.length < 4) return;
        const row = f.concat([Date.now(), new Date(), wer]);
        const i = keys.indexOf(k);
        if (i >= 0) g.getRange(i + 2, 1, 1, 14).setValues([row]);
        else { g.appendRow(row); keys.push(k); }
        g.getRange(i >= 0 ? i + 2 : g.getLastRow(), 1).setNumberFormat('@').setValue(f[0]);
        n++;
      });
      return String(n);
    }
    // Service-Mappe: gemeinsame Kundenliste (aus allen Formularen, für alle Techniker)
    if (ev === 'k_holen') {
      const b = sheet(SpreadsheetApp.getActiveSpreadsheet(), K_LISTE, K_KOPF);
      const seit = Number(d.seit) || 0, z = [String(Date.now())];
      if (b.getLastRow() >= 2) {
        b.getRange(2, 1, b.getLastRow() - 1, 9).getValues().forEach(function (row) {
          if (Number(row[6]) > seit) z.push(row.slice(0, 6).map(function (v) { return String(v).replace(/[\t\r\n]/g, ' '); }).join('\t'));
        });
      }
      return z.join('\n');
    }
    if (ev === 'k_speichern') {
      if (!programm) throw new Error('Nicht erlaubt.');
      const k = d.kunde || {};
      const f = [k.name, k.strasse, k.ort, k.ansprech, k.tel, k.mail].map(function (v) { return clean(v, 200); });
      if (vNorm(f[0]).length < 2) throw new Error('Kundenname fehlt.');
      const lock = LockService.getScriptLock(); lock.waitLock(20000);
      try {
        const b = sheet(SpreadsheetApp.getActiveSpreadsheet(), K_LISTE, K_KOPF), last = b.getLastRow();
        const schl = vNorm(f[0]) + '|' + vNorm(f[2]);
        const keys = last >= 2 ? b.getRange(2, 1, last - 1, 3).getValues().map(function (r) { return vNorm(r[0]) + '|' + vNorm(r[2]); }) : [];
        const i = keys.indexOf(schl), r = i >= 0 ? i + 2 : last + 1;
        let alt = i >= 0 ? b.getRange(r, 1, 1, 6).getValues()[0] : ['', '', '', '', '', ''];
        const neu = f.map(function (v, n) { return v || alt[n] || ''; });   // leere Felder löschen nichts
        b.getRange(r, 1, 1, 9).setValues([neu.concat([Date.now(), new Date(), clean(d.name, 100)])]);
        return '1';
      } finally { lock.releaseLock(); }
    }
    // Service-Mappe: firmenweite Einstellungen, z. B. Grenzwerte der Batterieprüfung (legt Jannik Niemann fest)
    if (ev === 'a_holen' || ev === 'a_speichern') {
      if (!programm) throw new Error('Nicht erlaubt.');
      const mail = String(d.email || '').trim().toLowerCase();
      const lock = LockService.getScriptLock(); lock.waitLock(20000);
      try {
        const b = sheet(SpreadsheetApp.getActiveSpreadsheet(), M_DATEN, M_KOPF), last = b.getLastRow();
        const mails = last >= 2 ? b.getRange(2, 1, last - 1, 1).getValues().map(function (r) { return String(r[0]).trim().toLowerCase(); }) : [];
        const i = mails.indexOf('*alle*');
        let daten = {};
        if (i >= 0) { try { daten = JSON.parse(String(b.getRange(i + 2, 2).getValue() || '{}')) || {}; } catch (x) { daten = {}; } }
        if (ev === 'a_speichern') {
          const neu = d.daten || {};
          Object.keys(neu).forEach(function (k) {
            const erlaubt = A_SCHLUESSEL[k];
            if (!erlaubt) throw new Error('Unbekannte Einstellung.');
            if (erlaubt.indexOf(mail) < 0) throw new Error('Diese Einstellung darf nur Jannik Niemann ändern.');
            daten[k] = { v: String(neu[k].v == null ? '' : neu[k].v).slice(0, 5000), t: Date.now(), von: clean(d.name, 100) };
          });
          const r = i >= 0 ? i + 2 : last + 1;
          b.getRange(r, 1, 1, 5).setValues([['*alle*', JSON.stringify(daten), Date.now(), new Date(), clean(d.name, 100)]]);
        }
        const darf = {};
        Object.keys(A_SCHLUESSEL).forEach(function (k) { darf[k] = A_SCHLUESSEL[k].indexOf(mail) >= 0; });
        return JSON.stringify({ daten: daten, darf: darf });
      } finally { lock.releaseLock(); }
    }
    // Service-Mappe: eigene Daten je Nutzer (E-Mail aus der Anmeldung) – z. B. eigene Artikel, Textbausteine,
    // Messgeräte – auf allen Rechnern des Nutzers gleich. Je Eintrag gewinnt der neueste Stand.
    if (ev === 'm_holen' || ev === 'm_speichern') {
      if (!programm) throw new Error('Nicht erlaubt.');
      const mail = String(d.email || '').trim().toLowerCase();
      if (!/^[^@\s]+@[^@\s]+$/.test(mail)) throw new Error('Nicht angemeldet.');
      const lock = LockService.getScriptLock(); lock.waitLock(20000);
      try {
        const b = sheet(SpreadsheetApp.getActiveSpreadsheet(), M_DATEN, M_KOPF), last = b.getLastRow();
        const mails = last >= 2 ? b.getRange(2, 1, last - 1, 1).getValues().map(function (r) { return String(r[0]).trim().toLowerCase(); }) : [];
        const i = mails.indexOf(mail);
        let daten = {};
        if (i >= 0) { try { daten = JSON.parse(String(b.getRange(i + 2, 2).getValue() || '{}')) || {}; } catch (x) { daten = {}; } }
        if (ev === 'm_speichern') {
          const neu = d.daten || {};
          let geaendert = false;
          Object.keys(neu).forEach(function (k) {
            const e = neu[k];
            if (!e || typeof e.t !== 'number' || String(k).length > 200) return;
            if (!daten[k] || Number(daten[k].t) < e.t) { daten[k] = { v: String(e.v == null ? '' : e.v).slice(0, 20000), t: e.t }; geaendert = true; }
          });
          if (geaendert) {
            const txt = JSON.stringify(daten);
            if (txt.length > 49000) throw new Error('Zu viele eigene Daten.');
            const r = i >= 0 ? i + 2 : last + 1;
            b.getRange(r, 1, 1, 5).setValues([[mail, txt, Date.now(), new Date(), clean(d.name, 100)]]);
          }
        }
        return JSON.stringify(daten);
      } finally { lock.releaseLock(); }
    }
    // Service-Mappe: gemeinsamer Batterie-Verlauf (Batterieprüfprotokolle aller Techniker)
    if (ev === 'b_holen') {
      const b = sheet(SpreadsheetApp.getActiveSpreadsheet(), B_PRUEF, B_KOPF);
      const seit = Number(d.seit) || 0, z = [String(Date.now())];
      if (b.getLastRow() >= 2) {
        b.getRange(2, 1, b.getLastRow() - 1, 14).getValues().forEach(function (row) {
          if (Number(row[11]) > seit) z.push(row.slice(0, 11).concat([row[13]]).map(function (v) {
            if (v instanceof Date) v = Utilities.formatDate(v, 'Europe/Berlin', 'yyyy-MM-dd');
            return String(v).replace(/[\t\r\n]/g, ' ');
          }).join('\t'));
        });
      }
      return z.join('\n');
    }
    if (ev === 'b_speichern') {
      if (!programm) throw new Error('Nicht erlaubt.');
      const b = sheet(SpreadsheetApp.getActiveSpreadsheet(), B_PRUEF, B_KOPF), p = d.pruefung || {};
      const f = [p.nr, p.datum, p.kunde, p.fahrzeug, p.hersteller, p.zellen, p.type, p.mv, p.ms, p.auff].map(function (v) { return clean(v, 200); }).concat([clean(p.werte, 3000)]);
      const k = vNorm(f[0]) + '|' + f[1];
      if (vNorm(f[0]).length < 2 || !/^\d{4}-\d{2}-\d{2}$/.test(f[1])) throw new Error('Batterie-Nr. oder Datum fehlt.');
      const last = b.getLastRow(), keys = last >= 2 ? b.getRange(2, 1, last - 1, 2).getValues().map(function (r) {
        return vNorm(r[0]) + '|' + (r[1] instanceof Date ? Utilities.formatDate(r[1], 'Europe/Berlin', 'yyyy-MM-dd') : String(r[1]));
      }) : [];
      const row = f.concat([Date.now(), new Date(), clean(d.name, 100)]), i = keys.indexOf(k), r = i >= 0 ? i + 2 : last + 1;
      b.getRange(r, 1, 1, 2).setNumberFormat('@');
      b.getRange(r, 1, 1, 14).setValues([row]);
      return '1';
    }
    if (ev === 'ts_start') {
      const token = Utilities.getUuid().replace(/-/g, '').slice(0, 24);
      CacheService.getScriptCache().put('ts_' + token, 'WARTEN', 1200);
      const url = HANDY_URL + '?ts=' + token;
      return token + '\n' + url + '\n' + qrBmpBase64(url, 6);
    }
    if (ev === 'ts_ergebnis') {
      const v = CacheService.getScriptCache().get('ts_' + String(d.token || '').replace(/[^a-f0-9]/g, ''));
      return v === null ? 'ABGELAUFEN' : v;
    }
    throw new Error('Unbekannte Anfrage.');
  } finally { lock.releaseLock(); }
}

// vom Handy (google.script.run): Foto lesen, Ergebnis für Excel bereitlegen
function tsFoto(token, b64, mime) {
  token = String(token || '').replace(/[^a-f0-9]/g, '');
  const c = CacheService.getScriptCache(), v = c.get('ts_' + token);
  if (v === null) throw new Error('Der QR-Code ist abgelaufen. Bitte in Excel nochmal auf „Typenschild einlesen“ klicken.');
  if (v !== 'WARTEN') {
    if (v.indexOf('OK') === 0) return true;   // schon gelesen (z. B. zweiter Versuch nach Verbindungsabbruch)
    throw new Error(v === 'LIEST' ? 'Das Foto wird schon gelesen – bitte kurz warten und in Excel nachsehen.' : 'Für diesen QR-Code wurde schon ein Foto geschickt.');
  }
  b64 = String(b64 || '');
  if (b64.length < 100 || b64.length > 15000000) throw new Error('Kein gültiges Foto.');
  c.put('ts_' + token, 'LIEST', 1200);
  let t;
  try { t = 'OK\n' + typenschildAuswerten(googleOcr(b64, /^image\//.test(String(mime)) ? mime : 'image/jpeg')); }
  catch (x) { t = 'FEHLER\n' + x.message; }
  c.put('ts_' + token, t.slice(0, 90000), 1200);
  return t.indexOf('OK') === 0;
}

var TS_SEITE = '<style>body{font:16px system-ui,sans-serif;margin:0;padding:20px;color:#1a1a1a;background:#f6f6f4}' +
  'h2{margin:0 0 6px;font-size:20px}.hint{color:#666;font-size:14px;margin:0 0 18px}' +
  'label.btn{display:block;text-align:center;background:#c8102e;color:#fff;border-radius:10px;padding:18px;font-weight:600;font-size:18px;margin-bottom:10px}' +
  'label.btn.zwei{background:#fff;color:#c8102e;border:2px solid #c8102e}input{display:none}#st{margin-top:18px;font-weight:600}.ok{color:#1d7a35}.err{color:#c8102e}</style>' +
  '<h2>Typenschild fotografieren</h2><p class="hint">Das Foto wird gelesen und die Werte erscheinen gleich in Excel.</p>' +
  '<label class="btn">📷 Foto aufnehmen<input type="file" id="k" accept="image/*" capture="environment"></label>' +
  '<label class="btn zwei">Aus der Galerie wählen<input type="file" id="g" accept="image/*"></label><div id="st"></div>' +
  '<script>var T="%TOKEN%",st=document.getElementById("st");' +
  'function senden(f){if(!f)return;st.className="";st.textContent="Foto wird gesendet und gelesen …";' +
  'createImageBitmap(f).then(function(b){var s=Math.min(1,2400/Math.max(b.width,b.height)),c=document.createElement("canvas");c.width=Math.round(b.width*s);c.height=Math.round(b.height*s);' +
  'c.getContext("2d").drawImage(b,0,0,c.width,c.height);return c.toDataURL("image/jpeg",0.9).split(",")[1];})' +
  '.then(function(b64){google.script.run.withSuccessHandler(function(ok){st.className=ok?"ok":"err";st.textContent=ok?"✓ Fertig – die Werte stehen jetzt in Excel. Du kannst diese Seite schließen.":"Das Foto konnte nicht gelesen werden. Bitte in Excel nochmal versuchen.";})' +
  '.withFailureHandler(function(e){st.className="err";st.textContent="Fehler: "+(e&&e.message||e);}).tsFoto(T,b64,"image/jpeg");})' +
  '.catch(function(e){st.className="err";st.textContent="Fehler: "+e;});}' +
  'document.getElementById("k").onchange=function(e){senden(e.target.files[0]);};document.getElementById("g").onchange=function(e){senden(e.target.files[0]);};</script>';

if (typeof module !== 'undefined') module.exports = { vNorm: vNorm };
