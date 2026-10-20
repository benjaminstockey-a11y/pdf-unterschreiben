// ---------------- Typenschild auswerten (für die Excel-Vorlage „Allgemeine Vorlagen“) ----------------
// Gleiche Regeln wie assets/TypenschildOCR.ps1 (-TextFile): aus dem von Google erkannten Text werden
// Hersteller, Typ, Fabr.-Nr., Baujahr, Schutzklasse, Ladestrom, Spannung und Strom geholt.
// Ergebnis: Zeilen „SCHLUESSEL=Wert“, danach „#ROH“ und der erkannte Text.

var TS_BRANDS = [
  ['micro\\s*power', 'Micropower'], ['\\bIEB\\b|ELEKTRONIK\\s+Brilon|\\bBrilon\\b', 'IEB'], ['AXIMA', 'AXIMA'],
  ['S\\.?\\s*p\\.?\\s*E\\.?\\s*(ELETTRONICA|INDUSTRIALE)|CREVALCORE|\\bS\\.P\\.E\\b', 'SPE'], ['Exide', 'Exide'], ['Nil\\s*f\\s*isk|vNiIfisk', 'Nilfisk'],
  ['\\biaE\\b|powertron|Untergr', 'iaE'], ['Fronius', 'Fronius'], ['Hoppecke', 'Hoppecke'], ['EnerSys', 'EnerSys'],
  ['Jungheinrich', 'Jungheinrich'], ['\\bLinde\\b', 'Linde'], ['\\bStill\\b', 'Still'], ['\\bBHT\\b', 'BHT'], ['Deltec', 'Deltec'],
  ['Zivan', 'Zivan'], ['Mentzer', 'Mentzer'], ['Benning', 'Benning'], ['Delta-?Q', 'Delta-Q'], ['CTEK', 'CTEK'],
  ['Hawker', 'Hawker'], ['\\bGNB\\b', 'GNB'], ['K(ä|a)rcher', 'Kärcher'], ['CHARIS|\\bIES\\b', 'IES'], ['\\bFerro\\b', 'Ferro']
];
var TS_STOP = /\s*[•·|].*$|\s+(?:Before|Charger|read|For|S\/?N|SN|SER(?:IAL)?|P\/N|CP\/N|DATE|DAT|Datum|INPUT|IN|OUT|OUTPUT|Ue|Ua|Batterie\S*|CE)\b.*$/i;
var TS_LABEL = /^\s*(TYPE?|Typ|MODEL|S\s*\/\s*N|SN|C?P\s*\/\s*N|Ser(?:ial)?\.?\s*(?:No|Nr|#+)?\.?|Art\.?\s*No\.?|Volts?|Amps?|Ue|Ua|[lI]e\s*N?|[iI]a\s*N?|Date|DAT\.?|Baujahr)\s*[:.#]*\s*$/i;

function tsM(s, p, ci) {
  var re = new RegExp(p, ci === false ? '' : 'i');
  var m = re.exec(s);
  return m ? { ok: true, g: m, index: m.index, len: m[0].length } : { ok: false, g: [] };
}
function tsNum(v) { return String(v).replace(/,/g, '.'); }
function tsFix(l) {
  l = l.replace(/\b(\d{1,2})\s?S\s?A\b/g, '$1.5A');
  return l.replace(/\b([0-9IOloS]{2,4})([AV])\b/g, function (all, n, u) {
    if (!/[0-9O]/i.test(n)) return all;
    return n.replace(/[Il]/gi, '1').replace(/[Oo]/gi, '0').replace(/S/gi, '5') + u;
  });
}
function tsVolt(v) {
  v = String(v).replace(/\s/g, '');
  var m = /^(\d{2,3})-(\d{3})$/.exec(v);
  if (m) {
    var lo = +m[1], hi = +m[2];
    if (lo <= 230 && hi >= 230 && hi < 300) return '230';
    if (lo <= 400 && hi >= 400) return '400';
    return v;
  }
  if (/^2[23]\d$/.test(v)) return '230';
  if (/^(38\d|39\d|40\d|41\d)$/.test(v)) return '400';
  return v;
}
function tsTrim(v, chars) {
  var s = v, i = 0, j = s.length;
  while (i < j && chars.indexOf(s[i]) >= 0) i++;
  while (j > i && chars.indexOf(s[j - 1]) >= 0) j--;
  return s.slice(i, j);
}

function tsParse(lines0) {
  var lines = lines0.filter(function (x) { return x && x.trim(); }).map(tsFix);
  var t = lines.join('\n');
  var d = {}, i, l, m;
  var has = function (k) { return Object.prototype.hasOwnProperty.call(d, k); };
  for (i = 0; i < TS_BRANDS.length; i++) if (tsM(t, TS_BRANDS[i][0]).ok) { d.HERSTELLER = TS_BRANDS[i][1]; break; }

  // Typ
  m = tsM(t, 'CHARIS\\s*HF\\s*(\\d)');
  if (m.ok) d.TYP = 'Charis HF' + m.g[1];
  if (!has('TYP')) {
    for (i = 0; i < lines.length; i++) {
      l = lines[i];
      if (tsM(l, 'Batter|Bat\\.|Setting|Gewicht|Weight').ok) continue;
      m = tsM(l, '(?:^|\\s)(?:TYPE?|Mod[eo]ll?)\\s*[:.]?\\s+([A-Za-z0-9].{2,})$');
      if (m.ok) {
        var v = tsTrim(m.g[1].replace(TS_STOP, '').trim(), ' :.');
        if (v.length >= 3 && (/\d/.test(v) || /^[A-Z]{4}/.test(v))) { d.TYP = v; break; }
      }
    }
  }
  if (!has('TYP')) {
    for (i = 0; i < lines.length; i++) {
      l = lines[i];
      m = tsM(l, '(powertron\\s+[A-Z]{1,2})\\b');
      if (!m.ok) m = tsM(l, '\\bM[.,]\\s*(C[A-Z0-9]+(?:\\s*NA)?\\s+\\d{2}-\\d{1,2})');
      if (!m.ok) m = tsM(l, '\\b(SP\\d{2}S\\d{3}S-?\\d{3,5})');
      if (m.ok) { d.TYP = m.g[1].trim(); break; }
    }
  }

  // Fabrik-/Seriennummer
  for (i = 0; i < lines.length; i++) {
    l = lines[i];
    m = tsM(l, 'Ger(ä|a)te-?\\s*Nr\\.?\\s*[:.]?\\s*(\\d[\\d ]{4,}\\d)');
    if (m.ok) { d.FABRNR = m.g[2].replace(/\s+/g, ' ').trim(); break; }
    m = tsM(l, '(?:S/N|\\bSN\\b|Se[rn][il]?(?:al)?\\.?\\s*(?:No|Nr|#)\\.?|Sen\\.?\\s*No\\.?|Serien-?\\s*Nr\\.?|N\\S?\\s*s(é|e)rie|Fabr?\\.?\\s*-?\\s*Nr\\.?|\\bSER\\b\\.?)\\s*[:.#]?\\s*([A-Z0-9][A-Z0-9\\-/]{3,})');
    if (m.ok && /\d/.test(m.g[2])) { d.FABRNR = m.g[2]; break; }
  }
  if (!has('FABRNR')) { m = tsM(t, '\\*\\s*(\\d{5,14})\\s*\\*'); if (m.ok) d.FABRNR = m.g[1]; }
  if (!has('FABRNR')) { m = tsM(t, '\\b([A-Z]{3}-SW\\d{6,})'); if (m.ok) d.FABRNR = m.g[1]; }
  if (!has('FABRNR')) {
    for (i = 0; i < lines.length; i++) {
      l = lines[i];
      if (tsM(l, 'P\\S?N|PIN|Art|Sales|Order|Prod|REF').ok) continue;
      m = /^\s*([1-9]\d{9,13})\s*$/.exec(l);
      if (m) { d.FABRNR = m[1]; break; }
    }
  }

  // Baujahr
  for (i = 0; i < lines.length; i++) {
    l = lines[i];
    m = tsM(l, "(?:Date|DAT|Datum|Baujahr|Bj|Year|Fab)\\.?\\s*[:.]?\\s*(?:\\d{1,2}\\s*[./-]\\s*)?(?:\\d{1,2}\\s*[./-]\\s*)((?:19|20)?\\d{2})(?![\\d'\\u2019`])");
    if (!m.ok) m = tsM(l, '(?:Baujahr|Bj\\.?|Year)\\s*[:.]?\\s*((?:19|20)\\d{2})');
    if (!m.ok) m = tsM(l, '^\\s*\\d{2}\\s*/\\s*((?:19|20)\\d{2})\\s*$');
    if (!m.ok) m = tsM(l, '(?<![\\d.])\\d{2}\\.\\d{2}\\.((?:19|20)\\d{2})(?![\\d.])');
    if (m.ok) {
      var y = m.g[1]; if (y.length === 2) y = '20' + y;
      if (+y >= 1980 && +y <= 2035) { d.BAUJAHR = y; break; }
    }
  }

  // Netzseite: Spannung + Strom
  var sec = '';
  for (i = 0; i < lines.length; i++) {
    l = lines[i];
    if (/^\s*(Input|Eingang)\b/i.test(l)) sec = 'in'; else if (/^\s*(Output|Ausgang)\b/i.test(l)) sec = 'out';
    var mv = tsM(l, '\\bVolt\\S*\\s*[:.]?\\s*(\\d{2,3})\\b');
    var ma = tsM(l, '\\bAmp\\S*\\s*[:.]?\\s*(\\d{1,3}(?:[.,]\\d+)?)\\b');
    if (sec === 'in') {
      if (mv.ok && !has('SPANNUNG')) d.SPANNUNG = mv.g[1];
      if (ma.ok && !has('STROM')) d.STROM = tsNum(ma.g[1]);
    } else if (sec === 'out') {
      if (ma.ok && !has('LEISTUNG')) d.LEISTUNG = tsNum(ma.g[1]);
    }
  }
  for (i = 0; i < lines.length; i++) {
    l = lines[i];
    m = tsM(l, '\\bUe\\b\\D{0,4}(\\d{3})\\s*V', false);
    if (m.ok && !has('SPANNUNG')) d.SPANNUNG = m.g[1];
    m = tsM(l, '\\bIe\\s*N?\\b\\D{0,4}(\\d{1,2}(?:[.,]\\d+)?)\\s*A\\b', false);
    if (m.ok && !has('STROM')) d.STROM = tsNum(m.g[1]);
    m = tsM(l, '\\bIa\\s*N?\\b\\D{0,4}(\\d{1,3}(?:[.,]\\d+)?)\\s*A\\b', false);
    if (m.ok && !has('LEISTUNG')) d.LEISTUNG = tsNum(m.g[1]);
  }
  for (i = 0; i < lines.length; i++) {
    l = lines[i].split(/\bOUT/i)[0];
    if (!tsM(l, 'input|^\\s*IN\\b|\\bIN\\.|\\bin\\s+AC|netz|eingang|prim|VAC|\\bAC\\b|1x|3x|~|N\\s*PE').ok) continue;
    m = tsM(l, '(\\d{3}(?:\\s*-\\s*\\d{3})?)\\s*(?:V|1/)', false);
    if (!m.ok) m = tsM(l, '^\\s*IN\\.?\\s+(\\d{3})\\b');
    if (!m.ok) continue;
    if (!has('SPANNUNG')) d.SPANNUNG = m.g[1].replace(/\s/g, '');
    var rest = l.substring(m.index + m.len);
    var nv = tsM(rest, '\\d+\\s*V', false);
    if (nv.ok) rest = rest.substring(0, nv.index);
    var a = tsM(rest, '(\\d{1,2}(?:[.,]\\d+)?)(?:\\s*-\\s*\\d{1,2}(?:[.,]\\d+)?)?\\s*A\\b', false);
    if (a.ok && !has('STROM')) d.STROM = tsNum(a.g[1]);
    if (has('STROM')) break;
  }
  if (!has('STROM')) {
    for (i = 0; i < lines.length; i++) {
      l = lines[i];
      if (tsM(l, 'DC|out|Ah\\b').ok) continue;
      m = tsM(l, '(?:N\\s*PE|\\bIN\\.?|Hz)\\D{0,6}?(\\d{1,2}(?:[.,]\\d+)?)\\s*A\\b');
      if (m.ok) { d.STROM = tsNum(m.g[1]); break; }
    }
  }
  if (has('SPANNUNG')) d.SPANNUNG = tsVolt(d.SPANNUNG);
  if (!has('SPANNUNG')) {
    if (tsM(t, '3\\s*/?\\s*N\\s*/?\\s*PE|3\\s*~|3x\\s*400').ok) d.SPANNUNG = '400';
    else if (tsM(t, '1\\s*/?\\s*N\\s*/?\\s*PE|1x\\s*2[23]0').ok) d.SPANNUNG = '230';
    else {
      for (i = 0; i < lines.length; i++) {
        l = lines[i];
        if (tsM(l, 'Ah|DC|out').ok) continue;
        m = tsM(l, '(?<![\\d.,])(220-240|100-240|200-240|380-415|230|400)(?![\\d.,])', false);
        if (m.ok) {
          d.SPANNUNG = tsVolt(m.g[1]);
          if (!has('STROM')) {
            var r2 = l.substring(m.index + m.len).replace(/^\s*V/i, '');
            var nv2 = tsM(r2, '\\d+\\s*V', false);
            if (nv2.ok) r2 = r2.substring(0, nv2.index);
            var a2 = tsM(r2, '(\\d{1,2}(?:[.,]\\d+)?)(?:\\s*-\\s*\\d{1,2}(?:[.,]\\d+)?)?\\s*A\\b', false);
            if (a2.ok) d.STROM = tsNum(a2.g[1]);
          }
          break;
        }
      }
    }
  }

  // Ladestrom (Ausgang) -> Feld „Nennleistung“ in A
  if (!has('LEISTUNG')) {
    for (i = 0; i < lines.length; i++) {
      m = tsM(lines[i], '\\b(\\d{1,2})\\s*V\\s*(?:DC|[^\\d\\s]{1,2})?\\s*(\\d{1,3}(?:[.,]\\d+)?)(?:\\s*-\\s*(\\d{1,3}))?\\s*A\\b', false);
      if (m.ok && +m.g[1] <= 96) { d.LEISTUNG = tsNum(m.g[3] !== undefined ? m.g[3] : m.g[2]); break; }
    }
  }
  if (!has('LEISTUNG')) {
    for (i = 0; i < lines.length; i++) {
      m = tsM(lines[i], '\\bOUT\\S*\\D{0,12}?(\\d{1,3})\\s*A\\b');
      if (m.ok) { d.LEISTUNG = m.g[1]; break; }
    }
  }
  if (!has('LEISTUNG')) {
    for (i = 0; i < lines.length; i++) {
      m = /^\s*(\d{2,3})\s*A\s*$/i.exec(lines[i]);
      if (m && +m[1] >= 20) { d.LEISTUNG = m[1]; break; }
    }
  }
  if (has('LEISTUNG')) d.EINHEIT = 'A';

  // Schutzklasse
  m = tsM(t, '(?:\\bcl|class|Schutzklasse|\\bSK)\\s*\\.?\\s*([I1l]{1,3}|[23])\\b');
  if (m.ok) { var g = m.g[1]; d.SCHUTZKLASSE = (g === '2' || g === '3') ? g : String(g.length); }
  if (!has('SCHUTZKLASSE')) d.SCHUTZKLASSE = '1';
  var o = {};
  ['HERSTELLER', 'TYP', 'FABRNR', 'BAUJAHR', 'SCHUTZKLASSE', 'LEISTUNG', 'EINHEIT', 'SPANNUNG', 'STROM'].forEach(function (k) { if (has(k)) o[k] = d[k]; });
  return o;
}

// Google liefert Beschriftung und Wert oft in getrennten Zeilen ("S/N" / "123456").
function tsJoinLabels(lines) {
  var out = [], i = 0;
  while (i < lines.length) {
    var labs = [];
    while (i < lines.length && TS_LABEL.test(lines[i])) { labs.push(lines[i].trim()); i++; }
    if (labs.length === 0) {
      var l = lines[i];
      if (i + 1 < lines.length && /(N\s*\/?\s*PE|Hz)\s*$/i.test(l) && /^\s*\d{1,2}(?:[.,]\d+)?\s*A\s*$/i.test(lines[i + 1])) { l = l + ' ' + lines[i + 1].trim(); i++; }
      if (i + 1 < lines.length && /Ger(ä|a)te-?\s*Nr/i.test(l) && /^\s*\d{3,5}\s*$/.test(lines[i + 1])) { l = l + ' ' + lines[i + 1].trim(); i++; }
      out.push(l); i++; continue;
    }
    var k = 0;
    while (k < labs.length && (i + k) < lines.length) {
      var v = lines[i + k];
      if (TS_LABEL.test(v) || /^\s*(Input|Output|IN|OUT|Eingang|Ausgang)\b/i.test(v) || !/\d/.test(v)) break;
      k++;
    }
    for (var j = 0; j < labs.length; j++) {
      var lab = labs[j].replace(/^\s*le/, 'Ie').replace(/^\s*ia/, 'Ia');
      if (j < k) {
        var val = lines[i + j].trim();
        if (k === 1 && /^\d+(?:[.,]\d+)?$/.test(val) && (i + 1) < lines.length && /^\s*A\b/i.test(lines[i + 1]) && /^(Ie|Ia)/i.test(lab)) val += ' A';
        out.push(lab + ' ' + val);
      } else out.push(lab);
    }
    i += k;
  }
  return out;
}

// Text von Google -> Ergebnis-Text für die Vorlage
function typenschildAuswerten(text) {
  var lines = String(text || '').replace(/﻿/g, '').replace(/\r/g, '').split('\n')
    .filter(function (x) { return x.trim() && !/^_+$/.test(x.trim()); });
  var joined = tsJoinLabels(lines);
  var res = tsParse(joined);
  if (!res.LEISTUNG) {
    var o = false;
    for (var i = 0; i < joined.length; i++) {
      var l = joined[i];
      if (/^\s*(Output|OUT)\b/i.test(l)) { o = true; continue; }
      var m = /^\s*(\d{2,3})\s*A\b/i.exec(l);
      if (o && m && +m[1] >= 20) { res.LEISTUNG = m[1]; res.EINHEIT = 'A'; break; }
    }
    if (res.LEISTUNG) {  // Reihenfolge wie im Skript: LEISTUNG/EINHEIT am Ende
      var r2 = {}; Object.keys(res).forEach(function (k) { if (k !== 'LEISTUNG' && k !== 'EINHEIT') r2[k] = res[k]; });
      r2.LEISTUNG = res.LEISTUNG; r2.EINHEIT = 'A'; res = r2;
    }
  }
  if (res.LEISTUNG && res.SPANNUNG && res.LEISTUNG === res.SPANNUNG) { delete res.LEISTUNG; delete res.EINHEIT; }
  if (res.TYP) res.TYP = res.TYP.replace(/^\s*MODEL\s+/i, '').trim();
  if (!res.TYP) {
    for (var j = 0; j < lines.length; j++) {
      var mm = /^\s*(SP\d{2}[A-Z0-9 ]{3,}-\d{4,5})\s*$/i.exec(lines[j]);
      if (mm) { res.TYP = mm[1]; break; }
    }
  }
  var s = Object.keys(res).map(function (k) { return k + '=' + res[k]; }).join('\r\n');
  return s + '\r\n#ROH\r\n' + lines.join('\r\n') + '\r\n';
}

if (typeof module !== 'undefined') module.exports = { typenschildAuswerten: typenschildAuswerten };
