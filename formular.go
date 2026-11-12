//go:build windows

package main

import (
	"bytes"
	_ "embed"
	"encoding/base64"
	"encoding/json"
	"errors"
	"fmt"
	"image"
	_ "image/png"
	"io"
	"net/http"
	"os"
	"path/filepath"
	"regexp"
	"strings"
	"time"
)

// Formulare: Die Service-Mappe füllt im Hintergrund das passende Blatt der bekannten Excel-Vorlage
// „Allgemeine Vorlagen“ aus (Excel, Makros aus) und speichert es als Excel-Datei und/oder PDF.
// So sieht das Ergebnis genau aus wie bisher.

//go:embed vorlagen/Allgemeine-Vorlagen.xlsm
var vorlagenEingebaut []byte

// Formulare mit eigener, im Programm mitgelieferter Vorlage (Anfrage: quelle "@name")
//
//go:embed vorlagen/Batteriepruefprotokoll.xlsx
var vorlageBatterie []byte

//go:embed vorlagen/Werkzeugbestellung.xlsx
var vorlageWerkzeug []byte

var eigeneVorlagen = map[string][]byte{"batterie": vorlageBatterie, "werkzeug": vorlageWerkzeug}

const vorlagenURL = "https://raw.githubusercontent.com/pdfunterschreiben-Admin/pdfunterschreiben/main/vorlagen/Allgemeine-Vorlagen.xlsm"

// neueste Vorlage: einmal am Tag von GitHub holen, sonst die zuletzt geholte bzw. die eingebaute
func vorlagenDatei() []byte {
	p := filepath.Join(dataDir(), "Allgemeine-Vorlagen.xlsm")
	if st, err := os.Stat(p); err == nil && time.Since(st.ModTime()) < 24*time.Hour {
		if b, err := os.ReadFile(p); err == nil && len(b) > 10000 {
			return b
		}
	}
	c := &http.Client{Timeout: 20 * time.Second}
	if resp, err := c.Get(vorlagenURL); err == nil {
		defer resp.Body.Close()
		if b, err := io.ReadAll(io.LimitReader(resp.Body, 30<<20)); err == nil && resp.StatusCode == 200 && len(b) > 10000 && bytes.HasPrefix(b, []byte("PK")) {
			_ = os.WriteFile(p, b, 0o644)
			return b
		}
	}
	if b, err := os.ReadFile(p); err == nil && len(b) > 10000 {
		return b
	}
	return vorlagenEingebaut
}

const formularScript = `$ErrorActionPreference = 'Stop'
$cfg = Get-Content -LiteralPath $env:PU_CFG -Raw -Encoding UTF8 | ConvertFrom-Json
$xl = $null
try {
  $xl = New-Object -ComObject Excel.Application
  $xl.Visible = $false
  $xl.DisplayAlerts = $false
  $xl.EnableEvents = $false
  try { $xl.AutomationSecurity = 3 } catch {}
  $wb = $xl.Workbooks.Open($env:PU_IN, 0, $false)
  try {
    if ($cfg.blatt) { $ws = $wb.Worksheets.Item([string]$cfg.blatt) } else { $ws = $wb.Worksheets.Item(1) }
    # zusätzliche Tabellenzeilen: letzte Zeile der Tabelle samt Format kopieren und darunter einfügen
    $nach = [int]$cfg.nach; $anzahl = [int]$cfg.anzahl
    if ($nach -gt 0 -and $anzahl -gt 0) {
      $pa = [string]$ws.PageSetup.PrintArea
      for ($i = 1; $i -le $anzahl; $i++) {
        [void]$ws.Rows($nach).Copy()
        [void]$ws.Rows($nach + $i).Insert(-4121)
      }
      try { $xl.CutCopyMode = 0 } catch {}
      [void]$ws.Range($ws.Rows($nach + 1), $ws.Rows($nach + $anzahl)).ClearContents()
      if ($pa -match '^(.*\$)(\d+)$' -and [int]$matches[2] -ge $nach) { $ws.PageSetup.PrintArea = $matches[1] + ([int]$matches[2] + $anzahl) }
    }
    foreach ($p in $cfg.werte.PSObject.Properties) {
      $c = $ws.Range($p.Name)
      $v = [string]$p.Value
      if ($v -match '^\d{4}-\d{2}-\d{2}$') { $c.Value2 = [double]([datetime]::ParseExact($v, 'yyyy-MM-dd', $null)).ToOADate() }
      else { $c.NumberFormat = '@'; $c.Value2 = $v }   # genau so, wie es eingetippt wurde
    }
    foreach ($p in $cfg.kreuze.PSObject.Properties) {
      try { $ws.CheckBoxes([string]$p.Name).Value = $(if ($p.Value) { 1 } else { -4146 }) } catch {}
    }
    # Unterschriften (z. B. Kunde) als Bild direkt über die Beschriftung setzen
    foreach ($b in @($cfg.bilder)) {
      if (-not $b) { continue }
      try {
        $c = $ws.Range([string]$b.zelle).MergeArea; $bw = [double]$b.w; $bh = [double]$b.h
        # „in“: oben in die Zelle (Beschriftung steht unten in der Zelle); sonst auf die Linie über der Beschriftung
        if ($b.modus -eq 'in') { $top = $c.Top + 1 } else { $top = $c.Top - $bh * 0.82 }   # Unterschrift reicht etwas über die Linie wie von Hand
        $left = $c.Left + [Math]::Max(2, ($c.Width - $bw) / 2)
        $sh = $ws.Shapes.AddPicture([string]$b.pfad, 0, -1, $left, $top, $bw, $bh)
        try { $sh.Placement = 2 } catch {}
      } catch {}
    }
    $ws.Activate()
    $ws.Range('A1').Select()
    # Vorlagen aus dem Drive-Ordner: genau eine Seite (Breite und Höhe)
    if ($cfg.einseite) {
      $ps = $ws.PageSetup
      if (-not $cfg.bereich) { try { $ps.PrintArea = $ws.UsedRange.Address() } catch {} }   # eigene Vorlagen haben einen festen Druckbereich
      $ps.Zoom = $false
      $ps.FitToPagesWide = 1
      $ps.FitToPagesTall = 1
      try { $ps.CenterHorizontally = $true } catch {}
    }
    # Knöpfe der Excel-Vorlage (Typenschild einlesen, Ersatzteile einfügen …) aus allen Registern entfernen
    foreach ($s in @($wb.Worksheets)) {
      foreach ($sh in @($s.Shapes)) {
        $typ = -1; try { $typ = $sh.FormControlType } catch {}
        if ($typ -eq 0 -or $sh.Name -like 'TS_QR*') { $sh.Delete() }
      }
    }
    # gelbe Markierungen der Vorlage für leere Pflichtfelder (bedingte Formatierung) nicht mit ausgeben
    try { [void]$ws.Cells.FormatConditions.Delete() } catch {}
    # Reparaturauftrag: die leere zweite Seite (Fortsetzung Ersatzteile) nicht mitdrucken
    if ($cfg.ersteSeite) {
      try {
        $ws.Activate(); $xl.ActiveWindow.View = 2; $r = $ws.HPageBreaks.Item(1).Location.Row - 1; $xl.ActiveWindow.View = 1
        $pa = [string]$ws.PageSetup.PrintArea
        if ($r -gt 10 -and $pa -match '^(.*\$)(\d+)$') { $ws.PageSetup.PrintArea = $matches[1] + $r }
      } catch {}
    }
    $ws.Activate()
    $ws.Range('A1').Select()
    if ($cfg.pdf) { $ws.ExportAsFixedFormat(0, $env:PU_PDF) }
    if ($cfg.drucken) { [void]$ws.PrintOut() }
    # Excel: die Allgemeine Vorlage mit ihren Registern (ohne Makros, Knöpfe, Geräteliste und Typenschild-Text), nur dieses Formular ausgefüllt
    if ($cfg.xlsx) {
      # Hilfsregister der Allgemeinen Vorlage (Geräteliste, Typenschild-Text) nicht mit in die Excel-Datei
      foreach ($n in @($cfg.ohneRegister)) {
        if (-not $n) { continue }
        try { $s = $wb.Worksheets.Item([string]$n); if ($s.Name -ne $ws.Name -and $wb.Worksheets.Count -gt 1) { $s.Visible = -1; [void]$s.Delete() } } catch {}
      }
      $ws.Activate()
      $wb.SaveAs($env:PU_XLSX, 51)
    }
  } finally { $wb.Close($false) }
} finally {
  if ($xl) { $xl.Quit(); [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) }
}`

var zelleRe = regexp.MustCompile(`^[A-Z]{1,3}[0-9]{1,4}$`)

func (a *app) handleFormular(w http.ResponseWriter, r *http.Request) {
	if r.Method != http.MethodPost {
		http.Error(w, "POST", http.StatusMethodNotAllowed)
		return
	}
	fehler := func(msg string) { writeJSON(w, map[string]any{"ok": false, "error": msg}) }
	var req struct {
		Blatt   string            `json:"blatt"`
		Werte   map[string]string `json:"werte"`
		Kreuze  map[string]bool   `json:"kreuze"`
		XLSX    bool              `json:"xlsx"`
		PDF     bool              `json:"pdf"`
		Drucken bool              `json:"drucken"` // direkt auf dem Standarddrucker ausgeben
		Nach    int               `json:"nach"`    // zusätzliche Tabellenzeilen unter dieser Zeile einfügen
		Anzahl  int               `json:"anzahl"`
		Quelle  string            `json:"quelle"` // Vorlage aus dem Drive-Ordner statt der Allgemeinen Vorlagen
		Bilder  []struct {
			Zelle  string  `json:"zelle"`
			PNG    string  `json:"png"`
			Hoehe  float64 `json:"hoehe"`
			Breite float64 `json:"breite"`
			Modus  string  `json:"modus"`
		} `json:"bilder"` // Unterschriften
	}
	if err := json.NewDecoder(io.LimitReader(r.Body, 4<<20)).Decode(&req); err != nil {
		fehler("Ungültige Anfrage.")
		return
	}
	if !req.XLSX && !req.PDF && !req.Drucken {
		fehler("Bitte Excel und/oder PDF auswählen.")
		return
	}
	for k := range req.Werte {
		if !zelleRe.MatchString(k) {
			fehler("Ungültige Zelle: " + k)
			return
		}
	}
	dir, err := os.MkdirTemp("", "servicemappe-")
	if err != nil {
		fehler("Temporärer Ordner nicht verfügbar.")
		return
	}
	defer os.RemoveAll(dir)
	vorlage, endung := vorlagenDatei(), "xlsm"
	if strings.HasPrefix(req.Quelle, "@") {
		d, ok := eigeneVorlagen[req.Quelle[1:]]
		if !ok {
			fehler("Unbekannte Vorlage.")
			return
		}
		vorlage, endung = d, "xlsx"
	} else if req.Quelle != "" {
		e, d, err := registerDatei(req.Quelle)
		if err != nil {
			fehler(err.Error())
			return
		}
		vorlage, endung = d, e
	}
	in := filepath.Join(dir, "vorlage."+endung)
	if err := os.WriteFile(in, vorlage, 0o600); err != nil {
		fehler("Die Vorlage konnte nicht bereitgelegt werden.")
		return
	}
	var bilder []map[string]any
	for i, b := range req.Bilder {
		png, err := base64.StdEncoding.DecodeString(b.PNG)
		if err != nil || len(png) < 50 || !zelleRe.MatchString(b.Zelle) {
			continue
		}
		zelle := b.Zelle
		if req.Nach > 0 && req.Anzahl > 0 { // zusätzliche Tabellenzeilen verschieben alles darunter
			j := strings.IndexAny(zelle, "0123456789")
			var zeile int
			fmt.Sscanf(zelle[j:], "%d", &zeile)
			if zeile > req.Nach {
				zelle = zelle[:j] + fmt.Sprint(zeile+min(req.Anzahl, 200))
			}
		}
		// Größe aus dem Bild: höchstens „hoehe“ hoch und „breite“ breit, Seitenverhältnis bleibt
		ic, _, err := image.DecodeConfig(bytes.NewReader(png))
		if err != nil || ic.Width == 0 || ic.Height == 0 {
			continue
		}
		hh, bb := max(10, min(b.Hoehe, 80)), max(40, min(b.Breite, 260))
		bw := hh * float64(ic.Width) / float64(ic.Height)
		if bw > bb {
			bw, hh = bb, bb*float64(ic.Height)/float64(ic.Width)
		}
		p := filepath.Join(dir, fmt.Sprintf("unterschrift%d.png", i))
		if os.WriteFile(p, png, 0o600) == nil {
			bilder = append(bilder, map[string]any{"zelle": zelle, "pfad": p, "w": bw, "h": hh, "modus": b.Modus})
		}
	}
	// Excel-Datei aus den Allgemeinen Vorlagen: ohne die Hilfsregister Geräteliste und Typenschild-Text
	ohneRegister := []string{}
	if req.Quelle == "" {
		ohneRegister = []string{"Geräteliste", "Typenschild-Text"}
	}
	cfg, _ := json.Marshal(map[string]any{"bilder": bilder, "blatt": req.Blatt, "nach": req.Nach, "anzahl": min(req.Anzahl, 200), "einseite": req.Quelle != "", "ersteSeite": req.Blatt == "Rep.auftrag+Kostenvor.", "bereich": strings.HasPrefix(req.Quelle, "@"), "drucken": req.Drucken, "werte": req.Werte, "kreuze": req.Kreuze, "xlsx": req.XLSX, "pdf": req.PDF, "ohneRegister": ohneRegister})
	cfgPath := filepath.Join(dir, "auftrag.json")
	_ = os.WriteFile(cfgPath, cfg, 0o600)
	xlsx, pdf := filepath.Join(dir, "formular.xlsx"), filepath.Join(dir, "formular.pdf")

	cmd := hidden("powershell.exe", "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-Command", formularScript)
	cmd.Env = append(os.Environ(), "PU_IN="+in, "PU_CFG="+cfgPath, "PU_XLSX="+xlsx, "PU_PDF="+pdf)
	var out bytes.Buffer
	cmd.Stdout, cmd.Stderr = &out, &out
	runErr := runWithTimeout(cmd, 2*time.Minute)
	res := map[string]any{"ok": true}
	if req.XLSX {
		if b, err := os.ReadFile(xlsx); err == nil && len(b) > 0 {
			res["xlsx"] = base64.StdEncoding.EncodeToString(b)
		}
	}
	if req.PDF {
		if b, err := os.ReadFile(pdf); err == nil && len(b) > 0 {
			res["pdf"] = base64.StdEncoding.EncodeToString(b)
		}
	}
	if (req.XLSX && res["xlsx"] == nil) || (req.PDF && res["pdf"] == nil) {
		s := out.String()
		msg := "Excel konnte das Formular nicht erstellen."
		switch {
		case strings.Contains(s, "80040154") || strings.Contains(strings.ToLower(s), "class not registered") || strings.Contains(s, "Klasse nicht registriert"):
			msg = "Dafür muss Microsoft Excel auf diesem Computer installiert sein."
		case runErr != nil && strings.Contains(runErr.Error(), "Zeit"):
			msg = "Excel hat zu lange gebraucht. Ist vielleicht ein Excel-Fenster mit einer Rückfrage offen?"
		}
		if t := strings.TrimSpace(s); t != "" {
			if len(t) > 400 {
				t = t[:400]
			}
			msg += "\n\n" + t
		}
		fehler(msg)
		return
	}
	writeJSON(w, res)
}

// Anfrage an den Anmeldeserver, Antwort als Text („OK\n…“ bzw. „FEHLER\n…“)
func serverText(payload map[string]any, timeout time.Duration) (string, error) {
	b, _ := json.Marshal(payload)
	c := &http.Client{Timeout: timeout}
	resp, err := c.Post(ReportURL, "application/json", bytes.NewReader(b))
	if err != nil {
		return "", err
	}
	defer resp.Body.Close()
	body, err := io.ReadAll(io.LimitReader(resp.Body, 20<<20))
	return string(body), err
}

func angemeldeterNutzer() *reportUser {
	var s struct {
		User *reportUser `json:"user"`
	}
	if b, err := os.ReadFile(settingsPath()); err == nil {
		_ = json.Unmarshal(b, &s)
	}
	return s.User
}

// Typenschild-Foto lesen (Google-Texterkennung + Auswertung auf dem Server)
func (a *app) handleFormularTypenschild(w http.ResponseWriter, r *http.Request) {
	if r.Method != http.MethodPost {
		http.Error(w, "POST", http.StatusMethodNotAllowed)
		return
	}
	img, err := io.ReadAll(io.LimitReader(r.Body, 15<<20))
	if err != nil || len(img) < 100 {
		writeJSON(w, map[string]any{"ok": false, "error": "Kein Foto."})
		return
	}
	if ReportURL == "" {
		writeJSON(w, map[string]any{"ok": false, "error": "Kein Server eingerichtet."})
		return
	}
	t, err := serverText(map[string]any{"key": ReportKey, "event": "typenschild",
		"b64": base64.StdEncoding.EncodeToString(img), "mime": http.DetectContentType(img)}, 90*time.Second)
	if err != nil {
		writeJSON(w, map[string]any{"ok": false, "error": "Keine Verbindung zum Server."})
		return
	}
	t = strings.ReplaceAll(t, "\r", "")
	if !strings.HasPrefix(t, "OK\n") {
		writeJSON(w, map[string]any{"ok": false, "error": strings.TrimSpace(strings.TrimPrefix(t, "FEHLER\n"))})
		return
	}
	werte := map[string]string{}
	roh := ""
	if i := strings.Index(t, "\n#ROH\n"); i >= 0 {
		roh = t[i+6:]
	}
	for _, z := range strings.Split(t[3:], "\n") {
		if z == "#ROH" {
			break
		}
		if i := strings.Index(z, "="); i > 0 && !strings.HasPrefix(z, "#") {
			if v := strings.TrimSpace(z[i+1:]); v != "" {
				werte[z[:i]] = v
			}
		}
	}
	// erkannten Text mitschreiben (hilft beim Verbessern der Erkennung), Datei bleibt klein
	lp := filepath.Join(dataDir(), "texterkennung.log")
	if alt, err := os.ReadFile(lp); err == nil && len(alt) > 200<<10 {
		_ = os.WriteFile(lp, alt[len(alt)-100<<10:], 0o644)
	}
	if f, err := os.OpenFile(lp, os.O_APPEND|os.O_CREATE|os.O_WRONLY, 0o644); err == nil {
		_, _ = f.WriteString("=== " + time.Now().Format("2006-01-02 15:04:05") + "\n" + roh + "\n")
		_ = f.Close()
	}
	writeJSON(w, map[string]any{"ok": true, "werte": werte, "roh": roh})
}

// gemeinsame Geräteliste (dieselbe wie in den Allgemeinen Vorlagen)
func (a *app) handleFormularGeraete(w http.ResponseWriter, r *http.Request) {
	if r.Method != http.MethodPost {
		http.Error(w, "POST", http.StatusMethodNotAllowed)
		return
	}
	var req struct {
		Event   string            `json:"event"`
		Seit    int64             `json:"seit"`
		Geraete [][]string        `json:"geraete"`
		Pruef   map[string]string `json:"pruefung"`
		Daten   map[string]any    `json:"daten"`
		Kunde   map[string]string `json:"kunde"`
		Kd      map[string]string `json:"kd"`    // persönliche Kundendaten (Firma, Name, E-Mail)
		Token   string            `json:"token"` // Visitenkarte per Handy: kd_scanergebnis
	}
	if err := json.NewDecoder(io.LimitReader(r.Body, 1<<20)).Decode(&req); err != nil || !map[string]bool{"g_holen": true, "g_speichern": true, "b_holen": true, "b_speichern": true, "m_holen": true, "m_speichern": true, "a_holen": true, "a_speichern": true, "k_holen": true, "k_speichern": true, "kd_holen": true, "kd_speichern": true, "kd_loeschen": true, "kd_scanstart": true, "kd_scanergebnis": true}[req.Event] {
		writeJSON(w, map[string]any{"ok": false, "error": "Ungültige Anfrage."})
		return
	}
	u := angemeldeterNutzer()
	if u == nil || u.Email == "" || ReportURL == "" {
		writeJSON(w, map[string]any{"ok": false, "error": "Bitte zuerst im Programm anmelden."})
		return
	}
	t, err := serverText(map[string]any{"key": ReportKey, "event": req.Event, "seit": req.Seit, "geraete": req.Geraete, "pruefung": req.Pruef, "daten": req.Daten, "kunde": req.Kunde, "kd": req.Kd, "token": req.Token,
		"name": u.Name, "email": u.Email, "version": Version}, 60*time.Second)
	if err != nil {
		writeJSON(w, map[string]any{"ok": false, "offline": true, "error": "Keine Verbindung zum Server."})
		return
	}
	t = strings.ReplaceAll(t, "\r", "")
	if !strings.HasPrefix(t, "OK\n") {
		writeJSON(w, map[string]any{"ok": false, "error": strings.TrimSpace(strings.TrimPrefix(t, "FEHLER\n"))})
		return
	}
	writeJSON(w, map[string]any{"ok": true, "text": t[3:]})
}

// Neue Register aus dem Drive-Ordner (wie in den Allgemeinen Vorlagen): Liste und „In Excel öffnen“
func (a *app) handleFormularRegister(w http.ResponseWriter, r *http.Request) {
	if ReportURL == "" {
		writeJSON(w, map[string]any{"ok": false, "error": "Kein Server eingerichtet."})
		return
	}
	if r.Method == http.MethodGet {
		t, err := serverText(map[string]any{"key": ReportKey, "event": "register_liste"}, 30*time.Second)
		if err != nil {
			writeJSON(w, map[string]any{"ok": false, "offline": true, "error": "Keine Verbindung zum Server."})
			return
		}
		t = strings.ReplaceAll(t, "\r", "")
		if !strings.HasPrefix(t, "OK\n") {
			writeJSON(w, map[string]any{"ok": false, "error": strings.TrimSpace(strings.TrimPrefix(t, "FEHLER\n"))})
			return
		}
		var liste []map[string]string
		for _, z := range strings.Split(t[3:], "\n") {
			f := strings.SplitN(z, "|", 3)
			if len(f) == 3 && f[0] != "" {
				liste = append(liste, map[string]string{"id": f[0], "stand": f[1], "name": strings.TrimSpace(f[2])})
			}
		}
		writeJSON(w, map[string]any{"ok": true, "register": liste})
		return
	}
	var req struct {
		ID   string `json:"id"`
		Name string `json:"name"`
	}
	if err := json.NewDecoder(io.LimitReader(r.Body, 1<<12)).Decode(&req); err != nil || !regexp.MustCompile(`^[A-Za-z0-9_\-]{10,80}$`).MatchString(req.ID) {
		writeJSON(w, map[string]any{"ok": false, "error": "Ungültige Anfrage."})
		return
	}
	ext, data, err := registerDatei(req.ID)
	if err != nil {
		writeJSON(w, map[string]any{"ok": false, "error": err.Error()})
		return
	}
	name := regexp.MustCompile(`[\\/:*?"<>|]`).ReplaceAllString(strings.TrimSpace(req.Name), " ")
	if name == "" {
		name = "Vorlage"
	}
	ordner := filepath.Join(os.Getenv("USERPROFILE"), "Documents", "Formulare")
	_ = os.MkdirAll(ordner, 0o755)
	pfad := filepath.Join(ordner, name+" "+time.Now().Format("2006-01-02 15-04")+"."+ext)
	if err := os.WriteFile(pfad, data, 0o644); err != nil {
		writeJSON(w, map[string]any{"ok": false, "error": "Die Datei konnte nicht gespeichert werden."})
		return
	}
	_ = hidden("rundll32.exe", "url.dll,FileProtocolHandler", pfad).Start()
	writeJSON(w, map[string]any{"ok": true, "pfad": pfad})
}

// Datei einer Vorlage aus dem Drive-Ordner holen (Endung, Inhalt)
func registerDatei(id string) (string, []byte, error) {
	if !regexp.MustCompile(`^[A-Za-z0-9_\-]{10,80}$`).MatchString(id) {
		return "", nil, errors.New("Ungültige Vorlage.")
	}
	t, err := serverText(map[string]any{"key": ReportKey, "event": "register_holen", "id": id}, 60*time.Second)
	if err != nil {
		return "", nil, errors.New("Keine Verbindung zum Server.")
	}
	t = strings.ReplaceAll(t, "\r", "")
	if !strings.HasPrefix(t, "OK\n") {
		return "", nil, errors.New(strings.TrimSpace(strings.TrimPrefix(t, "FEHLER\n")))
	}
	teile := strings.SplitN(t[3:], "\n", 2)
	ext := strings.ToLower(strings.TrimSpace(teile[0]))
	if len(teile) < 2 || (ext != "xlsx" && ext != "xlsm" && ext != "xls") {
		return "", nil, errors.New("Die Vorlage kam nicht richtig an.")
	}
	data, err := base64.StdEncoding.DecodeString(strings.ReplaceAll(teile[1], "\n", ""))
	if err != nil || len(data) == 0 {
		return "", nil, errors.New("Die Vorlage kam nicht richtig an.")
	}
	return ext, data, nil
}

// Sicherung der Formulareingaben: Bei jeder Änderung schickt das Fenster alle
// Eingaben hierher. Stürzt Fenster oder Programm ab, sind sie beim nächsten Öffnen wieder da.
func (a *app) handleFormularSicherung(w http.ResponseWriter, r *http.Request) {
	p := filepath.Join(dataDir(), "formulare-sicherung.json")
	if r.Method == http.MethodPost {
		b, err := io.ReadAll(io.LimitReader(r.Body, 20<<20))
		if err != nil || !json.Valid(b) {
			http.Error(w, "ungültig", http.StatusBadRequest)
			return
		}
		if alt, err := os.ReadFile(p); err == nil && len(alt) > 2 {
			_ = os.WriteFile(p+".bak", alt, 0o644)
		}
		tmp := p + ".tmp"
		if err := os.WriteFile(tmp, b, 0o644); err == nil {
			_ = os.Rename(tmp, p)
		}
		w.WriteHeader(http.StatusNoContent)
		return
	}
	b, err := os.ReadFile(p)
	if err != nil || !json.Valid(b) {
		b, err = os.ReadFile(p + ".bak")
		if err != nil || !json.Valid(b) {
			b = []byte("{}")
		}
	}
	w.Header().Set("Content-Type", "application/json")
	w.Header().Set("Cache-Control", "no-store")
	_, _ = w.Write(b)
}

// Ablage: jedes erstellte Formular zusätzlich unter Dokumente\Service-Mappe\Berichte\<Kunde> speichern
func ablageDir() string {
	return filepath.Join(os.Getenv("USERPROFILE"), "Documents", "Service-Mappe", "Berichte")
}

var ablageName = regexp.MustCompile(`[<>:"/\\|?*\x00-\x1f]+`)

func (a *app) handleFormularAblage(w http.ResponseWriter, r *http.Request) {
	if r.Method != http.MethodPost {
		http.Error(w, "POST", http.StatusMethodNotAllowed)
		return
	}
	basis := ablageDir()
	if r.URL.Query().Get("oeffnen") != "" {
		var req struct {
			Pfad string `json:"pfad"`
		}
		_ = json.NewDecoder(io.LimitReader(r.Body, 1<<16)).Decode(&req)
		p := filepath.Clean(req.Pfad)
		if req.Pfad != "" && strings.HasPrefix(strings.ToLower(p), strings.ToLower(basis)+string(os.PathSeparator)) {
			if _, err := os.Stat(p); err == nil {
				_ = hidden("explorer.exe", "/select,", p).Start()
				w.WriteHeader(http.StatusNoContent)
				return
			}
		}
		_ = os.MkdirAll(basis, 0o755)
		_ = hidden("explorer.exe", basis).Start()
		w.WriteHeader(http.StatusNoContent)
		return
	}
	var req struct {
		Kunde   string `json:"kunde"`
		Dateien []struct {
			Name string `json:"name"`
			Data string `json:"data"`
		} `json:"dateien"`
	}
	if err := json.NewDecoder(io.LimitReader(r.Body, 100<<20)).Decode(&req); err != nil || len(req.Dateien) == 0 {
		writeJSON(w, map[string]any{"ok": false, "error": "Keine Dateien."})
		return
	}
	kunde := strings.TrimSpace(ablageName.ReplaceAllString(req.Kunde, " "))
	kunde = strings.Trim(kunde, ". ")
	if len([]rune(kunde)) > 60 {
		kunde = strings.TrimSpace(string([]rune(kunde)[:60]))
	}
	if kunde == "" {
		kunde = "Ohne Kunde"
	}
	dir := filepath.Join(basis, kunde)
	if err := os.MkdirAll(dir, 0o755); err != nil {
		writeJSON(w, map[string]any{"ok": false, "error": err.Error()})
		return
	}
	var pfade []string
	for _, d := range req.Dateien {
		b, err := base64.StdEncoding.DecodeString(d.Data)
		if err != nil || len(b) == 0 {
			continue
		}
		name := strings.Trim(ablageName.ReplaceAllString(d.Name, " "), ". ")
		if name == "" {
			continue
		}
		p := filepath.Join(dir, name)
		ext := filepath.Ext(name)
		for i := 2; i < 100; i++ { // nichts überschreiben
			if _, err := os.Stat(p); err != nil {
				break
			}
			p = filepath.Join(dir, strings.TrimSuffix(name, ext)+fmt.Sprintf(" (%d)", i)+ext)
		}
		if os.WriteFile(p, b, 0o644) == nil {
			pfade = append(pfade, p)
		}
	}
	writeJSON(w, map[string]any{"ok": true, "pfade": pfade})
}
