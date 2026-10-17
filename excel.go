//go:build windows

package main

import (
	"archive/zip"
	"bytes"
	"encoding/base64"
	"encoding/json"
	"encoding/xml"
	"fmt"
	"io"
	"net/http"
	"os"
	"path/filepath"
	"strings"
	"time"
)

// Excel-Dateien werden mit dem installierten Microsoft Excel in eine PDF umgewandelt –
// so sieht das Ergebnis genau aus wie beim Drucken aus Excel. Makros laufen dabei nicht.

type sheetInfo struct {
	Name   string `json:"name"`
	Active bool   `json:"active"`
}

// listSheets liest die sichtbaren Tabellenblätter einer .xlsx/.xlsm direkt aus der Datei.
func listSheets(b []byte) ([]sheetInfo, error) {
	zr, err := zip.NewReader(bytes.NewReader(b), int64(len(b)))
	if err != nil {
		return nil, err
	}
	var wbXML []byte
	for _, f := range zr.File {
		if f.Name == "xl/workbook.xml" {
			rc, err := f.Open()
			if err != nil {
				return nil, err
			}
			wbXML, err = io.ReadAll(io.LimitReader(rc, 10<<20))
			rc.Close()
			if err != nil {
				return nil, err
			}
		}
	}
	if wbXML == nil {
		return nil, fmt.Errorf("keine Arbeitsmappe gefunden")
	}
	var wb struct {
		Views struct {
			View []struct {
				ActiveTab int `xml:"activeTab,attr"`
			} `xml:"workbookView"`
		} `xml:"bookViews"`
		Sheets struct {
			Sheet []struct {
				Name  string `xml:"name,attr"`
				State string `xml:"state,attr"`
			} `xml:"sheet"`
		} `xml:"sheets"`
	}
	if err := xml.Unmarshal(wbXML, &wb); err != nil {
		return nil, err
	}
	active := 0
	if len(wb.Views.View) > 0 {
		active = wb.Views.View[0].ActiveTab
	}
	var out []sheetInfo
	for i, s := range wb.Sheets.Sheet {
		if s.State == "hidden" || s.State == "veryHidden" {
			continue
		}
		out = append(out, sheetInfo{Name: s.Name, Active: i == active})
	}
	return out, nil
}

func isExcelName(n string) string {
	ext := strings.ToLower(filepath.Ext(n))
	switch ext {
	case ".xlsx", ".xlsm", ".xls", ".xlsb":
		return ext
	}
	return ""
}

func readUpload(w http.ResponseWriter, r *http.Request) ([]byte, string, bool) {
	if r.Method != http.MethodPost {
		http.Error(w, "POST", http.StatusMethodNotAllowed)
		return nil, "", false
	}
	ext := isExcelName(r.URL.Query().Get("name"))
	if ext == "" {
		writeJSON(w, map[string]any{"ok": false, "error": "Das ist keine Excel-Datei."})
		return nil, "", false
	}
	b, err := io.ReadAll(io.LimitReader(r.Body, 60<<20))
	if err != nil || len(b) == 0 {
		writeJSON(w, map[string]any{"ok": false, "error": "Die Datei konnte nicht gelesen werden."})
		return nil, "", false
	}
	return b, ext, true
}

func (a *app) handleExcelSheets(w http.ResponseWriter, r *http.Request) {
	b, ext, ok := readUpload(w, r)
	if !ok {
		return
	}
	if ext == ".xls" || ext == ".xlsb" {
		// Altes Format: Blätter lassen sich nicht ohne Excel auslesen – dann das aktive Blatt nehmen.
		writeJSON(w, map[string]any{"ok": true, "sheets": []sheetInfo{}})
		return
	}
	sheets, err := listSheets(b)
	if err != nil {
		writeJSON(w, map[string]any{"ok": true, "sheets": []sheetInfo{}})
		return
	}
	writeJSON(w, map[string]any{"ok": true, "sheets": sheets})
}

const excelScript = `$ErrorActionPreference = 'Stop'
$xl = $null
try {
  $xl = New-Object -ComObject Excel.Application
  $xl.Visible = $false
  $xl.DisplayAlerts = $false
  try { $xl.AutomationSecurity = 3 } catch {}
  $wb = $xl.Workbooks.Open($env:PU_IN, 0, $true)
  try {
    if ($env:PU_SHEET) { $ws = $wb.Worksheets.Item($env:PU_SHEET) } else { $ws = $wb.ActiveSheet }
    $ws.ExportAsFixedFormat(0, $env:PU_OUT)
  } finally { $wb.Close($false) }
} finally {
  if ($xl) { $xl.Quit(); [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) }
}`

func (a *app) handleExcelPDF(w http.ResponseWriter, r *http.Request) {
	b, ext, ok := readUpload(w, r)
	if !ok {
		return
	}
	dir, err := os.MkdirTemp("", "pdfunterschreiben-")
	if err != nil {
		writeJSON(w, map[string]any{"ok": false, "error": "Temporärer Ordner nicht verfügbar."})
		return
	}
	defer os.RemoveAll(dir)
	in := filepath.Join(dir, "eingabe"+ext)
	out := filepath.Join(dir, "ausgabe.pdf")
	if err := os.WriteFile(in, b, 0o600); err != nil {
		writeJSON(w, map[string]any{"ok": false, "error": "Die Datei konnte nicht zwischengespeichert werden."})
		return
	}
	cmd := hidden("powershell.exe", "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-Command", excelScript)
	cmd.Env = append(os.Environ(), "PU_IN="+in, "PU_OUT="+out, "PU_SHEET="+r.URL.Query().Get("sheet"))
	var stderr bytes.Buffer
	cmd.Stderr = &stderr
	cmd.Stdout = &stderr
	if err := cmd.Start(); err != nil {
		writeJSON(w, map[string]any{"ok": false, "error": "PowerShell konnte nicht gestartet werden."})
		return
	}
	done := make(chan error, 1)
	go func() { done <- cmd.Wait() }()
	select {
	case <-done:
	case <-time.After(2 * time.Minute):
		_ = cmd.Process.Kill()
		writeJSON(w, map[string]any{"ok": false, "error": "Excel hat zu lange gebraucht. Ist vielleicht ein Excel-Fenster mit einer Rückfrage offen?"})
		return
	}
	pdf, err := os.ReadFile(out)
	if err != nil || len(pdf) == 0 {
		msg := "Excel konnte die Datei nicht umwandeln."
		s := stderr.String()
		switch {
		case strings.Contains(s, "80040154") || strings.Contains(strings.ToLower(s), "class not registered") || strings.Contains(s, "Klasse nicht registriert"):
			msg = "Dafür muss Microsoft Excel auf diesem Computer installiert sein."
		case strings.Contains(s, "Worksheets") || strings.Contains(s, "8002000B"):
			msg = "Das gewählte Tabellenblatt wurde nicht gefunden."
		}
		writeJSON(w, map[string]any{"ok": false, "error": msg})
		return
	}
	w.Header().Set("Content-Type", "application/pdf")
	w.Header().Set("Cache-Control", "no-store")
	w.Write(pdf)
}

/* ---------- Unterschrift in die Excel-Datei selbst einfügen ---------- */

type overlay struct {
	PNG string  `json:"png"` // data:image/png;base64,...
	DX  float64 `json:"dx"`  // Abstand zur Beschriftung in PDF-Punkten
	DY  float64 `json:"dy"`
	W   float64 `json:"w"`
	H   float64 `json:"h"`
}

type signReq struct {
	Name     string    `json:"name"`
	Sheet    string    `json:"sheet"`
	File     string    `json:"file"` // base64
	Keyword  string    `json:"kw"`
	FontPt   float64   `json:"fh"` // Schriftgröße der Beschriftung in der PDF
	ItemW    float64   `json:"itemW"`
	Overlays []overlay `json:"overlays"`
}

// Die Positionen kommen relativ zur gefundenen Beschriftung („Servicetechniker“ o. ä.).
// Excel sucht dieselbe Beschriftung im Blatt und rechnet über die Schriftgröße den Druckmaßstab zurück.
const excelSignScript = `$ErrorActionPreference = 'Stop'
$cfg = Get-Content -LiteralPath $env:PU_CFG -Raw -Encoding UTF8 | ConvertFrom-Json
$xl = $null
try {
  $xl = New-Object -ComObject Excel.Application
  $xl.Visible = $false
  $xl.DisplayAlerts = $false
  try { $xl.AutomationSecurity = 3 } catch {}
  $wb = $xl.Workbooks.Open($env:PU_IN, 0, $false)
  try {
    if ($cfg.sheet) { $ws = $wb.Worksheets.Item([string]$cfg.sheet) } else { $ws = $wb.ActiveSheet }
    $rng = $ws.UsedRange
    $miss = [System.Reflection.Missing]::Value
    $cell = $rng.Find([string]$cfg.kw, $miss, -4123, 2)
    if (-not $cell) { $cell = $rng.Find([string]$cfg.kw, $miss, -4163, 2) }
    if (-not $cell) { throw "ANKER_FEHLT" }
    $fr = $cell.Row; $fc = $cell.Column; $best = $cell; $n = 0
    do {
      if ($cell.Top -gt $best.Top -or ($cell.Top -eq $best.Top -and $cell.Left -gt $best.Left)) { $best = $cell }
      $cell = $rng.FindNext($cell)
      $n++
    } while ($cell -and ($cell.Row -ne $fr -or $cell.Column -ne $fc) -and $n -lt 500)
    $c = $best; $m = $c.MergeArea
    $fs = [double]$c.Font.Size; if ($fs -le 0) { $fs = 10 }
    $f = [double]$cfg.fh / $fs; if ($f -le 0.05) { $f = 1 }
    $tw = [double]$cfg.itemW / $f
    $mw = [double]$m.Width; $mh = [double]$m.Height
    $ha = $c.HorizontalAlignment
    if ($ha -eq -4108 -or $ha -eq 7) { $tx = $m.Left + ($mw - $tw) / 2 }
    elseif ($ha -eq -4152) { $tx = $m.Left + $mw - $tw - 2 }
    else { $tx = $m.Left + 2 + [double]$c.IndentLevel * 9 }
    $va = $c.VerticalAlignment
    if ($va -eq -4160) { $base = $m.Top + 1 + $fs * 0.95 }
    elseif ($va -eq -4108) { $base = $m.Top + ($mh + $fs * 0.75) / 2 }
    else { $base = $m.Top + $mh - $fs * 0.22 - 1.5 }
    $i = 0
    foreach ($o in $cfg.overlays) {
      $p = Join-Path $env:PU_DIR ("bild$i.png"); $i++
      $s = $ws.Shapes.AddPicture($p, 0, -1, $tx + [double]$o.dx / $f, $base + [double]$o.dy / $f, [double]$o.w / $f, [double]$o.h / $f)
      $s.Placement = 2
    }
    $wb.Save()
  } finally { $wb.Close($false) }
} finally {
  if ($xl) { $xl.Quit(); [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) }
}`

func (a *app) handleExcelSign(w http.ResponseWriter, r *http.Request) {
	if r.Method != http.MethodPost {
		http.Error(w, "POST", http.StatusMethodNotAllowed)
		return
	}
	fehler := func(msg string) { writeJSON(w, map[string]any{"ok": false, "error": msg}) }
	var req signReq
	if err := json.NewDecoder(io.LimitReader(r.Body, 120<<20)).Decode(&req); err != nil {
		fehler("Ungültige Anfrage.")
		return
	}
	ext := isExcelName(req.Name)
	if ext == "" {
		fehler("Das ist keine Excel-Datei.")
		return
	}
	file, err := base64.StdEncoding.DecodeString(req.File)
	if err != nil || len(file) == 0 {
		fehler("Die Excel-Datei fehlt.")
		return
	}
	dir, err := os.MkdirTemp("", "pdfunterschreiben-")
	if err != nil {
		fehler("Temporärer Ordner nicht verfügbar.")
		return
	}
	defer os.RemoveAll(dir)
	in := filepath.Join(dir, "mappe"+ext)
	if err := os.WriteFile(in, file, 0o600); err != nil {
		fehler("Die Datei konnte nicht zwischengespeichert werden.")
		return
	}
	for i, o := range req.Overlays {
		data := o.PNG
		if k := strings.Index(data, ","); k >= 0 {
			data = data[k+1:]
		}
		png, err := base64.StdEncoding.DecodeString(data)
		if err != nil {
			fehler("Ein Bild konnte nicht gelesen werden.")
			return
		}
		if err := os.WriteFile(filepath.Join(dir, fmt.Sprintf("bild%d.png", i)), png, 0o600); err != nil {
			fehler("Ein Bild konnte nicht gespeichert werden.")
			return
		}
	}
	cfg, _ := json.Marshal(map[string]any{"sheet": req.Sheet, "kw": req.Keyword, "fh": req.FontPt, "itemW": req.ItemW, "overlays": req.Overlays})
	cfgPath := filepath.Join(dir, "auftrag.json")
	_ = os.WriteFile(cfgPath, cfg, 0o600)

	cmd := hidden("powershell.exe", "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-Command", excelSignScript)
	cmd.Env = append(os.Environ(), "PU_IN="+in, "PU_CFG="+cfgPath, "PU_DIR="+dir)
	var out bytes.Buffer
	cmd.Stdout, cmd.Stderr = &out, &out
	if err := cmd.Start(); err != nil {
		fehler("PowerShell konnte nicht gestartet werden.")
		return
	}
	done := make(chan error, 1)
	go func() { done <- cmd.Wait() }()
	var runErr error
	select {
	case runErr = <-done:
	case <-time.After(2 * time.Minute):
		_ = cmd.Process.Kill()
		fehler("Excel hat zu lange gebraucht. Ist vielleicht ein Excel-Fenster mit einer Rückfrage offen?")
		return
	}
	if runErr != nil {
		s := out.String()
		switch {
		case strings.Contains(s, "ANKER_FEHLT"):
			fehler("Das Suchwort „" + req.Keyword + "“ steht nicht im Excel-Blatt – bitte als PDF speichern.")
		case strings.Contains(s, "80040154") || strings.Contains(strings.ToLower(s), "class not registered") || strings.Contains(s, "Klasse nicht registriert"):
			fehler("Dafür muss Microsoft Excel auf diesem Computer installiert sein.")
		default:
			fehler("Excel konnte die Unterschrift nicht einfügen.")
		}
		return
	}
	res, err := os.ReadFile(in)
	if err != nil {
		fehler("Die fertige Datei konnte nicht gelesen werden.")
		return
	}
	w.Header().Set("Content-Type", "application/octet-stream")
	w.Header().Set("Cache-Control", "no-store")
	w.Write(res)
}
