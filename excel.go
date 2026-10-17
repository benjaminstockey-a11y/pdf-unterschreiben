//go:build windows

package main

import (
	"archive/zip"
	"bytes"
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
