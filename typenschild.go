//go:build windows

package main

import (
	"bytes"
	_ "embed"
	"encoding/json"
	"fmt"
	"io"
	"net/http"
	"os"
	"path/filepath"
	"strings"
	"time"
)

// Typenschild-Erkennung für die Excel-Datei „Allgemeine Vorlagen“.
// Früher startete Excel selbst PowerShell – Virenscanner wie Sophos beenden Excel dabei.
// Jetzt schickt Excel nur den Pfad des Fotos hierher, und dieses Programm führt die Erkennung aus.

//go:embed assets/TypenschildOCR.ps1
var typenschildScript []byte

// Port in eine Datei schreiben, damit die Excel-Vorlage das laufende Programm findet.
func writePortFile() {
	_ = os.WriteFile(filepath.Join(dataDir(), "port.txt"), []byte(fmt.Sprintf("%d", Port)), 0o644)
}

func (a *app) handleTypenschild(w http.ResponseWriter, r *http.Request) {
	if r.Method != http.MethodPost {
		http.Error(w, "POST", http.StatusMethodNotAllowed)
		return
	}
	var req struct {
		Image string `json:"image"`
	}
	if err := json.NewDecoder(io.LimitReader(r.Body, 1<<16)).Decode(&req); err != nil {
		http.Error(w, "Ungültige Anfrage.", http.StatusBadRequest)
		return
	}
	img := filepath.Clean(strings.TrimSpace(req.Image))
	ext := strings.ToLower(filepath.Ext(img))
	if ext != ".jpg" && ext != ".jpeg" && ext != ".png" && ext != ".bmp" {
		http.Error(w, "Nur Fotos (JPG, PNG, BMP).", http.StatusBadRequest)
		return
	}
	if st, err := os.Stat(img); err != nil || st.IsDir() {
		http.Error(w, "Foto nicht gefunden.", http.StatusBadRequest)
		return
	}
	dir, err := os.MkdirTemp("", "pu-typenschild-")
	if err != nil {
		http.Error(w, err.Error(), http.StatusInternalServerError)
		return
	}
	defer os.RemoveAll(dir)
	ps1 := filepath.Join(dir, "TypenschildOCR.ps1")
	out := filepath.Join(dir, "ergebnis.txt")
	// UTF-8 mit BOM, damit Windows PowerShell 5.1 Umlaute im Skript richtig liest
	if err := os.WriteFile(ps1, append([]byte{0xEF, 0xBB, 0xBF}, bytes.TrimPrefix(typenschildScript, []byte{0xEF, 0xBB, 0xBF})...), 0o600); err != nil {
		http.Error(w, err.Error(), http.StatusInternalServerError)
		return
	}
	args := []string{"-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-File", ps1, "-Image", img, "-Out", out}
	// Zuerst Google lesen lassen (genauer); das Skript wertet dann nur noch den Text aus.
	// Ohne Internet oder bei einem Fehler liest Windows das Foto wie bisher selbst.
	if data, err := os.ReadFile(img); err == nil {
		if txt, err := googleOCR(data); err == nil && len(strings.TrimSpace(txt)) >= 8 {
			tf := filepath.Join(dir, "google.txt")
			if os.WriteFile(tf, []byte(txt), 0o600) == nil {
				args = append(args, "-TextFile", tf)
			}
		}
	}
	cmd := hidden("powershell.exe", args...)
	var errb bytes.Buffer
	cmd.Stdout, cmd.Stderr = &errb, &errb
	runErr := runWithTimeout(cmd, 170*time.Second)
	b, err := os.ReadFile(out)
	if err != nil {
		msg := "Die Texterkennung hat kein Ergebnis geliefert."
		if lg, e := os.ReadFile(out + ".log"); e == nil {
			msg += "\n" + strings.TrimSpace(string(lg))
		} else if runErr != nil {
			msg += " " + runErr.Error()
		}
		http.Error(w, msg, http.StatusInternalServerError)
		return
	}
	b = bytes.TrimPrefix(b, []byte{0xEF, 0xBB, 0xBF})
	w.Header().Set("Content-Type", "text/plain; charset=utf-8")
	w.Header().Set("Cache-Control", "no-store")
	w.Write(b)
}
