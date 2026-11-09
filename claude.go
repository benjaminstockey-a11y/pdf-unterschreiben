//go:build windows

package main

import (
	"encoding/base64"
	"encoding/json"
	"net/http"
	"os"
	"path/filepath"
	"strings"
	"time"
)

// Auftrag an Claude (nur Testversion/Admin): Text und Dateien gehen per Outlook an das Postfach,
// das eine geplante Claude-Aufgabe werktags stündlich abarbeitet (Änderung einbauen → Testversion).
const claudePostfach = "benjaminstockey@gmail.com"
const claudeBetreff = "[Service-Mappe Auftrag]"

func (a *app) handleClaude(w http.ResponseWriter, r *http.Request) {
	if r.Method != http.MethodPost {
		http.Error(w, "POST", http.StatusMethodNotAllowed)
		return
	}
	var in struct {
		Text    string `json:"text"`
		Von     string `json:"von"`
		Art     string `json:"art"` // "freigabe": getestete Testversion für alle freigeben
		Dateien []struct {
			Name string `json:"name"`
			B64  string `json:"b64"`
		} `json:"dateien"`
	}
	if err := json.NewDecoder(http.MaxBytesReader(w, r.Body, 40<<20)).Decode(&in); err != nil {
		writeJSON(w, map[string]any{"ok": false, "error": "Zu groß oder ungültig (höchstens etwa 25 MB)."})
		return
	}
	in.Text = strings.TrimSpace(in.Text)
	if in.Art == "freigabe" {
		if !strings.Contains(Version, "-test") {
			writeJSON(w, map[string]any{"ok": false, "error": "Freigeben geht nur aus einer Testversion."})
			return
		}
		nr := strings.SplitN(Version, "-test", 2)[0]
		body := "Bitte Testversion " + Version + " als Version " + nr + " für alle freigeben.\n\n---\nGesendet aus der Service-Mappe " + Version + " am " + time.Now().Format("02.01.2006 15:04") + " von " + in.Von
		if err := olNativAuftrag(olAuftrag{art: "senden", subj: "[Service-Mappe Freigabe] " + nr, to: claudePostfach, body: body}, 90*time.Second); err != nil {
			writeJSON(w, map[string]any{"ok": false, "error": "Outlook konnte die Nachricht nicht senden: " + err.Error()})
			return
		}
		writeJSON(w, map[string]any{"ok": true, "version": nr})
		return
	}
	if len(in.Text) < 3 {
		writeJSON(w, map[string]any{"ok": false, "error": "Bitte beschreiben, was geändert werden soll."})
		return
	}
	dir := versandDir()
	var files []string
	for _, d := range in.Dateien {
		b, err := base64.StdEncoding.DecodeString(d.B64)
		if err != nil {
			continue
		}
		name := ablageName.ReplaceAllString(filepath.Base(d.Name), "_")
		if name == "" || name == "." {
			name = "Datei"
		}
		p := filepath.Join(dir, name)
		if os.WriteFile(p, b, 0o644) == nil {
			files = append(files, p)
		}
	}
	kurz := strings.Join(strings.Fields(in.Text), " ")
	if r := []rune(kurz); len(r) > 60 {
		kurz = string(r[:60]) + " …"
	}
	body := in.Text + "\n\n---\nGesendet aus der Service-Mappe " + Version + " am " + time.Now().Format("02.01.2006 15:04") + " von " + in.Von
	err := olNativAuftrag(olAuftrag{art: "senden", subj: claudeBetreff + " " + kurz, to: claudePostfach, body: body, files: files}, 90*time.Second)
	if err != nil {
		writeJSON(w, map[string]any{"ok": false, "error": "Outlook konnte die Nachricht nicht senden: " + err.Error()})
		return
	}
	writeJSON(w, map[string]any{"ok": true})
}
