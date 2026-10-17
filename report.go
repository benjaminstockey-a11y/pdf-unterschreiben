//go:build windows

package main

import (
	"bytes"
	"encoding/json"
	"io"
	"net/http"
	"os"
	"strings"
	"time"
)

// Nutzungsmeldung an die Admin-Tabelle (Google Apps Script). Adresse und Schlüssel kommen beim
// Bauen aus den GitHub-Secrets, damit sie nicht im öffentlichen Quelltext stehen.
var (
	ReportURL = ""
	ReportKey = ""
)

type reportUser struct {
	ID    string `json:"id"`
	Name  string `json:"name"`
	Email string `json:"email"`
}

func (a *app) handleReport(w http.ResponseWriter, r *http.Request) {
	if r.Method != http.MethodPost {
		http.Error(w, "POST", http.StatusMethodNotAllowed)
		return
	}
	var ev struct {
		Event string `json:"event"`
		Count int    `json:"count"`
		Code  string `json:"code"`
	}
	_ = json.NewDecoder(io.LimitReader(r.Body, 1<<16)).Decode(&ev)
	if ReportURL == "" || !strings.HasPrefix(ReportURL, "https://script.google.com/") {
		writeJSON(w, map[string]any{"ok": true, "disabled": true})
		return
	}
	// Den Nutzer aus den gespeicherten Einstellungen nehmen, nicht aus der Anfrage
	var s struct {
		User *reportUser `json:"user"`
	}
	if b, err := os.ReadFile(settingsPath()); err == nil {
		_ = json.Unmarshal(b, &s)
	}
	if s.User == nil || s.User.ID == "" {
		writeJSON(w, map[string]any{"ok": false, "error": "Nicht angemeldet."})
		return
	}
	if ev.Count <= 0 {
		ev.Count = 1
	}
	payload, _ := json.Marshal(map[string]any{
		"key": ReportKey, "id": s.User.ID, "name": s.User.Name, "email": s.User.Email,
		"event": ev.Event, "count": ev.Count, "code": ev.Code, "version": Version,
	})
	send := func() ([]byte, error) {
		c := &http.Client{Timeout: 30 * time.Second}
		resp, err := c.Post(ReportURL, "application/json", bytes.NewReader(payload))
		if err != nil {
			return nil, err
		}
		defer resp.Body.Close()
		return io.ReadAll(io.LimitReader(resp.Body, 1<<16))
	}
	switch ev.Event {
	case "save", "print", "mail":
		// Zählt nur mit – nicht auf die Antwort warten
		go send()
		writeJSON(w, map[string]any{"ok": true})
	default:
		// Anmeldung, Code prüfen, Start: Antwort wird gebraucht (Freischaltung / Sperre)
		b, err := send()
		if err != nil || !json.Valid(b) {
			writeJSON(w, map[string]any{"ok": false, "offline": true, "error": "Keine Verbindung zum Anmeldeserver. Bitte Internetverbindung prüfen."})
			return
		}
		w.Header().Set("Content-Type", "application/json; charset=utf-8")
		w.Header().Set("Cache-Control", "no-store")
		w.Write(b)
	}
}
