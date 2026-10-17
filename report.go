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
	}
	_ = json.NewDecoder(io.LimitReader(r.Body, 1<<16)).Decode(&ev)
	w.WriteHeader(http.StatusNoContent)
	if ReportURL == "" || !strings.HasPrefix(ReportURL, "https://script.google.com/") {
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
		return
	}
	if ev.Count <= 0 {
		ev.Count = 1
	}
	payload, _ := json.Marshal(map[string]any{
		"key": ReportKey, "id": s.User.ID, "name": s.User.Name, "email": s.User.Email,
		"event": ev.Event, "count": ev.Count, "version": Version,
	})
	go func() {
		c := &http.Client{Timeout: 20 * time.Second}
		resp, err := c.Post(ReportURL, "application/json", bytes.NewReader(payload))
		if err == nil {
			io.Copy(io.Discard, resp.Body)
			resp.Body.Close()
		}
	}()
}
