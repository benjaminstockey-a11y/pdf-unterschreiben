//go:build windows

package main

import (
	"bytes"
	"encoding/base64"
	"encoding/json"
	"errors"
	"io"
	"net/http"
	"os"
	"strings"
	"sync"
	"time"
)

// Texterkennung über den Anmeldeserver (Google). Liest Typenschilder deutlich genauer als die
// Windows-Texterkennung, braucht aber Internet. Schlägt sie fehl, wird Windows verwendet.
func googleOCR(img []byte) (string, error) {
	if ReportURL == "" || !strings.HasPrefix(ReportURL, "https://script.google.com/") {
		return "", errors.New("kein Server")
	}
	if len(img) < 100 || len(img) > 10<<20 {
		return "", errors.New("Foto zu groß")
	}
	var s struct {
		User *reportUser `json:"user"`
	}
	if b, err := os.ReadFile(settingsPath()); err == nil {
		_ = json.Unmarshal(b, &s)
	}
	if s.User == nil || s.User.Email == "" {
		return "", errors.New("nicht angemeldet")
	}
	payload, _ := json.Marshal(map[string]any{
		"key": ReportKey, "event": "ocr", "id": s.User.ID, "email": s.User.Email, "version": Version,
		"b64": base64.StdEncoding.EncodeToString(img), "mime": http.DetectContentType(img),
	})
	c := &http.Client{Timeout: 60 * time.Second}
	resp, err := c.Post(ReportURL, "application/json", bytes.NewReader(payload))
	if err != nil {
		return "", err
	}
	defer resp.Body.Close()
	var res struct {
		OK    bool   `json:"ok"`
		Text  string `json:"text"`
		Error string `json:"error"`
	}
	if err := json.NewDecoder(io.LimitReader(resp.Body, 4<<20)).Decode(&res); err != nil {
		return "", errors.New("keine gültige Antwort")
	}
	if !res.OK {
		return "", errors.New(res.Error)
	}
	// Google setzt an den Anfang eine Trennlinie für das eingebettete Bild
	t := strings.ReplaceAll(res.Text, "\r", "")
	t = strings.TrimLeft(strings.TrimPrefix(strings.TrimSpace(t), "\uFEFF"), "_\n ")
	return t, nil
}

// Mehrere Fotos gleichzeitig mit Google lesen; ok[i] sagt, ob Bild i erkannt wurde.
func googleOCRAll(imgs [][]byte) (texts []string, ok []bool) {
	texts, ok = make([]string, len(imgs)), make([]bool, len(imgs))
	var wg sync.WaitGroup
	sem := make(chan struct{}, 6)
	for i, b := range imgs {
		if b == nil {
			continue
		}
		wg.Add(1)
		go func(i int, b []byte) {
			defer wg.Done()
			sem <- struct{}{}
			defer func() { <-sem }()
			if t, err := googleOCR(b); err == nil {
				texts[i], ok[i] = t, true
			}
		}(i, b)
	}
	wg.Wait()
	return
}
