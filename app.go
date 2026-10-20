//go:build windows

package main

import (
	"bytes"
	"context"
	"crypto/sha256"
	"embed"
	"encoding/hex"
	"encoding/json"
	"errors"
	"fmt"
	"io"
	"io/fs"
	"net/http"
	"os"
	"path/filepath"
	"strconv"
	"strings"
	"sync"
	"time"
)

//go:embed all:web
var webFS embed.FS

const pdfLibVersion = "1.17.1"

var pdfLibMirrors = []string{
	"https://cdnjs.cloudflare.com/ajax/libs/pdf-lib/" + pdfLibVersion + "/pdf-lib.min.js",
	"https://cdn.jsdelivr.net/npm/pdf-lib@" + pdfLibVersion + "/dist/pdf-lib.min.js",
	"https://unpkg.com/pdf-lib@" + pdfLibVersion + "/dist/pdf-lib.min.js",
}

type app struct {
	srv      *http.Server
	mu       sync.Mutex
	clients  int
	hadOne   bool
	lastSeen time.Time
	started  time.Time
	quit     chan struct{}
	once     sync.Once
}

func url() string { return fmt.Sprintf("http://127.0.0.1:%d/", Port) }

func runApp(relaunch bool) {
	cleanupOld()
	_ = os.WriteFile(iconPath(), appIcon, 0o644)
	go func() { refreshIcon(); migrateName() }()

	ln, mine := listenOwn(relaunch)
	if ln == nil {
		if mine {
			// Läuft für diesen Windows-Benutzer schon – einfach das Fenster öffnen.
			openWindow(url())
			return
		}
		fail("Das Programm konnte nicht gestartet werden (kein freier Anschluss).")
		return
	}
	if relaunch {
		registerUninstall() // neue Versionsnummer in „Apps & Features“
	}

	a := &app{started: time.Now(), lastSeen: time.Now(), quit: make(chan struct{})}
	a.srv = &http.Server{Handler: a.routes(), ReadHeaderTimeout: 10 * time.Second}
	go a.srv.Serve(ln)
	if relaunch && Port != legacyPort {
		go legacyBridge()
	}
	if !relaunch {
		openWindow(url())
	}
	a.watchLifetime()
}

func (a *app) stop() {
	a.once.Do(func() { close(a.quit) })
}

// watchLifetime beendet das Programm, sobald kein Fenster mehr verbunden ist.
func (a *app) watchLifetime() {
	t := time.NewTicker(time.Second)
	defer t.Stop()
	for {
		select {
		case <-a.quit:
			a.shutdown()
			return
		case <-t.C:
			a.mu.Lock()
			idle := a.clients == 0 && time.Since(a.lastSeen) > 10*time.Second && a.hadOne
			never := !a.hadOne && time.Since(a.started) > 2*time.Minute
			a.mu.Unlock()
			if idle || never {
				a.shutdown()
				return
			}
		}
	}
}

func (a *app) shutdown() {
	ctx, cancel := context.WithTimeout(context.Background(), 2*time.Second)
	defer cancel()
	_ = a.srv.Shutdown(ctx)
}

func (a *app) routes() http.Handler {
	mux := http.NewServeMux()
	sub, _ := fs.Sub(webFS, "web")
	static := http.FileServer(http.FS(sub))

	mux.HandleFunc("/", func(w http.ResponseWriter, r *http.Request) {
		w.Header().Set("Cache-Control", "no-store")
		static.ServeHTTP(w, r)
	})
	mux.HandleFunc("/lib/pdf-lib.min.js", a.handlePdfLib)
	mux.HandleFunc("/api/owner", func(w http.ResponseWriter, r *http.Request) {
		w.Header().Set("Cache-Control", "no-store")
		w.Write([]byte(ownerToken()))
	})
	mux.HandleFunc("/api/info", func(w http.ResponseWriter, r *http.Request) {
		writeJSON(w, map[string]any{"version": Version, "repo": repoSlug(), "updates": RepoOwner != "", "report": ReportURL != ""})
	})
	mux.HandleFunc("/api/alive", a.handleAlive)
	mux.HandleFunc("/api/settings", a.handleSettings)
	mux.HandleFunc("/api/update/check", a.handleUpdateCheck)
	mux.HandleFunc("/api/update/apply", a.handleUpdateApply)
	mux.HandleFunc("/api/excel/sheets", a.handleExcelSheets)
	mux.HandleFunc("/api/excel/pdf", a.handleExcelPDF)
	mux.HandleFunc("/api/excel/sign", a.handleExcelSign)
	mux.HandleFunc("/api/mail", a.handleMail)
	mux.HandleFunc("/api/ocr", a.handleOCR)
	mux.HandleFunc("/api/report", a.handleReport)
	mux.HandleFunc("/api/open", func(w http.ResponseWriter, r *http.Request) {
		// Nur feste Adressen im Standardbrowser öffnen (dort ist man bei Google angemeldet)
		links := map[string]string{
			"tabelle": "https://docs.google.com/spreadsheets/d/1i7jifCaffYGLXdDVSidWPqHbZTkaWhAF6_wVwN3H0NA/edit",
		}
		u, ok := links[r.URL.Query().Get("ziel")]
		if r.Method != http.MethodPost || !ok {
			http.Error(w, "nein", http.StatusBadRequest)
			return
		}
		_ = hidden("rundll32.exe", "url.dll,FileProtocolHandler", u).Start()
		w.WriteHeader(http.StatusNoContent)
	})
	mux.HandleFunc("/api/quit", func(w http.ResponseWriter, r *http.Request) {
		if r.Method != http.MethodPost {
			http.Error(w, "POST", http.StatusMethodNotAllowed)
			return
		}
		w.WriteHeader(http.StatusNoContent)
		go func() { time.Sleep(200 * time.Millisecond); a.stop() }()
	})
	return guard(mux)
}

// guard lässt nur Anfragen von dieser Seite selbst zu (Schutz gegen fremde Webseiten im Browser).
func guard(next http.Handler) http.Handler {
	return http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		host := r.Host
		if host != fmt.Sprintf("127.0.0.1:%d", Port) && host != fmt.Sprintf("localhost:%d", Port) {
			http.Error(w, "forbidden", http.StatusForbidden)
			return
		}
		if o := r.Header.Get("Origin"); o != "" && o != strings.TrimSuffix(url(), "/") && o != fmt.Sprintf("http://localhost:%d", Port) {
			http.Error(w, "forbidden", http.StatusForbidden)
			return
		}
		next.ServeHTTP(w, r)
	})
}

func writeJSON(w http.ResponseWriter, v any) {
	w.Header().Set("Content-Type", "application/json; charset=utf-8")
	w.Header().Set("Cache-Control", "no-store")
	_ = json.NewEncoder(w).Encode(v)
}

func (a *app) handleAlive(w http.ResponseWriter, r *http.Request) {
	fl, ok := w.(http.Flusher)
	if !ok {
		http.Error(w, "stream", http.StatusInternalServerError)
		return
	}
	w.Header().Set("Content-Type", "text/event-stream")
	w.Header().Set("Cache-Control", "no-store")
	a.mu.Lock()
	a.clients++
	a.hadOne = true
	a.mu.Unlock()
	defer func() {
		a.mu.Lock()
		a.clients--
		a.lastSeen = time.Now()
		a.mu.Unlock()
	}()
	fmt.Fprint(w, ": ok\n\n")
	fl.Flush()
	t := time.NewTicker(15 * time.Second)
	defer t.Stop()
	for {
		select {
		case <-r.Context().Done():
			return
		case <-a.quit:
			return
		case <-t.C:
			if _, err := fmt.Fprint(w, ": ping\n\n"); err != nil {
				return
			}
			fl.Flush()
		}
	}
}

func settingsPath() string { return filepath.Join(dataDir(), "einstellungen.json") }

func (a *app) handleSettings(w http.ResponseWriter, r *http.Request) {
	switch r.Method {
	case http.MethodGet:
		b, err := os.ReadFile(settingsPath())
		if err != nil || !json.Valid(b) {
			b = []byte("{}")
		}
		w.Header().Set("Content-Type", "application/json; charset=utf-8")
		w.Header().Set("Cache-Control", "no-store")
		w.Write(b)
	case http.MethodPut:
		b, err := io.ReadAll(io.LimitReader(r.Body, 8<<20))
		if err != nil || !json.Valid(b) {
			http.Error(w, "ungültig", http.StatusBadRequest)
			return
		}
		tmp := settingsPath() + ".tmp"
		if err := os.WriteFile(tmp, b, 0o600); err != nil {
			http.Error(w, err.Error(), http.StatusInternalServerError)
			return
		}
		_ = os.Remove(settingsPath())
		if err := os.Rename(tmp, settingsPath()); err != nil {
			http.Error(w, err.Error(), http.StatusInternalServerError)
			return
		}
		w.WriteHeader(http.StatusNoContent)
	default:
		http.Error(w, "GET/PUT", http.StatusMethodNotAllowed)
	}
}

// pdf-lib wird mitgeliefert, wenn es beim Bauen vorlag – sonst einmalig geladen und zwischengespeichert.
func (a *app) handlePdfLib(w http.ResponseWriter, r *http.Request) {
	w.Header().Set("Content-Type", "text/javascript; charset=utf-8")
	if b, err := webFS.ReadFile("web/lib/pdf-lib.min.js"); err == nil {
		w.Write(b)
		return
	}
	cache := filepath.Join(os.Getenv("LOCALAPPDATA"), AppName, "lib", "pdf-lib-"+pdfLibVersion+".min.js")
	if b, err := os.ReadFile(cache); err == nil && len(b) > 100000 {
		w.Write(b)
		return
	}
	c := &http.Client{Timeout: 30 * time.Second}
	for _, u := range pdfLibMirrors {
		resp, err := c.Get(u)
		if err != nil {
			continue
		}
		b, err := io.ReadAll(io.LimitReader(resp.Body, 5<<20))
		resp.Body.Close()
		if err != nil || resp.StatusCode != 200 || len(b) < 100000 {
			continue
		}
		_ = os.MkdirAll(filepath.Dir(cache), 0o755)
		_ = os.WriteFile(cache, b, 0o644)
		w.Write(b)
		return
	}
	w.WriteHeader(http.StatusServiceUnavailable)
	fmt.Fprint(w, `window.PDFLIB_FEHLT = true;`)
}

/* ---------------- Updates ---------------- */

type release struct {
	Version string `json:"version"`
	Notes   string `json:"notes"`
	URL     string `json:"url"`
	SHA256  string `json:"sha256"`
}

func repoSlug() string {
	if RepoOwner == "" {
		return ""
	}
	return RepoOwner + "/" + RepoName
}

// Stabile Versionen liegen im neuesten Release, Testversionen im Vorab-Release „adminkanal“.
func fetchChannel(test bool) (*release, error) {
	if RepoOwner == "" {
		return nil, errors.New("keine Update-Quelle eingestellt")
	}
	u := fmt.Sprintf("https://github.com/%s/releases/latest/download/version.json", repoSlug())
	if test {
		u = fmt.Sprintf("https://github.com/%s/releases/download/adminkanal/version.json", repoSlug())
	}
	c := &http.Client{Timeout: 8 * time.Second}
	resp, err := c.Get(u)
	if err != nil {
		return nil, err
	}
	defer resp.Body.Close()
	if resp.StatusCode != 200 {
		return nil, fmt.Errorf("GitHub antwortet mit %d", resp.StatusCode)
	}
	var rel release
	if err := json.NewDecoder(io.LimitReader(resp.Body, 1<<20)).Decode(&rel); err != nil {
		return nil, err
	}
	return &rel, nil
}

// fetchLatest liefert die neueste passende Version; im Testkanal die neuere aus Test und Stabil.
func fetchLatest(test bool) (*release, error) {
	stable, err := fetchChannel(false)
	if !test {
		return stable, err
	}
	t, terr := fetchChannel(true)
	switch {
	case terr != nil:
		return stable, err
	case err != nil || newer(t.Version, stable.Version):
		return t, nil
	default:
		return stable, nil
	}
}

// Versionen wie 1.13.0, 1.13.0-test.4: die Testnummer zählt innerhalb derselben Version,
// und die fertige Version ist neuer als jede Testversion davon.
func parseVer(s string) (v [3]int, test int) {
	s = strings.TrimPrefix(strings.TrimSpace(s), "v")
	test = -1
	if i := strings.Index(s, "-"); i >= 0 {
		suf := s[i+1:]
		s = s[:i]
		if j := strings.LastIndex(suf, "."); j >= 0 {
			test, _ = strconv.Atoi(suf[j+1:])
		} else {
			test = 0
		}
	}
	if i := strings.Index(s, "+"); i >= 0 {
		s = s[:i]
	}
	for i, p := range strings.SplitN(s, ".", 3) {
		v[i], _ = strconv.Atoi(p)
	}
	return
}

func newer(latest, current string) bool {
	l, lt := parseVer(latest)
	c, ct := parseVer(current)
	for i := 0; i < 3; i++ {
		if l[i] != c[i] {
			return l[i] > c[i]
		}
	}
	if lt == ct {
		return false
	}
	if lt == -1 {
		return true // fertige Version schlägt Testversion
	}
	if ct == -1 {
		return false
	}
	return lt > ct
}

func (a *app) handleUpdateCheck(w http.ResponseWriter, r *http.Request) {
	rel, err := fetchLatest(r.URL.Query().Get("kanal") == "test")
	if err != nil {
		writeJSON(w, map[string]any{"ok": false, "current": Version, "error": err.Error(), "configured": RepoOwner != ""})
		return
	}
	writeJSON(w, map[string]any{"ok": true, "current": Version, "latest": rel.Version, "notes": rel.Notes, "available": newer(rel.Version, Version)})
}

func (a *app) handleUpdateApply(w http.ResponseWriter, r *http.Request) {
	if r.Method != http.MethodPost {
		http.Error(w, "POST", http.StatusMethodNotAllowed)
		return
	}
	fehler := func(msg string) { writeJSON(w, map[string]any{"ok": false, "error": msg}) }
	rel, err := fetchLatest(r.URL.Query().Get("kanal") == "test")
	if err != nil {
		fehler("Update-Info nicht erreichbar: " + err.Error())
		return
	}
	if !newer(rel.Version, Version) {
		fehler("Du hast schon die neueste Version.")
		return
	}
	if !strings.HasPrefix(rel.URL, "https://github.com/"+repoSlug()+"/") {
		fehler("Die Update-Adresse ist ungültig.")
		return
	}
	c := &http.Client{Timeout: 5 * time.Minute}
	resp, err := c.Get(rel.URL)
	if err != nil {
		fehler("Download fehlgeschlagen: " + err.Error())
		return
	}
	defer resp.Body.Close()
	if resp.StatusCode != 200 {
		fehler(fmt.Sprintf("Download fehlgeschlagen (%d).", resp.StatusCode))
		return
	}
	b, err := io.ReadAll(io.LimitReader(resp.Body, 200<<20))
	if err != nil {
		fehler("Download abgebrochen: " + err.Error())
		return
	}
	sum := sha256.Sum256(b)
	if rel.SHA256 != "" && !strings.EqualFold(hex.EncodeToString(sum[:]), rel.SHA256) {
		fehler("Die heruntergeladene Datei ist beschädigt. Bitte später nochmal versuchen.")
		return
	}
	if !bytes.HasPrefix(b, []byte("MZ")) {
		fehler("Die heruntergeladene Datei ist kein Windows-Programm.")
		return
	}
	tmp := installedExe() + ".tmp"
	if err := os.WriteFile(tmp, b, 0o755); err != nil {
		fehler("Konnte nicht speichern: " + err.Error())
		return
	}
	if err := replaceFile(tmp, installedExe()); err != nil {
		os.Remove(tmp)
		fehler("Konnte die alte Version nicht ersetzen: " + err.Error())
		return
	}
	if err := detached(installedExe(), "--relaunch"); err != nil {
		fehler("Neue Version installiert, aber der Neustart ging nicht. Bitte das Programm neu öffnen.")
		return
	}
	writeJSON(w, map[string]any{"ok": true, "version": rel.Version})
	go func() { time.Sleep(300 * time.Millisecond); a.stop() }()
}
