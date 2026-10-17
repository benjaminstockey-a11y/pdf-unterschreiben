//go:build windows

package main

import (
	"bytes"
	"encoding/base64"
	"encoding/json"
	"fmt"
	"io"
	"net/http"
	"os"
	"path/filepath"
	"regexp"
	"strings"
	"syscall"
	"time"
	"unsafe"
)

// Berichte per E-Mail: neue Nachricht mit den unterschriebenen PDFs als Anhang öffnen.
// 1. klassisches Outlook (COM)  2. Standard-Mailprogramm über MAPI  3. Ordner im Explorer zeigen.

type mailReq struct {
	Subject string `json:"subject"`
	Files   []struct {
		Name string `json:"name"`
		Data string `json:"data"`
	} `json:"files"`
}

var badChars = regexp.MustCompile(`[<>:"/\\|?*\x00-\x1f]`)

func versandDir() string {
	base := filepath.Join(os.Getenv("LOCALAPPDATA"), AppName, "Versand")
	// alte Versandordner aufräumen
	if entries, err := os.ReadDir(base); err == nil {
		for _, e := range entries {
			if info, err := e.Info(); err == nil && time.Since(info.ModTime()) > 7*24*time.Hour {
				_ = os.RemoveAll(filepath.Join(base, e.Name()))
			}
		}
	}
	d := filepath.Join(base, time.Now().Format("2006-01-02_150405"))
	_ = os.MkdirAll(d, 0o755)
	return d
}

const outlookScript = `$ErrorActionPreference = 'Stop'
$o = New-Object -ComObject Outlook.Application
$m = $o.CreateItem(0)
$m.Subject = $env:PU_SUBJ
foreach ($f in ($env:PU_FILES -split '\|')) { if ($f) { [void]$m.Attachments.Add($f) } }
$m.Display($false)
try { $m.GetInspector.Activate() } catch {}`

func (a *app) handleMail(w http.ResponseWriter, r *http.Request) {
	if r.Method != http.MethodPost {
		http.Error(w, "POST", http.StatusMethodNotAllowed)
		return
	}
	var req mailReq
	if err := json.NewDecoder(io.LimitReader(r.Body, 200<<20)).Decode(&req); err != nil || len(req.Files) == 0 {
		writeJSON(w, map[string]any{"ok": false, "error": "Keine Berichte zum Versenden."})
		return
	}
	dir := versandDir()
	var paths []string
	for _, f := range req.Files {
		b, err := base64.StdEncoding.DecodeString(f.Data)
		if err != nil {
			continue
		}
		name := strings.TrimSpace(badChars.ReplaceAllString(f.Name, "_"))
		if name == "" {
			name = "Bericht.pdf"
		}
		p := filepath.Join(dir, name)
		if os.WriteFile(p, b, 0o644) == nil {
			paths = append(paths, p)
		}
	}
	if len(paths) == 0 {
		writeJSON(w, map[string]any{"ok": false, "error": "Die Anhänge konnten nicht vorbereitet werden."})
		return
	}

	cmd := hidden("powershell.exe", "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-Command", outlookScript)
	cmd.Env = append(os.Environ(), "PU_SUBJ="+req.Subject, "PU_FILES="+strings.Join(paths, "|"))
	var out bytes.Buffer
	cmd.Stdout, cmd.Stderr = &out, &out
	if err := runWithTimeout(cmd, 60*time.Second); err == nil {
		writeJSON(w, map[string]any{"ok": true, "via": "outlook"})
		return
	}
	// Kein klassisches Outlook – Standard-Mailprogramm über MAPI probieren (blockiert, bis die Mail zu ist).
	mapiErr := make(chan error, 1)
	go func() { mapiErr <- mapiSendMail(req.Subject, paths) }()
	select {
	case err := <-mapiErr:
		if err == nil {
			writeJSON(w, map[string]any{"ok": true, "via": "mapi"})
			return
		}
	case <-time.After(4 * time.Second):
		// Fenster ist offen und wartet auf den Nutzer
		writeJSON(w, map[string]any{"ok": true, "via": "mapi"})
		return
	}
	// Letzte Möglichkeit: Ordner mit den Dateien zeigen
	_ = hidden("explorer.exe", "/select,", paths[0]).Start()
	writeJSON(w, map[string]any{"ok": true, "via": "explorer", "dir": dir})
}

func runWithTimeout(cmd interface {
	Start() error
	Wait() error
}, d time.Duration) error {
	if err := cmd.Start(); err != nil {
		return err
	}
	done := make(chan error, 1)
	go func() { done <- cmd.Wait() }()
	select {
	case err := <-done:
		return err
	case <-time.After(d):
		return fmt.Errorf("Zeitüberschreitung")
	}
}

/* ---------- Simple MAPI ---------- */

type mapiFileDescW struct {
	Reserved uint32
	Flags    uint32
	Position uint32
	PathName *uint16
	FileName *uint16
	FileType uintptr
}

type mapiMessageW struct {
	Reserved       uint32
	Subject        *uint16
	NoteText       *uint16
	MessageType    *uint16
	DateReceived   *uint16
	ConversationID *uint16
	Flags          uint32
	Originator     uintptr
	RecipCount     uint32
	Recips         uintptr
	FileCount      uint32
	Files          *mapiFileDescW
}

func mapiSendMail(subject string, paths []string) error {
	dll := syscall.NewLazyDLL("mapi32.dll")
	proc := dll.NewProc("MAPISendMailW")
	if err := proc.Find(); err != nil {
		return err
	}
	files := make([]mapiFileDescW, len(paths))
	for i, p := range paths {
		pp, _ := syscall.UTF16PtrFromString(p)
		fn, _ := syscall.UTF16PtrFromString(filepath.Base(p))
		files[i] = mapiFileDescW{Position: 0xFFFFFFFF, PathName: pp, FileName: fn}
	}
	subj, _ := syscall.UTF16PtrFromString(subject)
	empty, _ := syscall.UTF16PtrFromString("")
	msg := mapiMessageW{Subject: subj, NoteText: empty, FileCount: uint32(len(files)), Files: &files[0]}
	const mapiLogonUI, mapiDialog = 0x1, 0x8
	r, _, _ := proc.Call(0, 0, uintptr(unsafe.Pointer(&msg)), mapiLogonUI|mapiDialog, 0)
	if r != 0 && r != 1 { // 1 = vom Nutzer abgebrochen
		return fmt.Errorf("MAPI-Fehler %d", r)
	}
	return nil
}
