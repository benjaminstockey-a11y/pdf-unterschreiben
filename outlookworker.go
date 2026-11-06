//go:build windows

package main

import (
	"bufio"
	"encoding/base64"
	"encoding/json"
	"errors"
	"io"
	"net/http"
	"os/exec"
	"strings"
	"sync"
	"time"
)

// Schnelleres Outlook: Eine PowerShell bleibt im Hintergrund offen und hält die Verbindung zu Outlook.
// So entfällt bei jeder E-Mail das Starten von PowerShell (auf Firmenrechnern mit Virenscanner langsam),
// und Outlook wird schon vorgestartet, sobald man auf das Versenden zugeht (Dokumentation, Berichte geladen).
const outlookWorkerScript = `$ErrorActionPreference = 'Stop'
$script:o = $null
function Ol {
  if ($script:o) { try { [void]$script:o.Version; return $script:o } catch { $script:o = $null } }
  if (-not (Get-Process outlook -ErrorAction SilentlyContinue)) {
    Start-Process 'outlook.exe' -WindowStyle Minimized
    for ($i = 0; $i -lt 120 -and -not $script:o; $i++) {
      Start-Sleep -Milliseconds 500
      if (Get-Process outlook -ErrorAction SilentlyContinue) { try { $script:o = New-Object -ComObject Outlook.Application; [void]$script:o.Version } catch { $script:o = $null } }
    }
  }
  if (-not $script:o) { $script:o = New-Object -ComObject Outlook.Application }
  return $script:o
}
while ($true) {
  $l = [Console]::In.ReadLine()
  if ($l -eq $null) { break }
  try {
    $sw = [Diagnostics.Stopwatch]::StartNew(); $t = @()
    $j = [Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($l)) | ConvertFrom-Json
    if ($j.art -eq 'leicht') { if (Get-Process outlook -ErrorAction SilentlyContinue) { [void](Ol) }; [Console]::Out.WriteLine('OK'); continue }
    $o = Ol; $t += 'Outlook ' + $sw.ElapsedMilliseconds
    if ($j.art -eq 'mail') {
      $m = $o.CreateItem(0); $t += 'Neu ' + $sw.ElapsedMilliseconds
      $m.Subject = [string]$j.subj
      if ($j.to) { $m.To = [string]$j.to }
      foreach ($f in $j.files) { if ($f) { [void]$m.Attachments.Add([string]$f) } }
      $t += 'Anhaenge ' + $sw.ElapsedMilliseconds
      $m.Display($false); $t += 'Anzeigen ' + $sw.ElapsedMilliseconds
      try { $m.GetInspector.Activate() } catch {}
    }
    [Console]::Out.WriteLine('OK ' + ($t -join ', '))
  } catch { [Console]::Out.WriteLine('FEHLER ' + ($_.Exception.Message -replace '\s+', ' ')) }
}`

type outlookWorker struct {
	mu   sync.Mutex
	cmd  *exec.Cmd
	in   io.WriteCloser
	out  *bufio.Reader
	tot  bool
	letz time.Time
	// Zeitmessung der letzten Aufträge (ms), zum Nachsehen unter /api/mail/zeiten
	zeiten string
}

var mailZeiten []string

func mailZeitMerken(z string) {
	mailZeiten = append(mailZeiten, time.Now().Format("15:04:05")+" "+z)
	if len(mailZeiten) > 20 {
		mailZeiten = mailZeiten[len(mailZeiten)-20:]
	}
}

func (a *app) handleMailZeiten(w http.ResponseWriter, r *http.Request) {
	writeJSON(w, map[string]any{"zeiten": mailZeiten})
}

// beim Programmstart: PowerShell schon starten und – nur falls Outlook läuft – verbinden
func outlookLeichtVor() {
	go func() { time.Sleep(3 * time.Second); _ = olw.auftrag(map[string]any{"art": "leicht"}, 60*time.Second) }()
}

var olw = &outlookWorker{}

func (w *outlookWorker) starten() error {
	if w.cmd != nil && !w.tot {
		return nil
	}
	cmd := hidden("powershell.exe", "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-Command", outlookWorkerScript)
	in, err := cmd.StdinPipe()
	if err != nil {
		return err
	}
	out, err := cmd.StdoutPipe()
	if err != nil {
		return err
	}
	if err := cmd.Start(); err != nil {
		return err
	}
	w.cmd, w.in, w.out, w.tot = cmd, in, bufio.NewReader(out), false
	go func(c *exec.Cmd) {
		_ = c.Wait()
		w.mu.Lock()
		if w.cmd == c {
			w.tot = true
		}
		w.mu.Unlock()
	}(cmd)
	return nil
}

// auftrag schickt einen Auftrag an die Hintergrund-PowerShell und wartet auf die Antwort.
func (w *outlookWorker) auftrag(a map[string]any, timeout time.Duration) error {
	w.mu.Lock()
	defer w.mu.Unlock()
	if err := w.starten(); err != nil {
		return err
	}
	b, _ := json.Marshal(a)
	if _, err := io.WriteString(w.in, base64.StdEncoding.EncodeToString(b)+"\n"); err != nil {
		w.tot = true
		return err
	}
	ch := make(chan string, 1)
	go func(r *bufio.Reader) { s, _ := r.ReadString('\n'); ch <- strings.TrimSpace(s) }(w.out)
	select {
	case s := <-ch:
		w.letz = time.Now()
		if s == "OK" || strings.HasPrefix(s, "OK ") {
			w.zeiten = strings.TrimSpace(strings.TrimPrefix(s, "OK"))
			return nil
		}
		if s == "" {
			w.tot = true
			return errors.New("PowerShell beendet")
		}
		return errors.New(strings.TrimPrefix(s, "FEHLER "))
	case <-time.After(timeout):
		// hängt – beim nächsten Mal neu starten
		_ = w.cmd.Process.Kill()
		w.tot = true
		return errors.New("Zeitüberschreitung")
	}
}

// /api/mail/vor – Outlook schon im Hintergrund bereit machen (läuft asynchron)
func (a *app) handleMailVor(w http.ResponseWriter, r *http.Request) {
	go func() {
		olw.mu.Lock()
		frisch := olw.cmd != nil && !olw.tot && time.Since(olw.letz) < 2*time.Minute
		olw.mu.Unlock()
		if !frisch {
			_ = olw.auftrag(map[string]any{"art": "vor"}, 90*time.Second)
		}
	}()
	w.WriteHeader(http.StatusNoContent)
}
