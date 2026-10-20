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
	"strings"
	"time"
)

// Texterkennung (z. B. Seriennummer auf dem Typenschild) mit der in Windows 10/11 eingebauten
// OCR (Windows.Media.Ocr) – läuft ohne Internet. Mehrere Bilder werden in einem Aufruf erkannt.
const ocrScript = `$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [Text.Encoding]::UTF8
Add-Type -AssemblyName System.Runtime.WindowsRuntime
$asTask = ([System.WindowsRuntimeSystemExtensions].GetMethods() | Where-Object {
  $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncOperation` + "`" + `1' })[0]
function Await($op, [Type]$t) { $task = $asTask.MakeGenericMethod($t).Invoke($null, @($op)); [void]$task.Wait(-1); $task.Result }
[void][Windows.Storage.StorageFile, Windows.Storage, ContentType = WindowsRuntime]
[void][Windows.Media.Ocr.OcrEngine, Windows.Foundation, ContentType = WindowsRuntime]
[void][Windows.Graphics.Imaging.BitmapDecoder, Windows.Graphics, ContentType = WindowsRuntime]
$engine = [Windows.Media.Ocr.OcrEngine]::TryCreateFromUserProfileLanguages()
if (-not $engine) { $engine = [Windows.Media.Ocr.OcrEngine]::TryCreateFromLanguage([Windows.Globalization.Language]::new('en-US')) }
if (-not $engine) { throw 'Keine Texterkennung in Windows verfügbar.' }
$out = @()
foreach ($p in ($env:PU_OCR -split '\|')) {
  if (-not $p) { continue }
  try {
    $file = Await ([Windows.Storage.StorageFile]::GetFileFromPathAsync($p)) ([Windows.Storage.StorageFile])
    $stream = Await ($file.OpenAsync([Windows.Storage.FileAccessMode]::Read)) ([Windows.Storage.Streams.IRandomAccessStream])
    $dec = Await ([Windows.Graphics.Imaging.BitmapDecoder]::CreateAsync($stream)) ([Windows.Graphics.Imaging.BitmapDecoder])
    $bmp = Await ($dec.GetSoftwareBitmapAsync()) ([Windows.Graphics.Imaging.SoftwareBitmap])
    $res = Await ($engine.RecognizeAsync($bmp)) ([Windows.Media.Ocr.OcrResult])
    $out += ,(@($res.Lines | ForEach-Object { $_.Text }) -join "` + "`" + `n")
    $stream.Dispose()
  } catch { $out += ,'' }
}
ConvertTo-Json -InputObject @($out) -Compress
`

func (a *app) handleOCR(w http.ResponseWriter, r *http.Request) {
	if r.Method != http.MethodPost {
		http.Error(w, "POST", http.StatusMethodNotAllowed)
		return
	}
	var req struct {
		Images []string `json:"images"` // JPEG als Base64
	}
	if err := json.NewDecoder(io.LimitReader(r.Body, 200<<20)).Decode(&req); err != nil || len(req.Images) == 0 {
		writeJSON(w, map[string]any{"ok": false, "error": "Keine Bilder."})
		return
	}
	imgs := make([][]byte, len(req.Images))
	for i, s := range req.Images {
		if b, err := base64.StdEncoding.DecodeString(s); err == nil {
			imgs[i] = b
		}
	}
	// Zuerst Google (genauer), nur was dort nicht klappt, mit Windows lesen
	gTexts, gOK := googleOCRAll(imgs)
	need := false
	for i := range imgs {
		if imgs[i] != nil && !gOK[i] {
			need = true
		}
	}
	if !need {
		writeJSON(w, map[string]any{"ok": true, "texts": gTexts, "engine": "google"})
		return
	}
	dir, err := os.MkdirTemp("", "pu-ocr-")
	if err != nil {
		writeJSON(w, map[string]any{"ok": false, "error": err.Error()})
		return
	}
	defer os.RemoveAll(dir)
	paths := make([]string, len(imgs))
	for i, b := range imgs {
		if b == nil || gOK[i] {
			continue
		}
		p := filepath.Join(dir, fmt.Sprintf("%03d.jpg", i))
		if os.WriteFile(p, b, 0o600) == nil {
			paths[i] = p
		}
	}
	var list []string
	for _, p := range paths {
		if p != "" {
			list = append(list, p)
		}
	}
	cmd := hidden("powershell.exe", "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-Command", ocrScript)
	cmd.Env = append(os.Environ(), "PU_OCR="+strings.Join(list, "|"))
	var out, errb bytes.Buffer
	cmd.Stdout, cmd.Stderr = &out, &errb
	if err := runWithTimeout(cmd, time.Duration(30+5*len(list))*time.Second); err != nil {
		if hasAny(gOK) {
			writeJSON(w, map[string]any{"ok": true, "texts": gTexts, "engine": "google"})
			return
		}
		writeJSON(w, map[string]any{"ok": false, "error": "Texterkennung nicht möglich: " + strings.TrimSpace(errb.String()+" "+err.Error())})
		return
	}
	// Ausgabe enthält evtl. BOM/Zeilenumbrüche – nur das JSON-Array nehmen
	s := out.String()
	if i := strings.Index(s, "["); i >= 0 {
		s = s[i:]
	}
	var texts []string
	if json.Unmarshal([]byte(strings.TrimSpace(s)), &texts) != nil {
		if hasAny(gOK) {
			writeJSON(w, map[string]any{"ok": true, "texts": gTexts, "engine": "google"})
			return
		}
		writeJSON(w, map[string]any{"ok": false, "error": "Texterkennung lieferte keine Antwort."})
		return
	}
	// Windows-Ergebnisse an die richtigen Stellen setzen, Google-Treffer behalten
	res := gTexts
	j := 0
	for i, p := range paths {
		if p != "" && j < len(texts) {
			res[i] = texts[j]
			j++
		}
	}
	writeJSON(w, map[string]any{"ok": true, "texts": res})
}

func hasAny(b []bool) bool {
	for _, x := range b {
		if x {
			return true
		}
	}
	return false
}
