//go:build windows

// PDF Unterschreiben – ein einziges Programm für Setup, App und Updates.
//
// Wird die exe von irgendwo gestartet (z. B. aus dem Download-Ordner), arbeitet sie als Setup
// und installiert sich nach %LOCALAPPDATA%\Programs\PDF Unterschreiben. Aus dem Installationsordner
// gestartet, öffnet sie die Oberfläche in einem eigenen Edge-Fenster und prüft auf GitHub nach Updates.
package main

import (
	"os"
	"path/filepath"
	"strings"
)

// Werden beim Bauen per -ldflags gesetzt.
var (
	Version   = "0.0.0-dev"
	RepoOwner = ""
	RepoName  = "pdfunterschreiben"
)

const (
	AppName = "PDF Unterschreiben" // interner Name: Installations- und Datenordner (nicht ändern)
	// DisplayName ist der sichtbare Programmname (Fenster, Verknüpfungen, Apps & Features).
	DisplayName = "PDF Unterschreiben & Dokumentation"
	AppID       = "PDFUnterschreiben"
	ExeName     = "PDF Unterschreiben.exe"
	Port        = 47613 // fest, damit gespeicherte Einstellungen im Fenster erhalten bleiben
	Publisher   = DisplayName
)

func installDir() string {
	return filepath.Join(os.Getenv("LOCALAPPDATA"), "Programs", AppName)
}
func installedExe() string { return filepath.Join(installDir(), ExeName) }
func dataDir() string {
	d := filepath.Join(os.Getenv("APPDATA"), AppName)
	_ = os.MkdirAll(d, 0o755)
	return d
}

func main() {
	args := os.Args[1:]
	has := func(f string) bool {
		for _, a := range args {
			if a == f {
				return true
			}
		}
		return false
	}

	if has("--uninstall") {
		uninstall()
		return
	}

	self, _ := os.Executable()
	self, _ = filepath.Abs(self)
	if !strings.EqualFold(filepath.Clean(self), filepath.Clean(installedExe())) && !has("--portable") {
		runSetup(self)
		return
	}

	runApp(has("--relaunch"))
}
