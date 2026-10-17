//go:build windows

package main

import (
	_ "embed"
	"fmt"
	"io"
	"net/http"
	"os"
	"path/filepath"
	"strings"
	"time"
)

//go:embed assets/app.ico
var appIcon []byte

// quitRunning bittet eine laufende Instanz, sich zu beenden (damit die exe ersetzt werden kann).
func quitRunning() {
	c := &http.Client{Timeout: 2 * time.Second}
	resp, err := c.Post(fmt.Sprintf("http://127.0.0.1:%d/api/quit", Port), "text/plain", nil)
	if err == nil {
		resp.Body.Close()
		time.Sleep(1200 * time.Millisecond)
	}
}

func copyFile(src, dst string) error {
	in, err := os.Open(src)
	if err != nil {
		return err
	}
	defer in.Close()
	tmp := dst + ".tmp"
	out, err := os.Create(tmp)
	if err != nil {
		return err
	}
	if _, err := io.Copy(out, in); err != nil {
		out.Close()
		os.Remove(tmp)
		return err
	}
	if err := out.Close(); err != nil {
		return err
	}
	return replaceFile(tmp, dst)
}

// replaceFile setzt neu an die Stelle von dst. Eine laufende exe lässt sich unter Windows nicht
// überschreiben, aber umbenennen – deshalb wird die alte Datei zuerst beiseitegeschoben.
func replaceFile(neu, dst string) error {
	if _, err := os.Stat(dst); err == nil {
		old := fmt.Sprintf("%s.%d.old", dst, time.Now().UnixNano())
		if err := os.Rename(dst, old); err != nil {
			return err
		}
	}
	return os.Rename(neu, dst)
}

func cleanupOld() {
	matches, _ := filepath.Glob(filepath.Join(installDir(), "*.old"))
	for _, m := range matches {
		_ = os.Remove(m)
	}
	tmps, _ := filepath.Glob(filepath.Join(installDir(), "*.tmp"))
	for _, m := range tmps {
		_ = os.Remove(m)
	}
}

func runSetup(self string) {
	exists := false
	if _, err := os.Stat(installedExe()); err == nil {
		exists = true
	}
	q := fmt.Sprintf("%s %s installieren?\n\nDas Programm wird nur für dein Windows-Konto installiert, Administratorrechte sind nicht nötig. Es kommt eine Verknüpfung auf den Desktop und ins Startmenü.", AppName, Version)
	if exists {
		q = fmt.Sprintf("%s ist schon installiert.\n\nAuf Version %s aktualisieren? Deine Unterschrift und Einstellungen bleiben erhalten.", AppName, Version)
	}
	if !ask(q) {
		return
	}
	quitRunning()

	if err := os.MkdirAll(installDir(), 0o755); err != nil {
		fail("Der Installationsordner konnte nicht angelegt werden:\n" + err.Error())
		return
	}
	cleanupOld()
	if err := copyFile(self, installedExe()); err != nil {
		fail("Die Installation ist fehlgeschlagen:\n" + err.Error() + "\n\nSchließe PDF Unterschreiben und versuch es nochmal.")
		return
	}
	_ = os.WriteFile(iconPath(), appIcon, 0o644)
	if err := createShortcuts(true); err != nil {
		// Nicht kritisch – das Programm läuft trotzdem.
		info("Installiert, aber die Verknüpfungen konnten nicht angelegt werden.\nDu findest das Programm hier:\n" + installedExe())
	}
	registerUninstall()

	if err := detached(installedExe()); err != nil {
		fail("Installiert, aber das Programm konnte nicht gestartet werden:\n" + err.Error())
	}
}

func uninstall() {
	if !ask(AppName + " wirklich vom Computer entfernen?") {
		return
	}
	quitRunning()
	removeShortcuts()
	unregisterUninstall()

	paths := []string{installDir(), filepath.Join(os.Getenv("LOCALAPPDATA"), AppName)}
	if ask("Soll auch deine gespeicherte Unterschrift mit den Einstellungen gelöscht werden?\n\nWähle „Nein“, wenn du das Programm später wieder installieren willst.") {
		paths = append(paths, filepath.Join(os.Getenv("APPDATA"), AppName))
	}
	// Den eigenen Ordner kann die laufende exe nicht löschen – das erledigt cmd kurz nach dem Beenden.
	var parts []string
	for _, p := range paths {
		parts = append(parts, `rmdir /s /q "`+p+`"`)
	}
	script := "ping 127.0.0.1 -n 3 > nul & " + strings.Join(parts, " & ")
	_ = detached("cmd.exe", "/c", script)
	info(AppName + " wurde entfernt.")
}
