//go:build windows

package main

import (
	"os"
	"os/exec"
	"path/filepath"
	"syscall"
	"unsafe"
)

var (
	user32      = syscall.NewLazyDLL("user32.dll")
	pMessageBox = user32.NewProc("MessageBoxW")
)

const (
	mbOK           = 0x0
	mbYesNo        = 0x4
	mbIconError    = 0x10
	mbIconQuestion = 0x20
	mbIconInfo     = 0x40
	mbTopmost      = 0x40000
	idYes          = 6
)

func msgBox(text string, flags uintptr) int {
	t, _ := syscall.UTF16PtrFromString(text)
	c, _ := syscall.UTF16PtrFromString(AppName)
	r, _, _ := pMessageBox.Call(0, uintptr(unsafe.Pointer(t)), uintptr(unsafe.Pointer(c)), flags|mbTopmost)
	return int(r)
}
func info(text string)     { msgBox(text, mbOK|mbIconInfo) }
func fail(text string)     { msgBox(text, mbOK|mbIconError) }
func ask(text string) bool { return msgBox(text, mbYesNo|mbIconQuestion) == idYes }

const createNoWindow = 0x08000000
const detachedProcess = 0x00000008
const newProcessGroup = 0x00000200

// hidden startet ein Hilfsprogramm ohne sichtbares Konsolenfenster.
func hidden(name string, args ...string) *exec.Cmd {
	cmd := exec.Command(name, args...)
	cmd.SysProcAttr = &syscall.SysProcAttr{HideWindow: true, CreationFlags: createNoWindow}
	return cmd
}

// detached startet ein Programm, das unabhängig von diesem Prozess weiterläuft.
func detached(name string, args ...string) error {
	cmd := exec.Command(name, args...)
	cmd.SysProcAttr = &syscall.SysProcAttr{HideWindow: true, CreationFlags: detachedProcess | newProcessGroup}
	return cmd.Start()
}

func powershell(script string, env ...string) error {
	cmd := hidden("powershell.exe", "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-Command", script)
	cmd.Env = append(os.Environ(), env...)
	return cmd.Run()
}

func startMenuLink() string {
	return filepath.Join(os.Getenv("APPDATA"), `Microsoft\Windows\Start Menu\Programs`, AppName+".lnk")
}

// createShortcuts legt Startmenü- und Desktop-Verknüpfung an (Desktop-Pfad über Windows ermittelt,
// damit auch ein auf OneDrive umgeleiteter Desktop funktioniert).
func createShortcuts(withDesktop bool) error {
	desk := "0"
	if withDesktop {
		desk = "1"
	}
	script := `$w = New-Object -ComObject WScript.Shell
$paths = @($env:PU_START)
if ($env:PU_DESK -eq '1') { $paths += (Join-Path ([Environment]::GetFolderPath('Desktop')) ($env:PU_NAME + '.lnk')) }
foreach ($p in $paths) {
  $s = $w.CreateShortcut($p)
  $s.TargetPath = $env:PU_TARGET
  $s.WorkingDirectory = $env:PU_DIR
  $s.IconLocation = $env:PU_ICON
  $s.Description = 'PDFs mit deiner Unterschrift versehen'
  $s.Save()
}`
	return powershell(script,
		"PU_START="+startMenuLink(),
		"PU_DESK="+desk,
		"PU_NAME="+AppName,
		"PU_TARGET="+installedExe(),
		"PU_DIR="+installDir(),
		"PU_ICON="+iconPath(),
	)
}

func iconPath() string { return filepath.Join(installDir(), "logo.ico") }

// refreshIcon: Nach einem Update mit neuem Logo vorhandene Verknüpfungen auf die neue Symboldatei
// umstellen (neuer Dateiname, damit Windows nicht das alte Symbol aus dem Zwischenspeicher zeigt).
func refreshIcon() {
	old := filepath.Join(installDir(), "app.ico")
	if _, err := os.Stat(old); err != nil {
		return
	}
	_ = powershell(`$w = New-Object -ComObject WScript.Shell
$paths = @($env:PU_START, (Join-Path ([Environment]::GetFolderPath('Desktop')) ($env:PU_NAME + '.lnk')))
foreach ($p in $paths) { if (Test-Path $p) { $s = $w.CreateShortcut($p); $s.IconLocation = $env:PU_ICON; $s.Save() } }`,
		"PU_START="+startMenuLink(), "PU_NAME="+AppName, "PU_ICON="+iconPath())
	regSet("DisplayIcon", "REG_SZ", iconPath())
	_ = os.Remove(old)
	_ = hidden("ie4uinit.exe", "-show").Run()
}

func removeShortcuts() {
	_ = os.Remove(startMenuLink())
	_ = powershell(`$p = Join-Path ([Environment]::GetFolderPath('Desktop')) ($env:PU_NAME + '.lnk'); if (Test-Path $p) { Remove-Item $p -Force }`, "PU_NAME="+AppName)
}

const uninstallKey = `HKCU\Software\Microsoft\Windows\CurrentVersion\Uninstall\` + AppID

func regSet(name, typ, value string) {
	_ = hidden("reg.exe", "add", uninstallKey, "/v", name, "/t", typ, "/d", value, "/f").Run()
}

// registerUninstall trägt das Programm unter „Apps & Features“ ein.
func registerUninstall() {
	exe := installedExe()
	regSet("DisplayName", "REG_SZ", AppName)
	regSet("DisplayVersion", "REG_SZ", Version)
	regSet("Publisher", "REG_SZ", Publisher)
	regSet("DisplayIcon", "REG_SZ", iconPath())
	regSet("InstallLocation", "REG_SZ", installDir())
	regSet("UninstallString", "REG_SZ", `"`+exe+`" --uninstall`)
	regSet("NoModify", "REG_DWORD", "1")
	regSet("NoRepair", "REG_DWORD", "1")
}

func unregisterUninstall() {
	_ = hidden("reg.exe", "delete", uninstallKey, "/f").Run()
}

// findBrowser sucht Edge (auf jedem Windows 10/11 vorhanden), sonst Chrome.
func findBrowser() string {
	pf, pf86, lad := os.Getenv("ProgramFiles"), os.Getenv("ProgramFiles(x86)"), os.Getenv("LOCALAPPDATA")
	for _, p := range []string{
		filepath.Join(pf86, `Microsoft\Edge\Application\msedge.exe`),
		filepath.Join(pf, `Microsoft\Edge\Application\msedge.exe`),
		filepath.Join(lad, `Microsoft\Edge\Application\msedge.exe`),
		filepath.Join(pf, `Google\Chrome\Application\chrome.exe`),
		filepath.Join(pf86, `Google\Chrome\Application\chrome.exe`),
		filepath.Join(lad, `Google\Chrome\Application\chrome.exe`),
	} {
		if st, err := os.Stat(p); err == nil && !st.IsDir() {
			return p
		}
	}
	return ""
}

// openWindow öffnet die Oberfläche als eigenes App-Fenster ohne Adressleiste.
func openWindow(url string) {
	if b := findBrowser(); b != "" {
		profile := filepath.Join(os.Getenv("LOCALAPPDATA"), AppName, "Fenster")
		_ = os.MkdirAll(profile, 0o755)
		cmd := exec.Command(b,
			"--app="+url,
			"--user-data-dir="+profile,
			"--window-size=1280,900",
			"--no-first-run", "--no-default-browser-check",
			"--disable-features=Translate",
		)
		if cmd.Start() == nil {
			go cmd.Wait()
			return
		}
	}
	// Notlösung: Standardbrowser
	_ = hidden("rundll32.exe", "url.dll,FileProtocolHandler", url).Start()
}
