//go:build windows

package main

import (
	"errors"
	"fmt"
	"os/exec"
	"runtime"
	"strings"
	"syscall"
	"time"
	"unsafe"

	"pdfunterschreiben/internal/ole"
	"pdfunterschreiben/internal/ole/oleutil"
)

// Outlook direkt aus dem Programm steuern (COM), ganz ohne PowerShell.
// PowerShell wird auf den Firmen-Laptops vom Virenscanner (Sophos) bei jedem Start geprüft –
// das kostet Sekunden. Ein eigener Faden hält die Verbindung zu Outlook offen.

type olAuftrag struct {
	art   string // "leicht" (nur verbinden), "vor"/"start" (notfalls unsichtbar starten), "mail", "ende"
	subj  string
	to    string
	files []string
	antw  chan error
}

var olNativ = make(chan olAuftrag)

func init() { go olNativSchleife() }

func outlookLaeuft() bool {
	snap, err := syscall.CreateToolhelp32Snapshot(syscall.TH32CS_SNAPPROCESS, 0)
	if err != nil {
		return false
	}
	defer syscall.CloseHandle(snap)
	var e syscall.ProcessEntry32
	e.Size = uint32(unsafe.Sizeof(e))
	for err = syscall.Process32First(snap, &e); err == nil; err = syscall.Process32Next(snap, &e) {
		if strings.EqualFold(syscall.UTF16ToString(e.ExeFile[:]), "outlook.exe") {
			return true
		}
	}
	return false
}

func olNativSchleife() {
	runtime.LockOSThread()
	if err := ole.CoInitializeEx(0, ole.COINIT_APARTMENTTHREADED); err != nil {
		var oe *ole.OleError
		if !errors.As(err, &oe) { // S_FALSE (schon initialisiert) ist in Ordnung
			for a := range olNativ {
				a.antw <- err
			}
			return
		}
	}
	var app *ole.IDispatch
	wirGestartet := false // Outlook wurde von der Service-Mappe unsichtbar gestartet → beim Beenden wieder schließen
	verbinden := func(starten bool) error {
		if app != nil {
			if _, err := oleutil.GetProperty(app, "Version"); err == nil {
				return nil
			}
			app.Release()
			app = nil
		}
		if !outlookLaeuft() {
			if !starten {
				return errors.New("Outlook läuft nicht")
			}
			// Outlook unsichtbar im Hintergrund starten (ohne Fenster) und Postfach schon laden,
			// damit „Per E-Mail senden“ sofort reagiert
			unk, err := oleutil.CreateObject("Outlook.Application")
			if err == nil {
				d, err2 := unk.QueryInterface(ole.IID_IDispatch)
				unk.Release()
				if err2 == nil {
					app = d
					wirGestartet = true
					if nv, err := oleutil.CallMethod(app, "GetNamespace", "MAPI"); err == nil {
						ns := nv.ToIDispatch()
						_, _ = oleutil.CallMethod(ns, "Logon", "", "", false, false)
						if fv, err := oleutil.CallMethod(ns, "GetDefaultFolder", 6); err == nil { // Posteingang
							fv.ToIDispatch().Release()
						}
						ns.Release()
					}
					return nil
				}
			}
			// Rückfall: normal (minimiert) starten
			c := exec.Command("cmd", "/c", "start", "/min", "", "outlook.exe")
			c.SysProcAttr = &syscall.SysProcAttr{HideWindow: true, CreationFlags: createNoWindow}
			_ = c.Run()
			for i := 0; i < 120 && !outlookLaeuft(); i++ {
				time.Sleep(500 * time.Millisecond)
			}
		}
		// warten, bis sich das laufende Outlook bei Windows angemeldet hat
		var letzt error
		for i := 0; i < 60; i++ {
			unk, err := oleutil.GetActiveObject("Outlook.Application")
			if err != nil && outlookLaeuft() && i >= 4 {
				unk, err = oleutil.CreateObject("Outlook.Application") // verbindet mit dem laufenden Outlook
			}
			if err == nil {
				d, err2 := unk.QueryInterface(ole.IID_IDispatch)
				unk.Release()
				if err2 == nil {
					app = d
					return nil
				}
				err = err2
			}
			letzt = err
			time.Sleep(500 * time.Millisecond)
		}
		return fmt.Errorf("keine Verbindung zu Outlook: %v", letzt)
	}
	for a := range olNativ {
		start := time.Now()
		var t []string
		mark := func(n string) { t = append(t, fmt.Sprintf("%s %d", n, time.Since(start).Milliseconds())) }
		if a.art == "ende" {
			// nur schließen, wenn wir Outlook gestartet haben und kein Outlook-Fenster offen ist
			// (sonst ginge z. B. eine noch offene Mail verloren)
			if app != nil && wirGestartet {
				offen := 0
				for _, eig := range []string{"Explorers", "Inspectors"} {
					if v, err := oleutil.GetProperty(app, eig); err == nil {
						d := v.ToIDispatch()
						if c, err := oleutil.GetProperty(d, "Count"); err == nil {
							offen += int(c.Val)
						}
						d.Release()
					}
				}
				if offen == 0 {
					_, _ = oleutil.CallMethod(app, "Quit")
				}
			}
			if app != nil {
				app.Release()
				app = nil
			}
			a.antw <- nil
			continue
		}
		err := verbinden(a.art != "leicht")
		mark("Outlook")
		if err == nil && a.art == "mail" {
			err = func() (err error) {
				defer func() {
					if r := recover(); r != nil {
						err = fmt.Errorf("%v", r)
					}
				}()
				v, err := oleutil.CallMethod(app, "CreateItem", 0)
				if err != nil {
					return err
				}
				m := v.ToIDispatch()
				defer m.Release()
				mark("Neu")
				if _, err := oleutil.PutProperty(m, "Subject", a.subj); err != nil {
					return err
				}
				if a.to != "" {
					_, _ = oleutil.PutProperty(m, "To", a.to)
				}
				av, err := oleutil.GetProperty(m, "Attachments")
				if err != nil {
					return err
				}
				att := av.ToIDispatch()
				for _, f := range a.files {
					if _, err := oleutil.CallMethod(att, "Add", f); err != nil {
						att.Release()
						return err
					}
				}
				att.Release()
				mark("Anhaenge")
				if _, err := oleutil.CallMethod(m, "Display", false); err != nil {
					return err
				}
				mark("Anzeigen")
				if iv, err := oleutil.GetProperty(m, "GetInspector"); err == nil {
					if in := iv.ToIDispatch(); in != nil {
						_, _ = oleutil.CallMethod(in, "Activate")
						in.Release()
					}
				}
				return nil
			}()
			if err != nil && app != nil { // Verbindung evtl. kaputt – beim nächsten Mal neu
				app.Release()
				app = nil
			}
		}
		olw.zeiten = "nativ: " + strings.Join(t, ", ")
		a.antw <- err
	}
}

// olNativAuftrag schickt einen Auftrag an den Outlook-Faden (mit Zeitlimit).
func olNativAuftrag(a olAuftrag, timeout time.Duration) error {
	a.antw = make(chan error, 1)
	select {
	case olNativ <- a:
	case <-time.After(timeout):
		return errors.New("Outlook-Faden beschäftigt")
	}
	select {
	case err := <-a.antw:
		return err
	case <-time.After(timeout):
		return errors.New("Zeitüberschreitung")
	}
}
