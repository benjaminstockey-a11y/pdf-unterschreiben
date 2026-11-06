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
	art   string // "leicht" (nur verbinden, falls Outlook läuft), "vor" (notfalls starten), "mail"
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
