//go:build windows

package main

import (
	"encoding/json"
	"net/http"
	"regexp"
	"strings"
	"time"

	"pdfunterschreiben/internal/ole"
	"pdfunterschreiben/internal/ole/oleutil"
)

var kontaktMail = regexp.MustCompile(`^[^\s@]+@[^\s@]+\.[^\s@]+$`)

// Legt die E-Mail-Adresse eines Kunden als Kontakt in Outlook an, falls es sie dort noch nicht gibt.
// Dann schlägt Outlook die Adresse beim Tippen im „An“-Feld von selbst vor.
func (a *app) handleKontakt(w http.ResponseWriter, r *http.Request) {
	if r.Method != http.MethodPost {
		http.Error(w, "POST", http.StatusMethodNotAllowed)
		return
	}
	var in struct {
		Firma string `json:"firma"`
		Name  string `json:"name"`
		Email string `json:"email"`
	}
	if err := json.NewDecoder(http.MaxBytesReader(w, r.Body, 1<<16)).Decode(&in); err != nil {
		writeJSON(w, map[string]any{"ok": false, "error": "Ungültige Anfrage."})
		return
	}
	in.Email = strings.TrimSpace(in.Email)
	in.Name = strings.TrimSpace(in.Name)
	in.Firma = strings.TrimSpace(in.Firma)
	if !kontaktMail.MatchString(in.Email) {
		writeJSON(w, map[string]any{"ok": false, "error": "Keine gültige E-Mail."})
		return
	}
	err := olNativAuftrag(olAuftrag{art: "fn", fn: func(app *ole.IDispatch) error {
		nv, err := oleutil.CallMethod(app, "GetNamespace", "MAPI")
		if err != nil {
			return err
		}
		ns := nv.ToIDispatch()
		defer ns.Release()
		fv, err := oleutil.CallMethod(ns, "GetDefaultFolder", 10) // olFolderContacts
		if err != nil {
			return err
		}
		folder := fv.ToIDispatch()
		defer folder.Release()
		iv, err := oleutil.GetProperty(folder, "Items")
		if err != nil {
			return err
		}
		items := iv.ToIDispatch()
		defer items.Release()
		// Gibt es die E-Mail-Adresse schon? Dann nichts anlegen (keine Dubletten).
		if fv, err := oleutil.CallMethod(items, "Find", "[Email1Address] = '"+strings.ReplaceAll(in.Email, "'", "''")+"'"); err == nil {
			if d := fv.ToIDispatch(); d != nil {
				d.Release()
				return nil
			}
		}
		cv, err := oleutil.CallMethod(app, "CreateItem", 2) // olContactItem
		if err != nil {
			return err
		}
		c := cv.ToIDispatch()
		defer c.Release()
		name := in.Name
		if name == "" {
			name = in.Firma
		}
		if name != "" {
			_, _ = oleutil.PutProperty(c, "FullName", name)
		}
		if in.Firma != "" {
			_, _ = oleutil.PutProperty(c, "CompanyName", in.Firma)
		}
		_, _ = oleutil.PutProperty(c, "Email1Address", in.Email)
		_, _ = oleutil.PutProperty(c, "Email1DisplayName", strings.TrimSpace(name+" ("+in.Firma+")"))
		_, err = oleutil.CallMethod(c, "Save")
		return err
	}}, 60*time.Second)
	if err != nil {
		writeJSON(w, map[string]any{"ok": false, "error": err.Error()})
		return
	}
	writeJSON(w, map[string]any{"ok": true})
}
