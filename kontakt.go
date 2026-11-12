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
		Firma   string `json:"firma"`
		Name    string `json:"name"`
		Email   string `json:"email"`
		Tel     string `json:"tel"`
		Strasse string `json:"strasse"`
		Ort     string `json:"ort"`
	}
	if err := json.NewDecoder(http.MaxBytesReader(w, r.Body, 1<<16)).Decode(&in); err != nil {
		writeJSON(w, map[string]any{"ok": false, "error": "Ungültige Anfrage."})
		return
	}
	in.Email = strings.TrimSpace(in.Email)
	in.Name = strings.TrimSpace(in.Name)
	in.Firma = strings.TrimSpace(in.Firma)
	in.Tel = strings.TrimSpace(in.Tel)
	in.Strasse = strings.TrimSpace(in.Strasse)
	in.Ort = strings.TrimSpace(in.Ort)
	hasMail := kontaktMail.MatchString(in.Email)
	if !hasMail && in.Tel == "" {
		writeJSON(w, map[string]any{"ok": false, "error": "Weder E-Mail noch Telefon."})
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
		if hasMail {
			if fv, err := oleutil.CallMethod(items, "Find", "[Email1Address] = '"+strings.ReplaceAll(in.Email, "'", "''")+"'"); err == nil {
				if d := fv.ToIDispatch(); d != nil {
					d.Release()
					return nil
				}
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
		if hasMail {
			_, _ = oleutil.PutProperty(c, "Email1Address", in.Email)
			_, _ = oleutil.PutProperty(c, "Email1DisplayName", strings.TrimSpace(name+" ("+in.Firma+")"))
		}
		if in.Tel != "" {
			_, _ = oleutil.PutProperty(c, "BusinessTelephoneNumber", in.Tel)
		}
		if in.Strasse != "" {
			_, _ = oleutil.PutProperty(c, "BusinessAddressStreet", in.Strasse)
		}
		if in.Ort != "" {
			_, _ = oleutil.PutProperty(c, "BusinessAddressCity", in.Ort)
		}
		_, err = oleutil.CallMethod(c, "Save")
		return err
	}}, 60*time.Second)
	if err != nil {
		writeJSON(w, map[string]any{"ok": false, "error": err.Error()})
		return
	}
	writeJSON(w, map[string]any{"ok": true})
}

func kontaktEig(c *ole.IDispatch, eig string) string {
	v, err := oleutil.GetProperty(c, eig)
	if err != nil || v == nil {
		return ""
	}
	return strings.TrimSpace(v.ToString())
}

type kontaktEintrag struct {
	Firma   string `json:"firma"`
	Name    string `json:"name"`
	Email   string `json:"email"`
	Tel     string `json:"tel"`
	Strasse string `json:"strasse"`
	Ort     string `json:"ort"`
}

// Liest die vorhandenen Outlook-Kontakte (Firma, Name, E-Mail, Telefon, Adresse), damit sie als
// persönliche Kundendaten übernommen werden können (umgekehrte Richtung zu handleKontakt).
func (a *app) handleKontaktImport(w http.ResponseWriter, r *http.Request) {
	if r.Method != http.MethodPost {
		http.Error(w, "POST", http.StatusMethodNotAllowed)
		return
	}
	var liste []kontaktEintrag
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
		cv, err := oleutil.GetProperty(items, "Count")
		if err != nil {
			return err
		}
		n := int(cv.Val)
		if n > 3000 {
			n = 3000 // Sicherheitsgrenze, falls das Adressbuch riesig ist
		}
		for i := 1; i <= n; i++ {
			iv2, err := oleutil.CallMethod(items, "Item", i)
			if err != nil {
				continue
			}
			c := iv2.ToIDispatch()
			if c == nil {
				continue
			}
			em := kontaktEig(c, "Email1Address")
			tel := kontaktEig(c, "BusinessTelephoneNumber")
			if tel == "" {
				tel = kontaktEig(c, "MobileTelephoneNumber")
			}
			if em != "" || tel != "" {
				liste = append(liste, kontaktEintrag{
					Firma:   kontaktEig(c, "CompanyName"),
					Name:    kontaktEig(c, "FullName"),
					Email:   em,
					Tel:     tel,
					Strasse: kontaktEig(c, "BusinessAddressStreet"),
					Ort:     kontaktEig(c, "BusinessAddressCity"),
				})
			}
			c.Release()
		}
		return nil
	}}, 90*time.Second)
	if err != nil {
		writeJSON(w, map[string]any{"ok": false, "error": err.Error()})
		return
	}
	writeJSON(w, map[string]any{"ok": true, "liste": liste})
}
