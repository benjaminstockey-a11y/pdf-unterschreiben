//go:build windows

package main

import (
	"errors"
	"net/http"
	"strings"
	"time"

	"pdfunterschreiben/internal/ole"
	"pdfunterschreiben/internal/ole/oleutil"
)

// Termine aus dem Outlook-Kalender für den Kalender auf der Startseite (nur lesen).
// „Neuer Termin“ öffnet ein Outlook-Terminfenster für den gewählten Tag.

type termin struct {
	Start   string `json:"start"` // 2006-01-02T15:04 (Ortszeit)
	Ende    string `json:"ende"`
	Betreff string `json:"betreff"`
	Ort     string `json:"ort,omitempty"`
	Ganztag bool   `json:"ganztag,omitempty"`
}

func olDatum(v *ole.VARIANT) (time.Time, bool) {
	if v == nil {
		return time.Time{}, false
	}
	t, ok := v.Value().(time.Time)
	return t, ok
}

func termineLesen(von, bis time.Time) ([]termin, error) {
	var liste []termin
	err := olNativAuftrag(olAuftrag{art: "fn", fn: func(app *ole.IDispatch) error {
		nv, err := oleutil.CallMethod(app, "GetNamespace", "MAPI")
		if err != nil {
			return err
		}
		ns := nv.ToIDispatch()
		defer ns.Release()
		fv, err := oleutil.CallMethod(ns, "GetDefaultFolder", 9) // Kalender
		if err != nil {
			return err
		}
		f := fv.ToIDispatch()
		defer f.Release()
		iv, err := oleutil.GetProperty(f, "Items")
		if err != nil {
			return err
		}
		items := iv.ToIDispatch()
		defer items.Release()
		_, _ = oleutil.PutProperty(items, "IncludeRecurrences", true)
		_, _ = oleutil.CallMethod(items, "Sort", "[Start]")
		// Datumsformat für den Filter hängt von der Windows-Sprache ab: erst deutsch, dann englisch
		var r *ole.IDispatch
		for _, form := range []string{"02.01.2006 15:04", "01/02/2006 03:04 PM"} {
			filter := "[Start] < '" + bis.Format(form) + "' AND [End] > '" + von.Format(form) + "'"
			if rv, err := oleutil.CallMethod(items, "Restrict", filter); err == nil {
				r = rv.ToIDispatch()
				break
			}
		}
		if r == nil {
			return errors.New("Kalender konnte nicht gefiltert werden")
		}
		defer r.Release()
		mv, err := oleutil.CallMethod(r, "GetFirst")
		for n := 0; err == nil && mv != nil && mv.VT == ole.VT_DISPATCH && n < 600; n++ {
			m := mv.ToIDispatch()
			if m == nil {
				break
			}
			sv, _ := oleutil.GetProperty(m, "Start")
			ev, _ := oleutil.GetProperty(m, "End")
			s, ok1 := olDatum(sv)
			e, ok2 := olDatum(ev)
			if ok1 && ok2 && s.Before(bis) && e.After(von) {
				t := termin{Start: s.Format("2006-01-02T15:04"), Ende: e.Format("2006-01-02T15:04"), Betreff: strings.TrimSpace(smtText(m, "Subject")), Ort: strings.TrimSpace(smtText(m, "Location"))}
				if gv, err := oleutil.GetProperty(m, "AllDayEvent"); err == nil {
					t.Ganztag, _ = gv.Value().(bool)
				}
				liste = append(liste, t)
			}
			m.Release()
			mv, err = oleutil.CallMethod(r, "GetNext")
		}
		return nil
	}}, 60*time.Second)
	return liste, err
}

func (a *app) handleTermine(w http.ResponseWriter, r *http.Request) {
	w.Header().Set("Cache-Control", "no-store")
	von, err1 := time.ParseInLocation("2006-01-02", r.URL.Query().Get("von"), time.Local)
	bis, err2 := time.ParseInLocation("2006-01-02", r.URL.Query().Get("bis"), time.Local)
	if err1 != nil || err2 != nil || !bis.After(von) || bis.Sub(von) > 100*24*time.Hour {
		writeJSON(w, map[string]any{"ok": false, "error": "Zeitraum ungültig"})
		return
	}
	l, err := termineLesen(von, bis)
	if err != nil {
		f := err.Error()
		if f == "Outlook läuft nicht" {
			f = "Outlook ist nicht geöffnet"
		}
		writeJSON(w, map[string]any{"ok": false, "error": f})
		return
	}
	if l == nil {
		l = []termin{}
	}
	writeJSON(w, map[string]any{"ok": true, "termine": l})
}

func (a *app) handleTerminNeu(w http.ResponseWriter, r *http.Request) {
	if r.Method != http.MethodPost {
		http.Error(w, "POST", http.StatusMethodNotAllowed)
		return
	}
	tag, err := time.ParseInLocation("2006-01-02", r.URL.Query().Get("datum"), time.Local)
	if err != nil {
		tag = time.Now()
	}
	start := time.Date(tag.Year(), tag.Month(), tag.Day(), 8, 0, 0, 0, time.Local)
	err = olNativAuftrag(olAuftrag{art: "fn", fn: func(app *ole.IDispatch) error {
		v, err := oleutil.CallMethod(app, "CreateItem", 1) // Termin
		if err != nil {
			return err
		}
		t := v.ToIDispatch()
		defer t.Release()
		_, _ = oleutil.PutProperty(t, "Start", start)
		_, _ = oleutil.PutProperty(t, "Duration", 60)
		if _, err := oleutil.CallMethod(t, "Display", false); err != nil {
			return err
		}
		if iv, err := oleutil.GetProperty(t, "GetInspector"); err == nil {
			if in := iv.ToIDispatch(); in != nil {
				_, _ = oleutil.CallMethod(in, "Activate")
				in.Release()
			}
		}
		return nil
	}}, 60*time.Second)
	if err != nil {
		writeJSON(w, map[string]any{"ok": false, "error": err.Error()})
		return
	}
	writeJSON(w, map[string]any{"ok": true})
}
