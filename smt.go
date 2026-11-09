//go:build windows

package main

import (
	"encoding/json"
	"errors"
	"net/http"
	"os"
	"path/filepath"
	"regexp"
	"sort"
	"strings"
	"sync"
	"time"

	"pdfunterschreiben/internal/ole"
	"pdfunterschreiben/internal/ole/oleutil"
)

// Servicemitteilungen Toyota: Das Programm sucht in Outlook (Posteingang und Unterordner aller Postfächer)
// nach Mails mit „Servicemitteilung“ im Betreff, speichert die angehängten Mitteilungen (TB…/QM….pdf)
// in Dokumente\Service-Mappe\Servicemitteilungen Toyota und merkt sich Nummer und Titel aus dem Mailtext.

type smtEintrag struct {
	Nr      string   `json:"nr"`
	Titel   string   `json:"titel"`
	Datei   string   `json:"datei"`
	Datum   string   `json:"datum"`
	Betreff string   `json:"betreff"`
	Modelle []string `json:"modelle,omitempty"` // Gerätetypen aus der PDF (Abschnitt „Model(s)“)
	Text    string   `json:"text,omitempty"`    // Text der PDF für die Suche (klein geschrieben)
	Gelesen bool     `json:"gelesen,omitempty"` // PDF wurde schon einmal ausgelesen
}

type smtIndex struct {
	Liste  []smtEintrag `json:"liste"`
	Mails  []string     `json:"mails"` // schon ausgewertete Mails (EntryID)
	Stand  string       `json:"stand"`
	Fehler string       `json:"fehler,omitempty"`
}

var (
	smtMu     sync.Mutex
	smtLaeuft bool
	smtZeile  = regexp.MustCompile(`(?m)^[\s\x{00A0}]*((?:TB|QM)\s?\d{5})[\s\x{00A0}]*[-–—:][\s\x{00A0}]*(.+?)[\s\x{00A0}]*$`)
	smtDatei  = regexp.MustCompile(`(?i)^((?:TB|QM)\d{5})\b.*\.pdf$`)
)

func smtDir() string {
	return filepath.Join(os.Getenv("USERPROFILE"), "Documents", "Service-Mappe", "Servicemitteilungen Toyota")
}
func smtIndexPfad() string { return filepath.Join(smtDir(), "servicemitteilungen.json") }

func smtLaden() smtIndex {
	var x smtIndex
	if b, err := os.ReadFile(smtIndexPfad()); err == nil {
		_ = json.Unmarshal(b, &x)
	}
	return x
}

func smtSpeichern(x smtIndex) error {
	sort.Slice(x.Liste, func(i, j int) bool {
		if x.Liste[i].Datum != x.Liste[j].Datum {
			return x.Liste[i].Datum > x.Liste[j].Datum
		}
		return x.Liste[i].Nr > x.Liste[j].Nr
	})
	b, _ := json.MarshalIndent(x, "", " ")
	if err := os.MkdirAll(smtDir(), 0o755); err != nil {
		return err
	}
	return os.WriteFile(smtIndexPfad(), b, 0o644)
}

// smtAbruf durchsucht Outlook. Läuft Outlook nicht, passiert nichts (Fehler „Outlook läuft nicht“).
func smtAbruf() (smtIndex, error) {
	smtMu.Lock()
	if smtLaeuft {
		smtMu.Unlock()
		return smtLaden(), errors.New("Abruf läuft schon")
	}
	smtLaeuft = true
	smtMu.Unlock()
	defer func() { smtMu.Lock(); smtLaeuft = false; smtMu.Unlock() }()

	x := smtLaden()
	bekannt := map[string]bool{}
	for _, m := range x.Mails {
		bekannt[m] = true
	}
	nach := map[string]int{}
	for i, e := range x.Liste {
		nach[e.Nr] = i
	}
	_ = os.MkdirAll(smtDir(), 0o755)
	err := olNativAuftrag(olAuftrag{art: "fn", fn: func(app *ole.IDispatch) error {
		nv, err := oleutil.CallMethod(app, "GetNamespace", "MAPI")
		if err != nil {
			return err
		}
		ns := nv.ToIDispatch()
		defer ns.Release()
		var eingaenge []*ole.IDispatch
		if sv, err := oleutil.GetProperty(ns, "Stores"); err == nil {
			st := sv.ToIDispatch()
			n := 0
			if c, err := oleutil.GetProperty(st, "Count"); err == nil {
				n = int(c.Val)
			}
			for i := 1; i <= n; i++ {
				iv, err := oleutil.CallMethod(st, "Item", i)
				if err != nil {
					continue
				}
				s := iv.ToIDispatch()
				if fv, err := oleutil.CallMethod(s, "GetDefaultFolder", 6); err == nil && fv.ToIDispatch() != nil {
					eingaenge = append(eingaenge, fv.ToIDispatch())
				}
				s.Release()
			}
			st.Release()
		}
		if len(eingaenge) == 0 {
			fv, err := oleutil.CallMethod(ns, "GetDefaultFolder", 6)
			if err != nil {
				return err
			}
			eingaenge = append(eingaenge, fv.ToIDispatch())
		}
		filter := `@SQL=("urn:schemas:httpmail:subject" LIKE '%ervicemitteilung%' OR "urn:schemas:httpmail:subject" LIKE '%ervice-Mitteilung%')`
		var ordner func(f *ole.IDispatch, tiefe int)
		ordner = func(f *ole.IDispatch, tiefe int) {
			defer f.Release()
			if iv, err := oleutil.GetProperty(f, "Items"); err == nil {
				items := iv.ToIDispatch()
				if rv, err := oleutil.CallMethod(items, "Restrict", filter); err == nil {
					r := rv.ToIDispatch()
					n := 0
					if c, err := oleutil.GetProperty(r, "Count"); err == nil {
						n = int(c.Val)
					}
					for i := 1; i <= n; i++ {
						mv, err := oleutil.CallMethod(r, "Item", i)
						if err != nil {
							continue
						}
						m := mv.ToIDispatch()
						smtMail(m, &x, bekannt, nach)
						m.Release()
					}
					r.Release()
				}
				items.Release()
			}
			if tiefe >= 2 {
				return
			}
			if fv, err := oleutil.GetProperty(f, "Folders"); err == nil {
				fs := fv.ToIDispatch()
				n := 0
				if c, err := oleutil.GetProperty(fs, "Count"); err == nil {
					n = int(c.Val)
				}
				for i := 1; i <= n; i++ {
					if sv, err := oleutil.CallMethod(fs, "Item", i); err == nil {
						ordner(sv.ToIDispatch(), tiefe+1)
					}
				}
				fs.Release()
			}
		}
		for _, f := range eingaenge {
			ordner(f, 0)
		}
		return nil
	}}, 120*time.Second)
	x.Stand = time.Now().Format("2006-01-02 15:04")
	x.Fehler = ""
	if err != nil {
		x.Fehler = err.Error()
	}
	smtMu.Lock()
	jetzt := smtLaden() // inzwischen ausgelesene PDFs (Gerätetypen, Text) nicht überschreiben
	for i := range x.Liste {
		for _, j := range jetzt.Liste {
			if j.Nr == x.Liste[i].Nr && j.Datei == x.Liste[i].Datei && j.Gelesen && !x.Liste[i].Gelesen {
				x.Liste[i].Modelle, x.Liste[i].Text, x.Liste[i].Gelesen = j.Modelle, j.Text, true
			}
		}
	}
	_ = smtSpeichern(x)
	smtMu.Unlock()
	return x, err
}

func smtText(d *ole.IDispatch, eig string) string {
	v, err := oleutil.GetProperty(d, eig)
	if err != nil {
		return ""
	}
	return v.ToString()
}

func smtMail(m *ole.IDispatch, x *smtIndex, bekannt map[string]bool, nach map[string]int) {
	id := smtText(m, "EntryID")
	if id == "" || bekannt[id] {
		return
	}
	betreff := smtText(m, "Subject")
	datum := ""
	if v, err := oleutil.GetProperty(m, "ReceivedTime"); err == nil {
		if t, ok := v.Value().(time.Time); ok {
			datum = t.Format("2006-01-02")
		}
	}
	titel := map[string]string{}
	for _, z := range smtZeile.FindAllStringSubmatch(smtText(m, "Body"), -1) {
		nr := strings.ReplaceAll(strings.ToUpper(z[1]), " ", "")
		if _, ok := titel[nr]; !ok {
			titel[nr] = strings.TrimSpace(z[2])
		}
	}
	gespeichert := map[string]string{}
	if av, err := oleutil.GetProperty(m, "Attachments"); err == nil {
		att := av.ToIDispatch()
		n := 0
		if c, err := oleutil.GetProperty(att, "Count"); err == nil {
			n = int(c.Val)
		}
		for i := 1; i <= n; i++ {
			iv, err := oleutil.CallMethod(att, "Item", i)
			if err != nil {
				continue
			}
			a := iv.ToIDispatch()
			name := smtText(a, "FileName")
			if t := smtDatei.FindStringSubmatch(name); t != nil {
				nr := strings.ToUpper(t[1])
				ziel := filepath.Join(smtDir(), ablageName.ReplaceAllString(name, "_"))
				if _, err := os.Stat(ziel); err != nil {
					_, err = oleutil.CallMethod(a, "SaveAsFile", ziel)
					if err != nil {
						a.Release()
						continue
					}
				}
				gespeichert[nr] = filepath.Base(ziel)
			}
			a.Release()
		}
		att.Release()
	}
	if len(gespeichert) == 0 && len(titel) == 0 {
		bekannt[id] = true
		x.Mails = append(x.Mails, id)
		return
	}
	for nr, datei := range gespeichert {
		e := smtEintrag{Nr: nr, Titel: titel[nr], Datei: datei, Datum: datum, Betreff: betreff}
		if i, ok := nach[nr]; ok {
			if e.Titel == "" {
				e.Titel = x.Liste[i].Titel
			}
			if x.Liste[i].Datum > e.Datum {
				e.Datum = x.Liste[i].Datum
			}
			if x.Liste[i].Datei == e.Datei {
				e.Modelle, e.Text, e.Gelesen = x.Liste[i].Modelle, x.Liste[i].Text, x.Liste[i].Gelesen
			}
			x.Liste[i] = e
		} else {
			nach[nr] = len(x.Liste)
			x.Liste = append(x.Liste, e)
		}
	}
	bekannt[id] = true
	x.Mails = append(x.Mails, id)
}

func smtAutomatisch() {
	go func() {
		time.Sleep(20 * time.Second)
		for {
			if outlookLaeuft() {
				_, _ = smtAbruf()
			}
			time.Sleep(30 * time.Minute)
		}
	}()
}

func (a *app) handleSmt(w http.ResponseWriter, r *http.Request) {
	w.Header().Set("Cache-Control", "no-store")
	var x smtIndex
	var err error
	if r.Method == http.MethodPost {
		x, err = smtAbruf()
	} else {
		x = smtLaden()
	}
	smtMu.Lock()
	laeuft := smtLaeuft
	smtMu.Unlock()
	f := x.Fehler
	if err != nil {
		f = err.Error()
	}
	if x.Liste == nil {
		x.Liste = []smtEintrag{}
	}
	writeJSON(w, map[string]any{"liste": x.Liste, "stand": x.Stand, "fehler": f, "laeuft": laeuft, "ordner": smtDir()})
}

func (a *app) handleSmtOeffnen(w http.ResponseWriter, r *http.Request) {
	if r.Method != http.MethodPost {
		http.Error(w, "POST", http.StatusMethodNotAllowed)
		return
	}
	nr := strings.ToUpper(r.URL.Query().Get("nr"))
	ziel := smtDir()
	if nr != "" {
		ziel = ""
		for _, e := range smtLaden().Liste {
			if e.Nr == nr {
				ziel = filepath.Join(smtDir(), filepath.Base(e.Datei))
			}
		}
		if ziel == "" {
			http.Error(w, "unbekannt", http.StatusNotFound)
			return
		}
		if _, err := os.Stat(ziel); err != nil {
			http.Error(w, "Datei fehlt", http.StatusNotFound)
			return
		}
	} else {
		_ = os.MkdirAll(ziel, 0o755)
	}
	_ = hidden("rundll32.exe", "url.dll,FileProtocolHandler", ziel).Start()
	w.WriteHeader(http.StatusNoContent)
}

// PDF einer Mitteilung für den eingebauten PDF-Leser
func (a *app) handleSmtPdf(w http.ResponseWriter, r *http.Request) {
	nr := strings.ToUpper(r.URL.Query().Get("nr"))
	for _, e := range smtLaden().Liste {
		if e.Nr == nr {
			p := filepath.Join(smtDir(), filepath.Base(e.Datei))
			f, err := os.Open(p)
			if err != nil {
				http.Error(w, "Datei fehlt", http.StatusNotFound)
				return
			}
			defer f.Close()
			st, _ := f.Stat()
			w.Header().Set("Content-Type", "application/pdf")
			w.Header().Set("Content-Disposition", `inline; filename="`+strings.ReplaceAll(filepath.Base(e.Datei), `"`, "")+`"`)
			http.ServeContent(w, r, "", st.ModTime(), f)
			return
		}
	}
	http.Error(w, "unbekannt", http.StatusNotFound)
}

// Ergebnis des Auslesens (Gerätetypen und Text) aus der Oberfläche speichern
func (a *app) handleSmtInfo(w http.ResponseWriter, r *http.Request) {
	if r.Method != http.MethodPost {
		http.Error(w, "POST", http.StatusMethodNotAllowed)
		return
	}
	var in struct {
		Nr      string   `json:"nr"`
		Modelle []string `json:"modelle"`
		Text    string   `json:"text"`
	}
	if err := json.NewDecoder(http.MaxBytesReader(w, r.Body, 2<<20)).Decode(&in); err != nil {
		http.Error(w, "json", http.StatusBadRequest)
		return
	}
	if len(in.Text) > 60000 {
		in.Text = in.Text[:60000]
	}
	smtMu.Lock()
	defer smtMu.Unlock()
	x := smtLaden()
	for i := range x.Liste {
		if x.Liste[i].Nr == strings.ToUpper(in.Nr) {
			x.Liste[i].Modelle, x.Liste[i].Text, x.Liste[i].Gelesen = in.Modelle, in.Text, true
		}
	}
	_ = smtSpeichern(x)
	w.WriteHeader(http.StatusNoContent)
}
