//go:build windows

package main

import (
	"fmt"
	"hash/fnv"
	"io"
	"net"
	"net/http"
	"os"
	"strings"
	"time"
)

// Mehrere Windows-Benutzer am selben Rechner: Jeder bekommt seinen eigenen Anschluss (Port),
// und vor dem Verbinden wird geprüft, ob das laufende Programm wirklich diesem Benutzer gehört.
// Sonst landete das Fenster des einen im Programm des anderen – mit fremdem Namen und fremder
// Unterschrift.

const portSlots = 12

func ownerToken() string {
	h := fnv.New64a()
	h.Write([]byte(strings.ToLower(os.Getenv("APPDATA") + "|" + os.Getenv("USERDOMAIN") + `\` + os.Getenv("USERNAME"))))
	return fmt.Sprintf("%016x", h.Sum64())
}

func basePort() int {
	h := fnv.New32a()
	h.Write([]byte(ownerToken()))
	return 47613 + int(h.Sum32()%400)*2
}

// portOwner liefert die Kennung des Programms auf diesem Port ("" = keins von uns / nichts).
func portOwner(p int) string {
	c := &http.Client{Timeout: 800 * time.Millisecond}
	resp, err := c.Get(fmt.Sprintf("http://127.0.0.1:%d/api/owner", p))
	if err != nil {
		return ""
	}
	defer resp.Body.Close()
	b, _ := io.ReadAll(io.LimitReader(resp.Body, 64))
	if resp.StatusCode != 200 {
		return ""
	}
	return strings.TrimSpace(string(b))
}

// ownPort sucht die laufende Instanz dieses Benutzers (0 = läuft nicht).
func ownPort() int {
	me := ownerToken()
	for k := 0; k < portSlots; k++ {
		if p := basePort() + k; portOwner(p) == me {
			return p
		}
	}
	return 0
}

// listenOwn belegt einen freien Port für diesen Benutzer. Läuft das Programm für ihn schon,
// kommt (nil, true) zurück und Port zeigt auf die laufende Instanz.
func listenOwn(relaunch bool) (net.Listener, bool) {
	me := ownerToken()
	for k := 0; k < portSlots; k++ {
		p := basePort() + k
		for i := 0; i < 40; i++ {
			ln, err := net.Listen("tcp", fmt.Sprintf("127.0.0.1:%d", p))
			if err == nil {
				Port = p
				return ln, false
			}
			owner := portOwner(p)
			if owner == me {
				if relaunch { // die alte Version beendet sich gerade
					time.Sleep(250 * time.Millisecond)
					continue
				}
				Port = p
				return nil, true
			}
			if owner == "" && relaunch && i < 8 {
				time.Sleep(250 * time.Millisecond)
				continue
			}
			break // gehört jemand anderem – nächsten Port versuchen
		}
		if portOwner(p) == me {
			Port = p
			return nil, true
		}
	}
	return nil, false
}

// Bis Version 1.15.0 lief das Programm fest auf diesem Port (für alle Benutzer gemeinsam).
const legacyPort = 47613
