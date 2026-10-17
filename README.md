# PDF Unterschreiben

Setzt deine Unterschrift automatisch in PDF-Berichte – über dem Feld „Unterschrift Techniker“ (Suchwort frei einstellbar).

## Installieren

1. Unter **Releases** (rechts) die neueste `PDF-Unterschreiben-Setup.exe` herunterladen.
2. Doppelklick. Falls Windows „Der Computer wurde durch Windows geschützt“ anzeigt: **Weitere Informationen → Trotzdem ausführen**.
3. Mit **Ja** bestätigen. Keine Administratorrechte nötig.

Danach liegt „PDF Unterschreiben“ auf dem Desktop und im Startmenü.

## Benutzen

1. Einmalig links deine Unterschrift anlegen: zeichnen, als Bild laden oder aus einer schon unterschriebenen PDF ausschneiden.
2. PDFs oder Excel-Dateien (.xlsx, .xlsm, .xls) ins Fenster ziehen. Excel-Dateien werden mit dem installierten Excel in eine PDF umgewandelt; bei mehreren Blättern fragt das Programm, welches. Die Unterschrift wird automatisch platziert und lässt sich mit der Maus verschieben oder per Doppelklick an eine andere Stelle setzen. Kreise oder Kästchen zum Ankreuzen (z. B. „o Betriebssicher“) werden erkannt und per Klick abgehakt.
3. **Unterschreiben & speichern** – die neue Datei heißt `…_unterschrieben.pdf`. Bei Excel-Dateien kann man wählen: als PDF oder als Excel-Mappe mit eingefügter Unterschrift. Oder **Drucken** (auch mit Strg+P), dann wird die unterschriebene Fassung direkt auf dem gewählten Drucker ausgegeben.

## Updates

Beim Start prüft das Programm, ob es eine neue Version gibt, und zeigt dann oben einen Hinweis mit **Jetzt aktualisieren**. Unterschrift und Einstellungen bleiben dabei erhalten.

## Deinstallieren

Windows-Einstellungen → Apps → „PDF Unterschreiben“ → Deinstallieren.

---

### Neue Version veröffentlichen

Es gibt zwei Kanäle:

1. **Testen:** In `TESTVERSION` eine höhere Nummer eintragen als in `VERSION` (z. B. `1.13.1`) und nach `main` pushen. Jeder weitere Push baut automatisch eine neue Testversion (`1.13.1-test.N`) ins Vorab-Release `adminkanal`. Die bekommen nur Admins, die im Programm unter „Testkanal“ Testversionen eingeschaltet haben.
2. **Freigeben:** In `VERSION` dieselbe Nummer eintragen (z. B. `1.13.1`) und pushen. Dann wird die fertige Version als normales Release für alle veröffentlicht.

Die erste Zeile der Commit-Nachricht erscheint im Programm als Update-Hinweis.
