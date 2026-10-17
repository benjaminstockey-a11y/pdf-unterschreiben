# PDF Unterschreiben

Setzt deine Unterschrift automatisch in PDF-Berichte – über dem Feld „Unterschrift Techniker“ (Suchwort frei einstellbar).

## Installieren

1. Unter **Releases** (rechts) die neueste `PDF-Unterschreiben-Setup.exe` herunterladen.
2. Doppelklick. Falls Windows „Der Computer wurde durch Windows geschützt“ anzeigt: **Weitere Informationen → Trotzdem ausführen**.
3. Mit **Ja** bestätigen. Keine Administratorrechte nötig.

Danach liegt „PDF Unterschreiben“ auf dem Desktop und im Startmenü.

## Benutzen

1. Einmalig links deine Unterschrift anlegen: zeichnen, als Bild laden oder aus einer schon unterschriebenen PDF ausschneiden.
2. PDFs ins Fenster ziehen. Die Unterschrift wird automatisch platziert und lässt sich mit der Maus verschieben.
3. **Unterschreiben & speichern**. Die neue Datei heißt `…_unterschrieben.pdf`.

## Updates

Beim Start prüft das Programm, ob es eine neue Version gibt, und zeigt dann oben einen Hinweis mit **Jetzt aktualisieren**. Unterschrift und Einstellungen bleiben dabei erhalten.

## Deinstallieren

Windows-Einstellungen → Apps → „PDF Unterschreiben“ → Deinstallieren.

---

### Neue Version veröffentlichen

In der Datei `VERSION` die Nummer erhöhen (z. B. `1.0.1`) und nach `main` pushen. Die erste Zeile der Commit-Nachricht erscheint im Programm als Update-Hinweis.

GitHub Actions baut daraus die `Setup.exe` und das `version.json` und veröffentlicht beides als Release.
