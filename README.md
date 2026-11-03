# Elwis für Outlook (Demo)

Outlook-Add-in, das zur geöffneten E-Mail die passenden Daten aus dem Elwis ERP von fresh id anzeigt:
Kontakt, Kunde/Lieferant, Projekte (Anfragen) mit Projektverlauf, Angebote, Aufträge, Rechnungen,
offene Aufgaben und letzte Notizen. Erkennt Projektnummern wie `PID-3201-NBG` und 10-stellige
Belegnummern im Betreff.

**Stand: Demo mit erfundenen Beispieldaten.** Das Add-in greift in dieser Version nicht auf Elwis zu.
„In Elwis öffnen“ öffnet die passende Liste in Elwis; „Als Notiz speichern“ und „Aufgabe anlegen“
zeigen die geplante Stufe 2, speichern aber nichts.

## Vorschau im Browser

https://freshidbln.github.io/elwis-outlook-addin/taskpane.html – oben lässt sich eine Beispiel-E-Mail wählen.

## Installation im neuen Outlook (ohne Admin)

1. Im Browser **https://aka.ms/olksideload** öffnen (öffnet Outlook im Web mit dem Add-in-Dialog).
2. Links **Meine Add-ins** wählen, ganz unten bei **Benutzerdefinierte Add-ins** auf
   **+ Benutzerdefiniertes Add-in hinzufügen** klicken.
3. **Aus URL hinzufügen** und diese Adresse eintragen:
   `https://freshidbln.github.io/elwis-outlook-addin/manifest.xml`
   (alternativ **Aus Datei hinzufügen** mit der heruntergeladenen `manifest.xml`).
4. Warnhinweis bestätigen. Das Add-in erscheint nach kurzer Zeit auch im neuen Outlook für Windows.
5. Eine E-Mail öffnen → in der Leiste über der E-Mail **Apps** (bzw. „…“) → **Elwis**.
   Mit der Pin-Nadel oben im Bereich bleibt er beim Wechsel zwischen E-Mails offen.

Fehlt der Punkt „Benutzerdefinierte Add-ins“, hat die IT das Hinzufügen gesperrt. Dann kann ein
Microsoft-365-Admin das Add-in im **Microsoft 365 Admin Center → Einstellungen → Integrierte Apps →
Benutzerdefinierte Apps hochladen** mit derselben `manifest.xml` für ausgewählte Benutzer freigeben.

Tipp für die Demo: Echte Absender sind in den Beispieldaten nicht enthalten. Bei unbekanntem Absender
bietet der Bereich Beispiel-E-Mails zum Durchklicken an. Wer sich selbst eine Mail mit dem Betreff
`PID-3201-NBG` schickt, sieht die Projekt-Erkennung.

## Aufbau

| Datei | Inhalt |
|---|---|
| `manifest.xml` | Add-in-Manifest (Lesen + Verfassen, anheftbar) |
| `taskpane.html`, `src/styles.css` | Oberfläche für die Outlook-Seitenleiste |
| `src/app.js` | Logik: E-Mail lesen, Darstellung, Suche, Detailansichten |
| `src/data.js` | Datenschicht: `MockProvider` (Demo) und `ApiProvider` (Live) |
| `src/mock-data.js` | Erfundene Beispieldaten im Aufbau der Elwis-Typen |
| `src/status.js` | Status-Werte und Farben wie in Elwis konfiguriert |
| `src/config.js` | Modus (demo/live), Elwis-Adresse, Seiten für Direktlinks |
| `backend/` | Vermittlungs-Backend für die Live-Version (nur lesend, noch nicht in Betrieb) |

## Live-Version (Stufe 1, nur lesend)

1. Backend aus `backend/` auf einem Server mit Node 20 betreiben
   (`ELWIS_BASE_URL`, `ALLOWED_ORIGIN` setzen, nur über HTTPS erreichbar machen).
2. In `src/config.js` `mode: 'live'` und `backendUrl` setzen.
3. Jeder Benutzer meldet sich im Add-in einmal mit seinem eigenen Elwis-Zugang an; in Elwis gelten
   weiter seine Rechte. Das Backend speichert kein Passwort, nur die Elwis-Sitzung im Arbeitsspeicher.
4. Das Backend blockiert technisch alle schreibenden Elwis-Aufrufe; erlaubt sind nur Lese- und
   Abfrage-Endpunkte.

Im Elwis-System muss dafür nichts geändert werden.
