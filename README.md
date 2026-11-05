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

## Vertrieb und Auftragsabwicklung (ersetzt die Monday-Boards)

Ein Projekt hat im Add-in immer **genau eine Statuskarte**:

- **Vor dem Auftrag: Vertriebsstatus** (Offen, Angefragt (Lieferant), Angebot schreiben, Kontrolle (Preise), Raus an Kunden,
  In Verhandlung, On Hold, Gewonnen, Verloren, Übergeben) mit Closing Chance, Priorität und
  „Entscheidung bis“. Ist der Termin überschritten, erscheint ein Hinweis mit „Nachgefasst: +7 Tage“ –
  der eigentliche Status bleibt erhalten (in Monday hat „Überfällig“ ihn überschrieben).
- **Gewonnen → „An Auftragsabwicklung übergeben“** legt den Auftrag an; die Karte wechselt zum
  **Auftragsstatus** (AB, US Dollar, Korrekturabzug, Shippingmarks, Freigabemuster, Produktion,
  Frachtweg, Lieferstatus, Lieferung und Termine).
- **„Übersicht“** (Button oben im Add-in) öffnet `board.html` mit zwei Ansichten im gleichen Look:
  *Anfrage & Vertrieb* → *Auftragsabwicklung*. Gruppen werden aus den Status berechnet,
  Status lassen sich direkt ändern, Summen und Filter (Verkäufer/PM, Standort).
- **Regeln:** `src/sales.js` (V1–V7, T1–T2, ersetzt 22 Automationen) und `src/orders.js`
  (R1–R7, T1–T5, ersetzt 20 Automationen). Live laufen sie als Server-Abläufe in Elwis.
- **PID als durchgehende Nummer:** Beide Boards und das Add-in führen über die PID (Project ID), von der Anfrage bis zur Rechnung; AN-Nummern entfallen.
- **Felder:** „Regeln & Felder“ im Board zeigt, was in Elwis schon existiert und was neu anzulegen ist.
- In der Demo werden Änderungen nur im Browser gespeichert („Demo zurücksetzen“ im Board).

## Kalkulator (OEM-Kalkulation)

Der Kalkulator ersetzt die Excel-Vorlage „Automatic Calculation V15“ und liegt unter `kalkulator/`
(statischer Build der Browser-App *fresh-id-calculator*).

- **Aus der Vertriebskarte:** „Kalkulieren“ öffnet den Kalkulator mit Projekt, PID und Kunde.
  „Ergebnis an Vertrieb übergeben“ schreibt je Staffel Menge, Selbstkosten, VK und DB an den Deal:
  Die Karte zeigt die Staffeln, das Board markiert den Deal als „kalkuliert“ (Regel V7).
  Beim nächsten Öffnen wird der gespeicherte Stand geladen. Bei „Kontrolle (Preise)“ ohne Kalkulation erscheint ein Hinweis.
- **Im Board:** Knopf „Kalkulator“ oben, „∑“ je Deal.
- **Rechenkern:** zeilengetreue Portierung der Excel-Vorlage, gegen die Vorlage geprüft (33 echte Staffeln,
  1.360 Zufallsfälle). Rechenfehler der Vorlage sind korrigiert (K1–K13); weicht ein Ergebnis deshalb von Excel ab,
  zeigt der Kalkulator den Excel-Wert und den Grund.
- **Live (Stufe 2):** Übergabe als Angebotspositionen am Elwis-Angebot statt im Browser-Speicher.
- Neu bauen: im Kalkulator-Projekt `npx vite build --config vite.demo.config.ts`, dann `dist-demo/` nach `kalkulator/` kopieren.

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
| `src/sales.js` | Vertriebsstatus: Felder, Farben, Phasen, Regeln |
| `src/orders.js` | Auftragsstatus: Felder, Farben, Phasen, Regeln |
| `src/refs.js` | Erkennung von PID-, SC- und Belegnummern im Betreff (alle gelebten Schreibweisen) |
| `src/status-ui.js`, `src/status-ui.css` | Status-Chips und Auswahl (Add-in und Board) |
| `board.html`, `src/board.js`, `src/board.css` | Übersicht: Anfrage & Vertrieb, Auftragsabwicklung |
| `src/kalk.js` | Verbindung Vertrieb ↔ Kalkulator (Öffnen mit Deal, Übergabe der Staffeln) |
| `kalkulator/` | Kalkulator (Build), Tariftabellen unter `kalkulator/data/` |
| `backend/` | Vermittlungs-Backend für die Live-Version (nur lesend, noch nicht in Betrieb) |

## Live-Version (Stufe 1, nur lesend)

1. Backend aus `backend/` auf einem Server mit Node 20 betreiben
   (`ELWIS_BASE_URL`, `ALLOWED_ORIGIN` setzen, nur über HTTPS erreichbar machen).
2. In `src/config.js` `mode: 'live'` und `backendUrl` setzen.
3. Jeder Benutzer meldet sich im Add-in einmal mit seinem eigenen Elwis-Zugang an; in Elwis gelten
   weiter seine Rechte. Das Backend speichert kein Passwort, nur die Elwis-Sitzung im Arbeitsspeicher.
4. Das Backend blockiert technisch alle schreibenden Elwis-Aufrufe; erlaubt sind nur Lese- und
   Abfrage-Endpunkte.
5. Endpunkte: `POST /api/session`, `POST /api/lookup` (Absender, PID, SC-Nummer, Belegnummer),
   `GET /api/orders`, `GET /api/deals`. Statusänderungen (`PATCH /api/orders|deals/:id`) antworten
   bewusst mit 501, bis das Schreiben freigegeben ist.

Im Elwis-System muss dafür nichts geändert werden.
