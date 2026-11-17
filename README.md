# Receipt-Worker: Geteilte Einkaufsbelege aus Paperless-ngx abrechnen

Automatisiert die Aufteilung gemeinsamer Einkaufsbelege: Bons werden nach dem
Scannen per Vision-LLM (oder Azure Document Intelligence) vollständig
ausgelesen und in Excel-Dateien mit Summenblatt übertragen – eine Datei pro
Person, falls mehrere Personen abgerechnet werden.

## Ablauf im Alltag

1. **Scannen** wie gewohnt und in Paperless taggen (manuell, per
   Scanner-Profil oder Paperless-Workflow):
   - Tag `Auslagen` (Name konfigurierbar über `RECEIPT_TRIGGER_TAG`) –
     löst die Verarbeitung überhaupt erst aus.
   - Zusätzlich ein Personen-Tag (z. B. `Rose`), falls der Beleg einer
     bestimmten Person zugeordnet werden soll. Ohne passenden Personen-Tag
     landet der Beleg in der gemeinsamen Standarddatei. Mit mehreren
     passenden Personen-Tags landet er in jeder betroffenen Datei.
2. **Fertig.** Der Worker extrahiert alle Positionen vollständig (Anteil =
   Menge) und legt in `Auslagen_offen[_<Person>].xlsx` ein neues Blatt an;
   das Summenblatt zeigt den offenen Gesamtbetrag.
3. **Nur ein Teil sollte abgerechnet werden?** Direkt in der gelben Spalte
   „Anteil abrechnen" im Excel-Blatt korrigieren – der Betrag rechnet sich
   per Formel neu. Positionen, die grundsätzlich nie abgerechnet werden
   (z. B. ein Extra-Belag nur für dich), trägst du in
   `RECEIPT_EXCLUDE_ITEMS` ein (kommasepariert, Teiltext ohne
   Groß-/Kleinschreibung, z. B. `"Extra Bacon,Extra Speck"`): Sie bekommen
   automatisch Anteil 0 und den Hinweis „➖ nicht abgerechnet"; die
   Kontrollzeile rechnet sie aus der Bon-Summe heraus. Begriffe spezifisch
   wählen – `Bacon` allein träfe auch einen „Bacon King".
4. **Nach der Überweisung:** Datei umbenennen (z. B.
   `Auslagen_2026-07_beglichen.xlsx`). Beim nächsten Beleg für diese
   Person wird automatisch eine frische Datei mit Summenblatt angelegt.

Rote Blatt-Reiter bedeuten: extrahierte Positionssumme weicht von der
gedruckten Bon-Summe ab → Blatt kurz gegenprüfen.

**Plausibilitätskontrolle (Azure DI):** Vor dem Schreiben prüft der Worker,
ob die erkannten Positionen zur Bon-Summe passen und jede Zeile
Menge × Einzelpreis = Zeilensumme erfüllt. Passt es nicht, probiert er
typische Erkennungsfehler durch – z. B. Menü-Komponenten, die doppelt
gezählt wurden (Burger King: Menü bleibt als Pauschalbetrag, Komponenten
wandern in den Artikelnamen), in einer Zeile vermischte Netto-/Brutto-Beträge
oder als eigene Position erkannte Mengenzeilen („2 x 0,99", Aldi). Eine
Korrektur wird **nur übernommen, wenn danach alles exakt aufgeht**; sie
steht dann blau hinterlegt mit 🔧 und dem Originalwert in der
Hinweis-Spalte. Geht es auch so nicht auf, kann optional eine Zweitmeinung
per Claude eingeholt werden (siehe „Zweitmeinung" unten) – sonst bleibt das
Azure-Ergebnis mit ⚠-Markierung stehen.

**Aufpreis-Extras** (Burger King: „Extra Käse EUR 1,20", „Extra
Bacon@1,20 EUR 2,40") stehen auf dem Bon als eigene Zeile unter der
Hauptposition und sind in deren Zeilensumme nicht enthalten. Der Worker
rechnet sie immer in die Hauptposition darüber ein (Zeilensumme = Bon-Zeile
+ Extras, Einzelpreis neu berechnet, Artikel „Whopper + Extra Käse");
der 🔧-Hinweis nennt die Einzelbeträge und die ursprüngliche Bon-Zeile.
Extras, die auf `RECEIPT_EXCLUDE_ITEMS` passen, bleiben eigene Zeilen
(Anteil 0).

**Zweispaltig gelesene Bons (ALDI):** Azure liest bei manchen Scans die
Betragsspalte getrennt vom Text; Beträge, Stückpreise („6 x 0,89") und
Gewichtsangaben („0,598 kg × 1,99 EUR/kg") landen dann beim Nachbarartikel.
Der Worker schiebt sie zurück, wenn die Rechnung danach aufgeht – auch
anhand von Mengenzeilen, die nur im Rohtext stehen (blasse Stückzahl).
Hat Azure zwei Artikel zu einer Position zusammengezogen, teilt der Worker
sie auf; ein dabei verlorener Betrag kommt aus dem Rohtext oder – wenn er
dort unlesbar ist (z. B. „2,4%") und beide Artikel eine eigene
Artikelnummer haben – aus der Differenz zur Bon-Summe (🔧-Hinweis).
Ist eine Zeile trotzdem inkonsistent (Menge × Preis ≠ Zeilensumme), rechnet
die Betrag-Spalte anteilig mit der gedruckten Zeilensumme statt mit dem
vermutlich falsch gelesenen Preis.

**Kaufdatum:** Stehen mehrere Daten auf dem Bon, gilt das **oberste
plausible** (nicht nach dem Hinzufügen in Paperless, höchstens ein Jahr
davor). TSE-Signaturen („Letzte erfolgreiche Signatur", „TSE-Start" …),
Maschinen-Zeitstempel und Öffnungszeiten zählen nicht; OCR-Varianten wie
`07,11.2026` werden erkannt, ISO-Daten immer als Jahr-Monat-Tag gelesen.
Weicht das Erstellt-Datum in Paperless ab, korrigiert der Worker es, hängt
eine Notiz an und setzt den Tag `Datum prüfen` (`RECEIPT_DATE_CHECK_TAG`);
den Tag bekommen auch Belege mit unsicherem Datum (dann ohne Korrektur).
Abschalten der Korrektur: `PAPERLESS_FIX_CREATED=false`.

**Link zum Beleg:** Jedes Beleg-Blatt hat oben rechts (Spalte H) und in
der Quellzeile einen direkten Link „Paperless #<Nummer>"; das Summenblatt
führt die Dokumentnummer als Link in Spalte E. Damit die Links im Browser
funktionieren, `PAPERLESS_PUBLIC_URL` auf die Adresse setzen, unter der du
Paperless öffnest.

**Verarbeitet-Tag:** Nach erfolgreicher Verarbeitung setzt der Worker den Tag
`Verarbeitet` (konfigurierbar über `RECEIPT_PROCESSED_TAG`) – das ist der
Duplikatschutz. Ihn in Paperless zu entfernen lässt den Beleg beim nächsten
Scan (siehe unten) oder erneuten Consume noch einmal verarbeiten.

**Backlog-Scan:** Läuft periodisch im Hintergrund (`SCAN_INTERVAL_MINUTES`,
Standard 30) und über den Endpunkt `POST /scan` auch manuell auslösbar. Fängt
Belege ab, die zwar den Trigger-Tag tragen, aber – z. B. weil sie per E-Mail
importiert und erst nachträglich getaggt wurden – nie den Post-Consume-Hook
durchlaufen haben.

## Deployment

### 1. Paperless vorbereiten

- Trigger-Tag anlegen (Name frei wählbar, muss zu `RECEIPT_TRIGGER_TAG`
  passen): `Auslagen`.
- Je Person, die abgerechnet wird, einen eigenen Tag anlegen (z. B. `Rose`)
  und in `RECEIPT_PERSON_TAGS` eintragen (kommasepariert).
- API-Token holen: Paperless-Web-UI → Profilsymbol → **API-Token**.

### 2. Worker-Container deployen

Diesen Ordner auf den NAS kopieren (z. B. `/volume1/docker/receipt-worker`),
dann in `docker-compose.yml` anpassen:

- `PAPERLESS_URL`: Service-Name deines Paperless-Webservers im Docker-Netz
  (z. B. `http://paperless-webserver:8000` – Name per `docker ps` prüfen).
- `networks`: externes Netzwerk auf das des Paperless-Stacks setzen
  (`docker network ls`, oft `<stackname>_default`).
- Volume `/output`: NAS-Share, den du vom Handy/PC erreichst.
- `RECEIPT_PERSON_TAGS`: deine Personen-Tags eintragen.
- `AZURE_DI_ENDPOINT`: Endpoint deiner Azure-Document-Intelligence-Ressource
  eintragen.
- `.env.example` → `.env` kopieren und `PAPERLESS_TOKEN` +
  `AZURE_DI_KEY` eintragen (Azure Portal → Document Intelligence → Keys).

Dann: `docker compose pull && docker compose up -d` – die Compose zieht das
fertige Multi-Arch-Image aus der GitHub Container Registry
(`ghcr.io/hanborns/paperless-auslagen`), ein lokaler Build ist nicht nötig
(oder als Portainer-Stack ohne Repository-Build deployen).

Test: `curl http://<nas-ip>:...` entfällt – der Worker lauscht nur im
Docker-Netz. Stattdessen aus dem Paperless-Container:
`docker exec <paperless-container> curl -s http://receipt-worker:5000/health`

### 3. Post-Consume-Hook in Paperless einhängen

1. `post-consume.sh` in einen Ordner legen, der in den **Paperless**-Container
   gemountet wird, z. B.:

   ```yaml
   # im Paperless-Stack, Service webserver:
   volumes:
     - /volume1/docker/paperless/scripts:/scripts:ro
   environment:
     PAPERLESS_POST_CONSUME_SCRIPT: /scripts/post-consume.sh
     RECEIPT_TRIGGER_TAG: Auslagen
     RECEIPT_WORKER_URL: http://receipt-worker:5000/process
   ```

2. `chmod +x post-consume.sh` nicht vergessen.
3. Paperless-Stack **neu deployen** (Recreate – ein Restart übernimmt keine
   neuen Environment-Variablen).

### 4. Testlauf

Einen Bon scannen (oder ein Foto in den Consume-Ordner legen) und die Tags
vergeben. Im Log nachsehen:

```
docker logs -f receipt-worker
# bzw. die Datei receipt-worker.log im Output-Ordner
```

### 5. Optional: Benachrichtigung per ntfy.sh

Default-Kanal ist [ntfy.sh](https://ntfy.sh) – kein eigener Server nötig,
einfach ein (möglichst unerratbares, da öffentlich abonnierbares) Thema
wählen, in der [ntfy-App](https://ntfy.sh/app) abonnieren und in der
`docker-compose.yml` setzen:

```yaml
NTFY_TOPIC: "paperless-auslagen-<zufälliger-string>"
```

Optional: `NTFY_URL` für einen selbst gehosteten ntfy-Server (statt des
öffentlichen `https://ntfy.sh`), `NTFY_TOKEN` falls dieser Auth verlangt.

**Alternative/Ergänzung – Home-Assistant-Webhook:** In HA eine Automation
mit Webhook-Trigger anlegen (Trigger-Typ „Webhook", ID frei wählen), dann
zusätzlich oder statt ntfy setzen:

```yaml
HA_WEBHOOK_URL: "http://homeassistant.local:8123/api/webhook/<deine-id>"
```

Beide Kanäle bekommen dieselbe Nachricht; ntfy nur den Text (`message`),
der HA-Webhook zusätzlich JSON mit `person`, `betrag_auslagen`,
`plausibel` und `unsichere_positionen` – direkt verwendbar in einer
`notify`-Action: `{{ trigger.json.message }}`.

## Alternative Provider

Standard ist **Azure Document Intelligence** (Option A, s. o.) – ein
spezialisiertes Kassenbon-OCR-Modell statt eines allgemeinen Vision-Modells.

**Feintuning für deutsche Bons (nur Option A):**

- `AZURE_DI_LOCALE` (Standard `de-DE`): Sprach-Hint für OCR und Analyse –
  verbessert die Erkennung von Dezimalkommas (`1,99`), deutschen
  Datumsformaten (`dd.mm.yyyy`) und EUR-Beträgen. Leer setzen
  (`AZURE_DI_LOCALE: ""`) aktiviert die automatische Spracherkennung –
  sinnvoll, wenn auch nicht-deutsche Belege verarbeitet werden (ein
  falscher Hint kann die Genauigkeit verschlechtern).
- `AZURE_DI_FEATURES` (kommasepariert, Standard leer): optionale
  Analyse-Add-ons, z. B. `ocrHighResolution` für sehr klein gedruckte
  Bons. **Achtung:** Add-ons berechnet Azure pro Seite extra – nur bei
  hartnäckigen Lesefehlern aktivieren.

- `AZURE_DI_DEBUG_DUMP` (Standard aus): `true` legt pro Beleg die
  Rohantwort von Azure DI und das Originaldokument unter
  `<Output-Ordner>/.azure_raw/<Dokument-ID>.*` ab. Gedacht zum Sammeln
  problematischer Belege, um die Plausibilitätskontrolle zu verbessern
  (Nachstellen lokal mit `python tools/replay.py`). Enthält vollständige
  Belegdaten – nach dem Sammeln wieder ausschalten.

**Zweitmeinung (nur Option A):** `SECOND_OPINION_PROVIDER: "anthropic"`
(plus `ANTHROPIC_API_KEY`, optional `ANTHROPIC_MODEL`) schickt einen Beleg
**nur dann** zusätzlich an Claude, wenn das Azure-Ergebnis auch nach der
automatischen Korrektur nicht aufgeht. Das Claude-Ergebnis wird nur
übernommen, wenn es selbst exakt aufgeht und zur von Azure gelesenen
Bon-Summe passt; im Blatt steht dann „Erkennung: Zweitmeinung …", die
Benachrichtigung enthält 🤖. Fehler bei der Zweitmeinung (z. B. kein
Guthaben) brechen die Verarbeitung nicht ab – dann gilt das
Azure-Ergebnis. `"openai"` nutzt stattdessen die `OPENAI_*`-Einstellungen.

Alternativ:

**Anthropic Claude** (Option B), z. B. wenn Azure DI bei einem bestimmten
Bon-Layout schwächelt:

```yaml
PROVIDER: "anthropic"
ANTHROPIC_API_KEY: "${ANTHROPIC_API_KEY}"
ANTHROPIC_MODEL: "claude-sonnet-4-6"
```

**Lokales Modell via Ollama** (Option C):

```yaml
PROVIDER: "openai"
OPENAI_BASE_URL: "http://ollama:11434"
OPENAI_MODEL: "qwen2.5vl:7b"
```

Container neu starten – sonst ändert sich nichts. Zum Qualitätsvergleich:
denselben Bon einmal mit jedem Provider verarbeiten (in Paperless den Tag
`Verarbeitet` vom Dokument entfernen, sonst wird es als Duplikat
übersprungen).

## Build & Hosting über GitHub

Das Repository enthält einen GitHub-Actions-Workflow
(`.github/workflows/docker-publish.yml`), der bei jedem Push auf `main`
und bei Version-Tags (`v1.0.0` …) automatisch ein Multi-Arch-Image
(amd64 + arm64) baut und in die GitHub Container Registry pusht:
`ghcr.io/hanborns/paperless-auslagen`. Der Workflow nutzt ausschließlich den
eingebauten `GITHUB_TOKEN` – es müssen keine persönlichen Tokens oder
Secrets im Repository hinterlegt werden.

**Fertiges Image direkt ziehen** (kein eigener Build nötig):

```bash
docker pull ghcr.io/hanborns/paperless-auslagen:latest
```

Versionierte Tags (z. B. `:2.1.0`, `:2.1`, `:2`) entstehen automatisch aus
GitHub-Releases; `docker-compose.yml` nutzt standardmäßig `:latest`.

Alle instanzspezifischen Werte (API-Keys, Tokens, Tag-Namen, Dateiname,
ntfy-Thema, Webhook-URL) werden zur Laufzeit per Environment gesetzt und
sind weder im Code noch im Image enthalten. Die Datei `.env` ist per
`.gitignore` vom Commit ausgeschlossen.

Eigene Tag-Namen (z. B. andere Sprachen oder Kürzel) werden nicht im Code
geändert, sondern über `RECEIPT_TRIGGER_TAG` / `RECEIPT_PERSON_TAGS` /
`RECEIPT_PROCESSED_TAG` / `RECEIPT_DATE_CHECK_TAG` konfiguriert; der Excel-Basisdateiname über
`OPEN_FILENAME`.

## Dateien

| Datei | Zweck |
|---|---|
| `worker.py` | Worker: HTTP-Endpoints, Paperless-Zugriff, Extraktion, Excel |
| `plausibility.py` | Plausibilitätskontrolle + automatische Korrektur der Positionen |
| `tests/` | pytest-Suite (läuft im CI vor dem Image-Build) |
| `tools/replay.py` | Plausibilitätskontrolle lokal gegen gespeicherte Azure-Rohdaten nachstellen |
| `Dockerfile` | Container-Image (Python 3.12 + poppler) |
| `docker-compose.yml` | Stack-Definition (GHCR-Image, Build-Fallback) |
| `post-consume.sh` | Hook-Script für den Paperless-Container |
| `.github/workflows/docker-publish.yml` | CI: Image-Build & Push nach GHCR |
| `.env.example` | Vorlage für Secrets (lokal, nie committen) |
| `DEPLOYMENT.md` | Ausführliche Schritt-für-Schritt-Anleitung |
| `LICENSE` | MIT |

## Hinweise

- **Duplikatschutz:** über den Paperless-Tag `Verarbeitet` (siehe oben) –
  keine lokale Datenbank oder State-Datei nötig.
- **Nebenläufigkeit:** eine Lock-Datei im Output-Ordner
  (`.receipt_worker.lock`) stellt sicher, dass immer nur eine Verarbeitung
  gleichzeitig läuft (verhindert doppelte Verarbeitung, wenn Hook und
  Backlog-Scan sich zeitlich überschneiden).
- **Kosten (Option A, Azure Document Intelligence):** Abrechnung pro
  analysiertem Dokument nach Azure-Pricing für `prebuilt-receipt` (aktuelle
  Preise im Azure-Portal); kein Rasterisieren/Bild-Resize nötig. Der
  `AZURE_DI_LOCALE`-Hint ist kostenlos; Add-ons über `AZURE_DI_FEATURES`
  (z. B. `ocrHighResolution`) kosten pro Seite extra.
- **Kosten Zweitmeinung:** fällt nur bei Belegen an, die Azure auch nach
  Auto-Korrektur nicht stimmig liefert – Kosten wie Option B pro solchem Beleg.
- **Kosten (Option B, Anthropic):** Ein Bon ≈ 1 Bild + ~500 Output-Token →
  deutlich unter 1 Cent pro Beleg mit `claude-sonnet-4-6`. Sparsamer:
  `claude-haiku-4-5-20251001`.
- **Mehrseitige Bons:** bis `MAX_PAGES` (Standard 4) Seiten werden übertragen
  (gilt für Anthropic/OpenAI; Azure DI verarbeitet das Originaldokument ohne
  dieses Limit).
- **Excel offen halten:** Die Datei nicht dauerhaft in Excel geöffnet lassen –
  der Worker schreibt direkt hinein; bei gleichzeitigem Zugriff kann das
  Speichern fehlschlagen (steht dann im Log).
