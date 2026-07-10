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
   per Formel neu.
4. **Nach der Überweisung:** Datei umbenennen (z. B.
   `Auslagen_2026-07_beglichen.xlsx`). Beim nächsten Beleg für diese
   Person wird automatisch eine frische Datei mit Summenblatt angelegt.

Rote Blatt-Reiter bedeuten: extrahierte Positionssumme weicht von der
gedruckten Bon-Summe ab → Blatt kurz gegenprüfen.

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

### 5. Optional: Home-Assistant-Benachrichtigung

In HA eine Automation mit Webhook-Trigger anlegen (Trigger-Typ „Webhook",
ID frei wählen), dann in der `docker-compose.yml` setzen:

```yaml
HA_WEBHOOK_URL: "http://homeassistant.local:8123/api/webhook/<deine-id>"
```

Der Worker schickt JSON mit `message`, `person`, `betrag_auslagen`,
`plausibel` und `unsichere_positionen` – direkt verwendbar in einer
`notify`-Action: `{{ trigger.json.message }}`.

## Alternative Provider

Standard ist **Azure Document Intelligence** (Option A, s. o.) – ein
spezialisiertes Kassenbon-OCR-Modell statt eines allgemeinen Vision-Modells.
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

Alle instanzspezifischen Werte (API-Keys, Tokens, Tag-Namen, Dateiname,
Webhook-URL) werden zur Laufzeit per Environment gesetzt und sind weder
im Code noch im Image enthalten. Die Datei `.env` ist per `.gitignore`
vom Commit ausgeschlossen.

Eigene Tag-Namen (z. B. andere Sprachen oder Kürzel) werden nicht im Code
geändert, sondern über `RECEIPT_TRIGGER_TAG` / `RECEIPT_PERSON_TAGS` /
`RECEIPT_PROCESSED_TAG` konfiguriert; der Excel-Basisdateiname über
`OPEN_FILENAME`.

## Dateien

| Datei | Zweck |
|---|---|
| `worker.py` | Worker: HTTP-Endpoints, Paperless-Zugriff, Extraktion, Excel |
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
  Preise im Azure-Portal); kein Rasterisieren/Bild-Resize nötig.
- **Kosten (Option B, Anthropic):** Ein Bon ≈ 1 Bild + ~500 Output-Token →
  deutlich unter 1 Cent pro Beleg mit `claude-sonnet-4-6`. Sparsamer:
  `claude-haiku-4-5-20251001`.
- **Mehrseitige Bons:** bis `MAX_PAGES` (Standard 4) Seiten werden übertragen
  (gilt für Anthropic/OpenAI; Azure DI verarbeitet das Originaldokument ohne
  dieses Limit).
- **Excel offen halten:** Die Datei nicht dauerhaft in Excel geöffnet lassen –
  der Worker schreibt direkt hinein; bei gleichzeitigem Zugriff kann das
  Speichern fehlschlagen (steht dann im Log).
