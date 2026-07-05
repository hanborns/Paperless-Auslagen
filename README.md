# Receipt-Worker: Geteilte Einkaufsbelege aus Paperless-ngx abrechnen

Automatisiert die Aufteilung gemeinsamer Einkaufsbelege: Handschriftlich
markierte Bons werden nach dem Scannen per Vision-LLM ausgelesen und in eine
Excel-Datei mit Summenblatt übertragen.

## Ablauf im Alltag

1. **An der Kasse / zu Hause:** abzurechnende Positionen auf dem Bon
   markieren – **R** = komplette Menge, Zahl daneben (z. B. „4") = dieser
   Anteil, keine Markierung = gehört dir.
2. **Scannen** wie gewohnt, Tag `abrechnung` vergeben (manuell, per
   Scanner-Profil oder Paperless-Workflow). **Sonderfall kompletter
   abzurechnender Einkauf:** stattdessen Tag `abrechnung-komplett`
   vergeben – dann geht der gesamte Beleg ohne jede Markierung in die
   Abrechnung ein.
3. **Fertig.** Der Worker legt in `Abrechnung_offen.xlsx` ein neues Blatt
   an, das Summenblatt zeigt den offenen Gesamtbetrag.
4. **Nach der Überweisung:** Datei umbenennen (z. B.
   `Abrechnung_2026-07_beglichen.xlsx`). Beim nächsten Beleg wird
   automatisch eine frische Datei mit Summenblatt angelegt.

Gelbe Zellen (Spalte „Anteil") sind die vom Modell gelesenen Handschrift-
Werte – bei Bedarf einfach in Excel korrigieren, die Beträge rechnen sich per
Formel neu. Rote Blatt-Reiter bedeuten: extrahierte Positionssumme weicht von
der gedruckten Bon-Summe ab → Blatt kurz gegenprüfen.

## Deployment

### 1. Paperless vorbereiten

- Zwei Tags anlegen (Namen frei wählbar, müssen zu `RECEIPT_TRIGGER_TAG`
  bzw. `RECEIPT_FULL_TAG` passen):
  - `abrechnung` – markierte Positionen werden abgerechnet
  - `abrechnung-komplett` – der gesamte Beleg wird abgerechnet
    (hat Vorrang, falls beide gesetzt sind)
- API-Token holen: Paperless-Web-UI → Profilsymbol → **API-Token**.

### 2. Worker-Container deployen

Diesen Ordner auf den NAS kopieren (z. B. `/volume1/docker/receipt-worker`),
dann in `docker-compose.yml` anpassen:

- `PAPERLESS_URL`: Service-Name deines Paperless-Webservers im Docker-Netz
  (z. B. `http://paperless-webserver:8000` – Name per `docker ps` prüfen).
- `networks`: externes Netzwerk auf das des Paperless-Stacks setzen
  (`docker network ls`, oft `<stackname>_default`).
- Volume `/output`: NAS-Share, den du vom Handy/PC erreichst.
- `.env.example` → `.env` kopieren und `PAPERLESS_TOKEN` +
  `ANTHROPIC_API_KEY` eintragen (API-Key: https://console.anthropic.com).

Dann: `docker compose up -d --build` (oder als Portainer-Stack mit
Build-from-Repository/Upload).

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
     RECEIPT_TRIGGER_TAG: abrechnung
     RECEIPT_FULL_TAG: abrechnung-komplett
     RECEIPT_WORKER_URL: http://receipt-worker:5000/process
   ```

2. `chmod +x post-consume.sh` nicht vergessen.
3. Paperless-Stack **neu deployen** (Recreate – ein Restart übernimmt keine
   neuen Environment-Variablen).

### 4. Testlauf

Einen markierten Bon scannen (oder ein Foto in den Consume-Ordner legen) und
den Tag vergeben. Im Log nachsehen:

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

Der Worker schickt JSON mit `message`, `betrag_se`, `offen_gesamt`,
`plausibel` und `unsichere_positionen` – direkt verwendbar in einer
`notify`-Action: `{{ trigger.json.message }}`.

## Umschalten auf lokales Modell (Option C)

Ollama-Container auf dem NAS starten, Vision-Modell ziehen
(`ollama pull qwen2.5vl:7b`), dann im Worker nur die Provider-Zeilen
umstellen:

```yaml
PROVIDER: "openai"
OPENAI_BASE_URL: "http://ollama:11434"
OPENAI_MODEL: "qwen2.5vl:7b"
```

Container neu starten – sonst ändert sich nichts. Zum Qualitätsvergleich:
denselben Bon einmal mit jedem Provider verarbeiten (Dokument-ID aus dem
State-File `.receipt_worker_state.json` unter `processed_ids` entfernen,
sonst wird er als Duplikat übersprungen).

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
geändert, sondern über `RECEIPT_TRIGGER_TAG` / `RECEIPT_FULL_TAG` im
Paperless-Stack konfiguriert; der Excel-Dateiname über `OPEN_FILENAME`.

## Dateien

| Datei | Zweck |
|---|---|
| `worker.py` | Worker: HTTP-Endpoint, Paperless-Download, LLM-Extraktion, Excel |
| `Dockerfile` | Container-Image (Python 3.12 + poppler) |
| `docker-compose.yml` | Stack-Definition (GHCR-Image, Build-Fallback) |
| `post-consume.sh` | Hook-Script für den Paperless-Container |
| `.github/workflows/docker-publish.yml` | CI: Image-Build & Push nach GHCR |
| `.env.example` | Vorlage für Secrets (lokal, nie committen) |
| `DEPLOYMENT.md` | Ausführliche Schritt-für-Schritt-Anleitung |
| `LICENSE` | MIT |

## Hinweise

- **Duplikatschutz:** Verarbeitete Dokument-IDs stehen im State-File im
  Output-Ordner; ein erneuter Consume desselben Dokuments wird übersprungen.
- **Kosten (Option A):** Ein Bon ≈ 1 Bild + ~500 Output-Token → deutlich
  unter 1 Cent pro Beleg mit `claude-sonnet-4-6`. Sparsamer:
  `claude-haiku-4-5-20251001` (bei Handschrift etwas schwächer).
- **Mehrseitige Bons:** bis `MAX_PAGES` (Standard 4) Seiten werden übertragen.
- **Excel offen halten:** Die Datei nicht dauerhaft in Excel geöffnet lassen –
  der Worker schreibt direkt hinein; bei gleichzeitigem Zugriff kann das
  Speichern fehlschlagen (steht dann im Log, Beleg per State-File-Eintrag
  entfernen und neu consumen).
