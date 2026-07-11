# Deployment-Anleitung: Receipt-Worker für Paperless-ngx

Schritt-für-Schritt-Anleitung für die Installation auf einem NAS oder
Homeserver mit Docker (z. B. mit Portainer). Geschätzte Dauer: 45–60 Minuten inkl. Test.

---

## Voraussetzungen

- Paperless-ngx läuft als Docker-Stack (Portainer) auf dem NAS
- SSH-Zugang zum NAS oder Portainer-Zugriff
- Azure-Document-Intelligence-Key (Schritt 2, Default-Provider) – oder ein
  Anthropic API-Key, siehe „Alternative Provider" weiter unten
- Ein NAS-Share für die Excel-Ausgabe, den du vom PC/Handy erreichst

---

## Schritt 1: Vorbereitung – Namen und Pfade ermitteln

Diese drei Werte brauchst du später. Per SSH auf dem NAS:

```bash
# 1a) Container-Name des Paperless-Webservers:
docker ps --format '{{.Names}}\t{{.Image}}' | grep -i paperless
# Beispiel-Ausgabe: paperless-webserver   ghcr.io/paperless-ngx/paperless-ngx:latest

# 1b) Docker-Netzwerk des Paperless-Stacks:
docker inspect <paperless-container-name> --format '{{range $k, $v := .NetworkSettings.Networks}}{{$k}}{{end}}'
# Beispiel-Ausgabe: paperless_default

# 1c) Host-Pfad des Paperless-Stacks (wo docker-compose.yml liegt), z. B.:
#     /volume1/docker/paperless
```

Notiere dir:

| Wert | Dein Eintrag |
|---|---|
| Paperless-Container-Name | ______________ |
| Docker-Netzwerk | ______________ |
| Paperless-Stack-Pfad | ______________ |
| Output-Share für Excel | z. B. `/volume1/shares/abrechnung` |

---

## Schritt 2: API-Zugänge anlegen

**Azure Document Intelligence Key (Default-Provider):**

1. Im Azure-Portal eine „Document Intelligence"-Ressource anlegen (Region
   frei wählbar, Tarif „Free" reicht zum Testen).
2. Unter *Keys and Endpoint* Key 1 sowie den Endpoint notieren. Sicher
   ablegen (z. B. Vaultwarden).

**Paperless API-Token:**

1. Paperless-Web-UI öffnen → Profilsymbol oben rechts → **API-Token**
   (bzw. *Mein Profil* → Token generieren).
2. Token kopieren.

**Optional – Anthropic API-Key** (nur nötig bei `PROVIDER=anthropic`, siehe
„Alternative Provider" weiter unten):

1. https://console.anthropic.com öffnen, Konto anlegen (unabhängig von der
   Claude-Pro-Subscription).
2. Unter *Billing* ein kleines Prepaid-Guthaben aufladen (5 € reichen für
   sehr lange – ein Bon kostet deutlich unter 1 Cent).
3. Unter *API Keys* einen Key erzeugen, beginnt mit `sk-ant-`. Sicher
   ablegen (z. B. Vaultwarden), er wird nur einmal angezeigt.

---

## Schritt 3: Paperless vorbereiten (Web-UI, 5 Minuten)

1. **Tags anlegen:** *Verwaltung → Tags → Neu* (Namen frei wählbar, müssen
   dann in Schritt 6 identisch eingetragen werden):
   - `Auslagen` – Trigger-Tag, löst die Verarbeitung überhaupt erst aus.
   - Je Person, die abgerechnet wird, ein eigener Tag, z. B. `Rose`. Ein
     Beleg mit `Auslagen` + `Rose` landet in der Excel-Datei für Rose; ohne
     passenden Personen-Tag in der gemeinsamen Standarddatei; mit mehreren
     passenden Personen-Tags in jeder betroffenen Datei.
   - `Verarbeitet` wird vom Worker automatisch angelegt und gesetzt – hier
     nichts zu tun, außer du willst ihn schon vorab anlegen.
2. Optional, wenn der Scanner nicht selbst taggt: *Verwaltung → Workflows* →
   je Tag ein Workflow mit Trigger „Konsumierung gestartet" und Filter auf
   einen dedizierten Consume-Unterordner (z. B. `consume/rose/`), Aktion:
   entsprechende Tags zuweisen. Dann steuert der Ablageordner das Tagging.

---

## Schritt 4: Worker-Dateien auf das NAS kopieren

```bash
# Zielordner anlegen
mkdir -p /volume1/docker/receipt-worker

# ZIP-Inhalt dorthin kopieren (per SMB/File-Station oder scp), danach:
ls /volume1/docker/receipt-worker
# Erwartete Dateien: worker.py  Dockerfile  docker-compose.yml
#                    post-consume.sh  .env.example  README.md  DEPLOYMENT.md

# Output-Ordner anlegen:
mkdir -p /volume1/shares/abrechnung
```

---

## Schritt 5: Worker-Stack konfigurieren und deployen

### 5a) `docker-compose.yml` anpassen

Öffne `/volume1/docker/receipt-worker/docker-compose.yml` und trage die
Werte aus Schritt 1 sowie deine Personen-Tags ein:

```yaml
PAPERLESS_URL: "http://<paperless-container-name>:8000"
```

```yaml
RECEIPT_TRIGGER_TAG: "Auslagen"
RECEIPT_PERSON_TAGS: "Rose"        # kommasepariert, z. B. "Rose,Anna"
RECEIPT_PROCESSED_TAG: "Verarbeitet"
```

```yaml
AZURE_DI_ENDPOINT: "https://<dein-resource-name>.cognitiveservices.azure.com"
AZURE_DI_LOCALE: "de-DE"    # Sprach-Hint (Standard); leer = Auto-Erkennung
```

Optional, nur falls Zahlen/Kleingedrucktes trotz `de-DE`-Hint falsch
gelesen werden (kostenpflichtiges Azure-Add-on, wird pro Seite extra
berechnet):

```yaml
AZURE_DI_FEATURES: "ocrHighResolution"
```

```yaml
volumes:
  - /volume1/shares/abrechnung:/output      # dein Output-Share
```

```yaml
networks:
  paperless:
    external: true
    name: <dein-netzwerkname>               # z. B. paperless_default
```

### 5b) Secrets hinterlegen

```bash
cd /volume1/docker/receipt-worker
cp .env.example .env
nano .env    # PAPERLESS_TOKEN und AZURE_DI_KEY eintragen
chmod 600 .env
```

(Alternativ in Portainer: Stack-Editor → *Environment variables* →
`PAPERLESS_TOKEN` und `AZURE_DI_KEY` dort setzen, dann keine
`.env` nötig.)

### 5c) Deployen

Die Compose nutzt standardmäßig das fertige Image aus der GitHub
Container Registry – kein eigener Build nötig:

```bash
docker pull ghcr.io/hanborns/paperless-auslagen:latest
```

**Variante SSH:**

```bash
cd /volume1/docker/receipt-worker
docker compose pull && docker compose up -d
```

**Variante Portainer:** *Stacks → Add Stack* → Compose-Inhalt einfügen,
Env-Variablen setzen, deployen. Kein lokaler Build nötig.

**Updates einspielen:** `docker compose pull && docker compose up -d`
(oder in Portainer „Re-pull image and redeploy").

**Hinweis für private Repositories:** Ist das GitHub-Repo privat, ist auch
das GHCR-Image privat – dann entweder das Package in den GitHub-Settings
auf *public* stellen (das Image enthält keinerlei Secrets, alles Sensible
kommt zur Laufzeit per Environment) oder auf dem NAS einmalig per
`docker login ghcr.io` mit einem fein granulierten Token (nur Scope
`read:packages`) anmelden. Alternativ bleibt der lokale Build:
in der Compose `image:` auskommentieren, `build: .` aktivieren,
`docker compose up -d --build`.

### 5d) Funktionstest

```bash
# Läuft der Container?
docker ps | grep receipt-worker

# Ist er aus dem Paperless-Container erreichbar? (wichtig: gleiches Netz)
docker exec <paperless-container-name> curl -s http://receipt-worker:5000/health
# Erwartet: {"provider":"azure_di","status":"ok"}
```

Wenn `curl` im Paperless-Container fehlt oder der Health-Check fehlschlägt:
Netzwerkname in der Compose prüfen (häufigste Fehlerquelle).

---

## Schritt 6: Post-Consume-Hook in Paperless einhängen

### 6a) Script bereitstellen

```bash
mkdir -p /volume1/docker/paperless/scripts
cp /volume1/docker/receipt-worker/post-consume.sh /volume1/docker/paperless/scripts/
chmod +x /volume1/docker/paperless/scripts/post-consume.sh
```

### 6b) Paperless-Stack erweitern

In der `docker-compose.yml` des **Paperless**-Stacks beim Webserver-Service
ergänzen:

```yaml
services:
  webserver:                # Name ggf. abweichend
    # ... bestehende Konfiguration ...
    volumes:
      # ... bestehende Volumes ...
      - /volume1/docker/paperless/scripts:/scripts:ro
    environment:
      # ... bestehende Variablen ...
      PAPERLESS_POST_CONSUME_SCRIPT: /scripts/post-consume.sh
      RECEIPT_TRIGGER_TAG: Auslagen
      RECEIPT_WORKER_URL: http://receipt-worker:5000/process
```

`RECEIPT_TRIGGER_TAG` muss hier denselben Wert haben wie im
Worker-Stack (Schritt 5a) – das Script nutzt ihn nur als günstigen
Vorfilter, bevor es den Worker überhaupt anspricht. Personen-Tags und der
Verarbeitet-Tag werden nur vom Worker gebraucht, hier ist nichts weiter
einzutragen.

### 6c) Paperless neu deployen

```bash
cd <paperless-stack-pfad>
docker compose up -d
```

**Wichtig:** Ein reiner `docker restart` übernimmt keine neuen
Environment-Variablen – es muss ein Recreate über Compose/Portainer sein.

Kontrolle:

```bash
docker exec <paperless-container-name> printenv PAPERLESS_POST_CONSUME_SCRIPT
# Erwartet: /scripts/post-consume.sh
```

---

## Schritt 7: End-to-End-Test

1. Einen echten Bon scannen und in Paperless mit den Tags `Auslagen` +
   deinem Personen-Tag (z. B. `Rose`) konsumieren lassen.
2. Log beobachten:

   ```bash
   docker logs -f receipt-worker
   ```

   Erwartete Zeilen: `Verarbeite Paperless-Dokument …` →
   `Fertig: <blattname> (<person>) | Auslagenbetrag x,xx € …`
3. In Paperless prüfen, dass der Tag `Verarbeitet` gesetzt wurde.
4. `Auslagen_offen_Rose.xlsx` (bzw. der Dateiname für deine Person) im
   Output-Share öffnen und prüfen:
   - Summenblatt vorn mit offenem Gesamtbetrag
   - Beleg-Blatt mit Positionen; gelbe Spalte „Anteil abrechnen" ist
     standardmäßig die volle Menge – bei Bedarf reduzieren, Beträge rechnen
     sich per Formel neu
   - Kontrollzeilen unten: extrahierte Positionssumme bzw. Abrechnungsbetrag
     vs. gedruckte Bon-Summe (grün = plausibel, rot + roter Blatt-Reiter =
     prüfen)

Beim ersten Beleg alle Positionen einmal gegen den Papier-Bon prüfen, um
ein Gefühl für die Erkennungsqualität zu bekommen.

---

## Schritt 8 (optional): Benachrichtigung

### ntfy.sh (Default-Kanal)

1. [ntfy-App](https://ntfy.sh/app) auf dem Handy installieren (oder Web-App
   nutzen) und ein eigenes, unerratbares Thema abonnieren (z. B.
   `paperless-auslagen-x7k2m9` – jeder, der das Thema kennt, kann mitlesen,
   da `ntfy.sh` standardmäßig öffentlich ist).
2. Im Worker-Stack ergänzen und Container neu starten:

   ```yaml
   NTFY_TOPIC: "paperless-auslagen-x7k2m9"
   ```

3. Optional bei selbst gehostetem ntfy-Server: `NTFY_URL` (statt
   `https://ntfy.sh`) und `NTFY_TOKEN` (falls Auth verlangt wird) setzen.

Der Worker schickt den reinen Nachrichtentext (`message`), inkl. ⚠-Hinweisen
bei Summen-Differenzen oder Fehlern (dann mit Priorität „urgent").

### Alternative/Ergänzung: Home-Assistant-Webhook

1. In HA eine Automation anlegen, Trigger-Typ **Webhook**, eigene ID
   vergeben (z. B. `beleg-erfasst`), Methode POST, nur lokales Netzwerk.
2. Action, z. B.:

   ```yaml
   action: notify.mobile_app_dein_handy
   data:
     title: "Einkaufs-Abrechnung"
     message: "{{ trigger.json.message }}"
   ```

3. Im Worker-Stack ergänzen und Container neu starten (zusätzlich zu oder
   statt `NTFY_TOPIC`):

   ```yaml
   HA_WEBHOOK_URL: "http://homeassistant.local:8123/api/webhook/beleg-erfasst"
   ```

Der Webhook liefert zusätzlich `person`, `betrag_auslagen`, `plausibel`,
`unsichere_positionen`, `sheet` und `doc_id` – nutzbar für ein
Dashboard-Kärtchen oder bedingte Warn-Notifications
(`{{ trigger.json.plausibel == false }}`).

---

## Betrieb

**Abrechnungsrunde abschließen:** Nach der Überweisung die passende
`Auslagen_offen[_<Person>].xlsx` umbenennen (z. B.
`Auslagen_2026-07_Rose_beglichen.xlsx`). Der nächste Beleg für diese
Person legt automatisch eine neue Datei an.

**Fehlerhafter Beleg / erneut verarbeiten:** In Paperless beim Dokument den
Tag `Verarbeitet` entfernen. Der nächste periodische Scan (Standard alle
30 Minuten, `SCAN_INTERVAL_MINUTES`) holt es automatisch nach – oder sofort
per Hand:

```bash
# Option 1: Backlog-Scan sofort anstoßen (verarbeitet alle offenen Belege)
docker exec <paperless-container-name> \
  curl -s -X POST http://receipt-worker:5000/scan

# Option 2: gezielt ein einzelnes Dokument
docker exec <paperless-container-name> \
  curl -s -X POST http://receipt-worker:5000/process \
  -H 'Content-Type: application/json' -d '{"document_id": 4711}'
```

**Logs:** `docker logs receipt-worker` oder `receipt-worker.log` im
Output-Share.

**Excel nicht dauerhaft geöffnet lassen** – der Worker schreibt direkt in
die Datei; kollidiert ein Schreibvorgang, steht der Fehler im Log und der
Beleg kann wie oben erneut angestoßen werden.

---

## Alternative Provider

Standard ist **Azure Document Intelligence** (s. Schritt 2 und 5) – ein
spezialisiertes Kassenbon-OCR-Modell statt eines allgemeinen Vision-Modells.

### Anthropic Claude (Option B)

Sinnvoll, wenn Azure DI bei einem bestimmten Bon-Layout Erkennungsprobleme
hat.

1. Anthropic API-Key anlegen, siehe „Optional – Anthropic API-Key" in
   Schritt 2.
2. In der Worker-Compose:

   ```yaml
   PROVIDER: "anthropic"
   ANTHROPIC_API_KEY: "${ANTHROPIC_API_KEY}"
   ANTHROPIC_MODEL: "claude-sonnet-4-6"
   ```

3. `ANTHROPIC_API_KEY` in `.env` eintragen, `docker compose up -d`.

### Lokales Modell via Ollama (Option C)

1. Ollama als Container auf dem NAS deployen (im selben Netz oder mit
   erreichbarer Adresse), Modell ziehen: `ollama pull qwen2.5vl:7b`
2. In der Worker-Compose die drei Provider-Zeilen tauschen:

   ```yaml
   PROVIDER: "openai"
   OPENAI_BASE_URL: "http://ollama:11434"
   OPENAI_MODEL: "qwen2.5vl:7b"
   ```

3. `docker compose up -d` – mehr ändert sich nicht.

**Qualitätsvergleich zwischen Providern:** Denselben Beleg einmal pro
Provider verarbeiten (in Paperless vorher den Tag `Verarbeitet` entfernen,
sonst wird er als Duplikat übersprungen) und die Blätter gegen den
Papier-Bon vergleichen. Rechne bei reiner CPU-Ollama-Inferenz mit mehreren
Minuten pro Beleg und voller CPU-Last während der Verarbeitung.

---

## Fehlerbehebung

| Symptom | Ursache / Lösung |
|---|---|
| Health-Check aus Paperless-Container schlägt fehl | Falscher Netzwerkname in Worker-Compose; `docker network ls` prüfen |
| Kein Log-Eintrag nach Consume | Tag stimmt nicht mit `RECEIPT_TRIGGER_TAG` überein; Env-Variablen im Paperless-Container prüfen (`printenv`); Recreate statt Restart durchgeführt? |
| `401` im Worker-Log | `PAPERLESS_TOKEN` oder Provider-API-Key falsch |
| `Permission denied` beim Script | `chmod +x post-consume.sh` fehlt |
| Excel-Blatt fehlt trotz Log „Fertig" | Falscher Output-Mount; `docker inspect receipt-worker` → Mounts prüfen |
| Summen-Differenz auf jedem Beleg | Scan-Qualität prüfen (Auflösung ≥ 200 dpi, Bon glatt auflegen) |
| Dokument bleibt trotz Tag unverarbeitet | Personen-Tag nicht in `RECEIPT_PERSON_TAGS`? Läuft bereits eine andere Verarbeitung (Lock-Datei `.receipt_worker.lock` im Output-Ordner)? |
