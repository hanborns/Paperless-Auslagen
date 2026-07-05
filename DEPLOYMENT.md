# Deployment-Anleitung: Receipt-Worker für Paperless-ngx

Schritt-für-Schritt-Anleitung für die Installation auf einem NAS oder
Homeserver mit Docker (z. B. mit Portainer). Geschätzte Dauer: 45–60 Minuten inkl. Test.

---

## Voraussetzungen

- Paperless-ngx läuft als Docker-Stack (Portainer) auf dem NAS
- SSH-Zugang zum NAS oder Portainer-Zugriff
- Anthropic API-Key (Schritt 2)
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

**Anthropic API-Key:**

1. https://console.anthropic.com öffnen, Konto anlegen (unabhängig von der
   Claude-Pro-Subscription).
2. Unter *Billing* ein kleines Prepaid-Guthaben aufladen (5 € reichen für
   sehr lange – ein Bon kostet deutlich unter 1 Cent).
3. Unter *API Keys* einen Key erzeugen, beginnt mit `sk-ant-`. Sicher
   ablegen (z. B. Vaultwarden), er wird nur einmal angezeigt.

**Paperless API-Token:**

1. Paperless-Web-UI öffnen → Profilsymbol oben rechts → **API-Token**
   (bzw. *Mein Profil* → Token generieren).
2. Token kopieren.

---

## Schritt 3: Paperless vorbereiten (Web-UI, 5 Minuten)

1. **Zwei Tags anlegen:** *Verwaltung → Tags → Neu* (Kleinschreibung;
   andere Namen gehen, müssen dann in Schritt 6 identisch eingetragen
   werden):
   - `abrechnung` → nur die handschriftlich markierten Positionen
     („R" = ganze Menge, Zahl = Anteil) gehen in die Abrechnung
   - `abrechnung-komplett` → der **gesamte Beleg** geht in die
     Abrechnung, keine Markierungen nötig (hat Vorrang, falls beide
     Tags gesetzt sind)
2. Optional, wenn der Scanner nicht selbst taggt: *Verwaltung → Workflows* →
   je Tag ein Workflow mit Trigger „Konsumierung gestartet" und Filter auf
   einen dedizierten Consume-Unterordner (z. B. `consume/abrechnung/`
   und `consume/abrechnung-komplett/`), Aktion: entsprechenden Tag
   zuweisen. Dann steuert der Ablageordner das Tagging.

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
Werte aus Schritt 1 ein:

```yaml
PAPERLESS_URL: "http://<paperless-container-name>:8000"
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
nano .env    # PAPERLESS_TOKEN und ANTHROPIC_API_KEY eintragen
chmod 600 .env
```

(Alternativ in Portainer: Stack-Editor → *Environment variables* →
`PAPERLESS_TOKEN` und `ANTHROPIC_API_KEY` dort setzen, dann keine
`.env` nötig.)

### 5c) Deployen

Die Compose nutzt standardmäßig das fertige Image aus der GitHub
Container Registry (`ghcr.io/hanborns/paperless-auslagen:latest`).

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
# Erwartet: {"provider":"anthropic","status":"ok"}
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
      RECEIPT_TRIGGER_TAG: abrechnung
      RECEIPT_FULL_TAG: abrechnung-komplett
      RECEIPT_WORKER_URL: http://receipt-worker:5000/process
```

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

1. Einen echten Bon mit der Markierungs-Konvention versehen:
   - **R** neben Position = komplette Menge für Abrechnungspartei
   - **Zahl** neben Position = dieser Anteil für Abrechnungspartei
   - **nichts** = gehört dir
2. Scannen und in Paperless mit Tag `abrechnung` konsumieren lassen.
   Danach zusätzlich einen unmarkierten Bon mit Tag
   `abrechnung-komplett` testen – dessen „Abrechnungsbetrag" muss der
   Bon-Endsumme entsprechen.
3. Log beobachten:

   ```bash
   docker logs -f receipt-worker
   ```

   Erwartete Zeilen: `Verarbeite Paperless-Dokument …` →
   `Fertig: <blattname> | Anteil x,xx € | offen gesamt …`
4. `Abrechnung_offen.xlsx` im Output-Share öffnen und prüfen:
   - Summenblatt vorn mit offenem Gesamtbetrag
   - Beleg-Blatt mit Positionen; gelbe Spalte „Anteil" = gelesene
     Handschrift-Werte (bei Bedarf korrigieren, Beträge rechnen sich neu)
   - Kontrollzeile unten: extrahierte Positionssumme vs. gedruckte
     Bon-Summe (grün = plausibel, rot + roter Blatt-Reiter = prüfen)

Beim ersten Beleg alle Positionen einmal gegen den Papier-Bon prüfen, um
ein Gefühl für die Lesequalität deiner Handschrift zu bekommen.

---

## Schritt 8 (optional): Home-Assistant-Benachrichtigung

1. In HA eine Automation anlegen, Trigger-Typ **Webhook**, eigene ID
   vergeben (z. B. `beleg-erfasst`), Methode POST, nur lokales Netzwerk.
2. Action, z. B.:

   ```yaml
   action: notify.mobile_app_dein_handy
   data:
     title: "Einkaufs-Abrechnung"
     message: "{{ trigger.json.message }}"
   ```

3. Im Worker-Stack ergänzen und Container neu starten:

   ```yaml
   HA_WEBHOOK_URL: "http://homeassistant.local:8123/api/webhook/beleg-erfasst"
   ```

Der Webhook liefert zusätzlich `betrag_se`, `offen_gesamt`, `plausibel`,
`unsichere_positionen`, `sheet` und `doc_id` – nutzbar für ein
Dashboard-Kärtchen oder bedingte Warn-Notifications
(`{{ trigger.json.plausibel == false }}`).

---

## Betrieb

**Abrechnungsrunde abschließen:** Nach der Überweisung
`Abrechnung_offen.xlsx` umbenennen (z. B.
`Abrechnung_2026-07_beglichen.xlsx`). Der nächste Beleg legt
automatisch eine neue Datei an; der offene Zähler startet bei 0.

**Fehlerhafter Beleg / erneut verarbeiten:** Im Output-Share liegt
`.receipt_worker_state.json`. Die Dokument-ID aus `processed_ids`
entfernen, dann das Dokument in Paperless erneut konsumieren (oder den
Worker direkt anstoßen):

```bash
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

## Umschalten auf lokales Modell (Option C)

1. Ollama als Container auf dem NAS deployen (im selben Netz oder mit
   erreichbarer Adresse), Modell ziehen: `ollama pull qwen2.5vl:7b`
2. In der Worker-Compose die drei Provider-Zeilen tauschen:

   ```yaml
   PROVIDER: "openai"
   OPENAI_BASE_URL: "http://ollama:11434"
   OPENAI_MODEL: "qwen2.5vl:7b"
   ```

3. `docker compose up -d` – mehr ändert sich nicht.

**Qualitätsvergleich:** Denselben Beleg einmal pro Provider verarbeiten
(ID vorher aus `processed_ids` entfernen) und die beiden Blätter gegen den
Papier-Bon vergleichen. Rechne bei reiner CPU-Inferenz mit
mehreren Minuten pro Beleg und voller CPU-Last während der Verarbeitung.

---

## Fehlerbehebung

| Symptom | Ursache / Lösung |
|---|---|
| Health-Check aus Paperless-Container schlägt fehl | Falscher Netzwerkname in Worker-Compose; `docker network ls` prüfen |
| Kein Log-Eintrag nach Consume | Tag stimmt nicht mit `RECEIPT_TRIGGER_TAG`/`RECEIPT_FULL_TAG` überein; Env-Variablen im Paperless-Container prüfen (`printenv`); Recreate statt Restart durchgeführt? |
| `401` im Worker-Log | `PAPERLESS_TOKEN` oder `ANTHROPIC_API_KEY` falsch |
| `Permission denied` beim Script | `chmod +x post-consume.sh` fehlt |
| Excel-Blatt fehlt trotz Log „Fertig" | Falscher Output-Mount; `docker inspect receipt-worker` → Mounts prüfen |
| Summen-Differenz auf jedem Beleg | Scan-Qualität prüfen (Auflösung ≥ 200 dpi, Bon glatt auflegen) |
