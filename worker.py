#!/usr/bin/env python3
"""
Receipt-Worker für Paperless-ngx
--------------------------------
Nimmt per HTTP eine Paperless-Dokument-ID entgegen, lädt den Beleg,
extrahiert Positionen + handschriftliche Markierungen per Vision-LLM
(Anthropic API oder OpenAI-kompatibel, z. B. Ollama) und pflegt eine
Excel-Datei mit einem Blatt pro Beleg und einem Summenblatt.

Handschrift-Konvention auf dem Bon:
  - "R" neben Position              -> komplette Menge wird abgerechnet
  - Zahl neben Position (z. B. "4") -> genau dieser Anteil wird abgerechnet
  - keine Markierung                -> Position gehört komplett dem Käufer

Zusätzlich: Trägt das Dokument den Komplett-Tag (z. B. "abrechnung-komplett"),
geht der gesamte Beleg ohne Markierungen in die Abrechnung ein.
"""

import base64
import io
import json
import logging
import os
import re
import threading
import time
from datetime import date, datetime
from pathlib import Path

import requests
from flask import Flask, jsonify, request
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter
from pdf2image import convert_from_bytes
from PIL import Image

# ----------------------------------------------------------------------------
# Konfiguration (per Environment)
# ----------------------------------------------------------------------------
PAPERLESS_URL = os.environ["PAPERLESS_URL"].rstrip("/")   # z. B. http://paperless-webserver:8000
PAPERLESS_TOKEN = os.environ["PAPERLESS_TOKEN"]

PROVIDER = os.environ.get("PROVIDER", "anthropic")        # "anthropic" | "openai" (Ollama etc.)
ANTHROPIC_API_KEY = os.environ.get("ANTHROPIC_API_KEY", "")
ANTHROPIC_MODEL = os.environ.get("ANTHROPIC_MODEL", "claude-sonnet-4-6")
OPENAI_BASE_URL = os.environ.get("OPENAI_BASE_URL", "http://ollama:11434").rstrip("/")
OPENAI_MODEL = os.environ.get("OPENAI_MODEL", "qwen2.5vl:7b")

OUTPUT_DIR = Path(os.environ.get("OUTPUT_DIR", "/output"))
OPEN_FILENAME = os.environ.get("OPEN_FILENAME", "Abrechnung_offen.xlsx")
STATE_FILE = OUTPUT_DIR / ".receipt_worker_state.json"
LOCK_FILE = OUTPUT_DIR / ".receipt_worker.lock"

HA_WEBHOOK_URL = os.environ.get("HA_WEBHOOK_URL", "")     # optional
MAX_PAGES = int(os.environ.get("MAX_PAGES", "4"))
MAX_IMG_WIDTH = int(os.environ.get("MAX_IMG_WIDTH", "1100"))

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s %(levelname)s %(message)s",
    handlers=[
        logging.StreamHandler(),
        logging.FileHandler(OUTPUT_DIR / "receipt-worker.log", encoding="utf-8"),
    ],
)
log = logging.getLogger("receipt-worker")

app = Flask(__name__)
_excel_lock = threading.Lock()

# ----------------------------------------------------------------------------
# Extraktions-Prompt
# ----------------------------------------------------------------------------
PROMPT = """Du siehst einen eingescannten deutschen Supermarkt-Kassenbon mit \
handschriftlichen Markierungen. Extrahiere ALLE Kaufpositionen.

Handschrift-Konvention (Markierungen stehen links oder rechts neben der Zeile):
- Handschriftlicher Buchstabe "R" neben einer Position: die KOMPLETTE Menge \
dieser Position wird abgerechnet.
- Handschriftliche Zahl neben einer Position (z. B. "4"): genau dieser Anteil \
der Menge wird abgerechnet.
- Keine Markierung: Anteil ist 0 (Position wird nicht abgerechnet).
- Achtung: Nur HANDSCHRIFTLICHE Markierungen zählen. Gedruckte Buchstaben auf \
dem Bon (z. B. Steuerkennzeichen "A"/"B" am Zeilenende) sind KEINE Markierungen.

Wichtige Regeln:
- Mengenzeilen wie "6 x 1,15" gehören zur darüber/darunter stehenden Artikelzeile: \
Menge 6, Einzelpreis 1,15.
- Positionen ohne explizite Menge haben Menge 1 und Einzelpreis = Zeilenpreis.
- Rabatt-/Nachlasszeilen als eigene Position mit negativem Preis aufnehmen \
(Anteil-Konvention gilt auch hier).
- Pfandzeilen als normale Positionen behandeln.
- Preise als Dezimalzahl mit Punkt (1.15), niemals mit Komma.
- Lies auch den Markt-/Filialnamen, das Kaufdatum und die gedruckte GESAMTSUMME des Bons.
- Wenn du eine handschriftliche Markierung nicht sicher lesen kannst, setze \
"unsicher": true bei dieser Position.

Antworte AUSSCHLIESSLICH mit einem JSON-Objekt, ohne Markdown, ohne Erklärung:
{
  "markt": "REWE",
  "datum": "2026-07-04",
  "bon_summe": 47.23,
  "positionen": [
    {"artikel": "Frische Vollmilch 1L", "menge": 6, "einzelpreis": 1.15,
     "zeilensumme": 6.90, "anteil": 4, "unsicher": false}
  ]
}
"anteil" ist die Menge (Stückzahl), die abgerechnet wird (0 bis menge).
Falls Datum oder Bon-Summe nicht lesbar: null."""

PROMPT_FULL = """Du siehst einen eingescannten deutschen Supermarkt-Kassenbon. \
Extrahiere ALLE Kaufpositionen. Handschriftliche Markierungen sind zu IGNORIEREN, \
der komplette Beleg wird abgerechnet.

Wichtige Regeln:
- Mengenzeilen wie "6 x 1,15" gehören zur darüber/darunter stehenden Artikelzeile: \
Menge 6, Einzelpreis 1,15.
- Positionen ohne explizite Menge haben Menge 1 und Einzelpreis = Zeilenpreis.
- Rabatt-/Nachlasszeilen als eigene Position mit negativem Preis aufnehmen.
- Pfandzeilen als normale Positionen behandeln.
- Preise als Dezimalzahl mit Punkt (1.15), niemals mit Komma.
- Lies auch den Markt-/Filialnamen, das Kaufdatum und die gedruckte GESAMTSUMME des Bons.

Antworte AUSSCHLIESSLICH mit einem JSON-Objekt, ohne Markdown, ohne Erklärung:
{
  "markt": "REWE",
  "datum": "2026-07-04",
  "bon_summe": 47.23,
  "positionen": [
    {"artikel": "Frische Vollmilch 1L", "menge": 6, "einzelpreis": 1.15,
     "zeilensumme": 6.90}
  ]
}
Falls Datum oder Bon-Summe nicht lesbar: null."""


# ----------------------------------------------------------------------------
# Paperless-Zugriff
# ----------------------------------------------------------------------------
def paperless_get(path, **kwargs):
    return requests.get(
        f"{PAPERLESS_URL}{path}",
        headers={"Authorization": f"Token {PAPERLESS_TOKEN}"},
        timeout=60,
        **kwargs,
    )


def fetch_document_images(doc_id):
    """Lädt das Dokument aus Paperless und liefert eine Liste von PIL-Images."""
    meta = paperless_get(f"/api/documents/{doc_id}/").json()
    resp = paperless_get(f"/api/documents/{doc_id}/download/")
    resp.raise_for_status()
    content_type = resp.headers.get("Content-Type", "")

    if "pdf" in content_type:
        pages = convert_from_bytes(resp.content, dpi=200)[:MAX_PAGES]
    else:
        pages = [Image.open(io.BytesIO(resp.content))]

    images = []
    for img in pages:
        if img.mode != "RGB":
            img = img.convert("RGB")
        if img.width > MAX_IMG_WIDTH:
            ratio = MAX_IMG_WIDTH / img.width
            img = img.resize((MAX_IMG_WIDTH, int(img.height * ratio)), Image.LANCZOS)
        # API-Limit: max. 8000 px pro Kante – sehr lange Bons zusätzlich stauchen
        if img.height > 7900:
            ratio = 7900 / img.height
            img = img.resize((int(img.width * ratio), 7900), Image.LANCZOS)
        images.append(img)
    return meta, images


def images_to_base64_png(images):
    out = []
    for img in images:
        buf = io.BytesIO()
        img.save(buf, format="PNG")
        out.append(base64.standard_b64encode(buf.getvalue()).decode("ascii"))
    return out


# ----------------------------------------------------------------------------
# LLM-Aufruf (Anthropic nativ oder OpenAI-kompatibel/Ollama)
# ----------------------------------------------------------------------------
def call_llm(b64_images, prompt):
    if PROVIDER == "anthropic":
        content = []
        for b64 in b64_images:
            content.append({
                "type": "image",
                "source": {"type": "base64", "media_type": "image/png", "data": b64},
            })
        content.append({"type": "text", "text": prompt})
        resp = requests.post(
            "https://api.anthropic.com/v1/messages",
            headers={
                "x-api-key": ANTHROPIC_API_KEY,
                "anthropic-version": "2023-06-01",
                "content-type": "application/json",
            },
            json={
                "model": ANTHROPIC_MODEL,
                "max_tokens": 4000,
                "messages": [{"role": "user", "content": content}],
            },
            timeout=300,
        )
        if resp.status_code >= 400:
            raise RuntimeError(
                f"Anthropic API {resp.status_code}: {resp.text[:500]}")
        return "".join(
            b.get("text", "") for b in resp.json()["content"] if b.get("type") == "text"
        )

    # OpenAI-kompatibel (Ollama & Co.) – Umschalten für Option C
    content = [
        {"type": "image_url", "image_url": {"url": f"data:image/png;base64,{b64}"}}
        for b64 in b64_images
    ]
    content.append({"type": "text", "text": prompt})
    resp = requests.post(
        f"{OPENAI_BASE_URL}/v1/chat/completions",
        json={
            "model": OPENAI_MODEL,
            "messages": [{"role": "user", "content": content}],
            "max_tokens": 4000,
            "temperature": 0,
        },
        timeout=3600,  # lokale CPU-Inferenz kann dauern
    )
    if resp.status_code >= 400:
        raise RuntimeError(
            f"LLM-API {resp.status_code}: {resp.text[:500]}")
    return resp.json()["choices"][0]["message"]["content"]


def parse_llm_json(text):
    """Extrahiert das JSON-Objekt aus der Modellantwort (robust gegen Fences)."""
    text = re.sub(r"```(json)?", "", text).strip()
    start, end = text.find("{"), text.rfind("}")
    if start == -1 or end == -1:
        raise ValueError(f"Keine JSON-Struktur in Antwort: {text[:200]}")
    data = json.loads(text[start:end + 1])
    if not isinstance(data.get("positionen"), list) or not data["positionen"]:
        raise ValueError("Antwort enthält keine Positionen")
    return data


# ----------------------------------------------------------------------------
# State (Dedupe + laufende Summe)
# ----------------------------------------------------------------------------
def load_state():
    if STATE_FILE.exists():
        return json.loads(STATE_FILE.read_text(encoding="utf-8"))
    return {"processed_ids": [], "open_receipts": []}


def save_state(state):
    STATE_FILE.write_text(json.dumps(state, indent=2), encoding="utf-8")


# ----------------------------------------------------------------------------
# Excel
# ----------------------------------------------------------------------------
HEADER_FONT = Font(bold=True)
CHECK_FILL = PatternFill("solid", fgColor="FFF3B0")   # gelb: bitte prüfen
OK_FILL = PatternFill("solid", fgColor="C6EFCE")      # grün
WARN_FILL = PatternFill("solid", fgColor="FFC7CE")    # rot

SUM_SHEET = "Summe"
EUR_FMT = '#,##0.00\u00a0"€"'


def create_workbook(path):
    wb = Workbook()
    ws = wb.active
    ws.title = SUM_SHEET
    ws["A1"] = "Offener Betrag:"
    ws["A1"].font = Font(bold=True, size=13)
    ws["D1"] = "=SUM(D4:D500)"
    ws["D1"].font = Font(bold=True, size=13)
    ws["D1"].number_format = EUR_FMT
    for col, head in enumerate(["Datum", "Markt", "Beleg-Blatt", "Betrag"], start=1):
        c = ws.cell(row=3, column=col, value=head)
        c.font = HEADER_FONT
    ws.column_dimensions["A"].width = 12
    ws.column_dimensions["B"].width = 22
    ws.column_dimensions["C"].width = 26
    ws.column_dimensions["D"].width = 14
    wb.save(path)
    log.info("Neue Abrechnungsdatei angelegt: %s", path)


def sanitize_sheet_name(name):
    name = re.sub(r"[\[\]:*?/\\]", " ", name).strip()
    return name[:31] or "Beleg"


def unique_sheet_name(wb, base):
    name, i = base, 2
    while name in wb.sheetnames:
        suffix = f" ({i})"
        name = base[: 31 - len(suffix)] + suffix
        i += 1
    return name


def add_receipt_sheet(data, doc_id, doc_title, full=False):
    """Fügt ein Beleg-Blatt hinzu, aktualisiert das Summenblatt.
    Gibt (betrag_se, plausibel, inkonsistente_zeilen, blattname) zurück."""
    path = OUTPUT_DIR / OPEN_FILENAME
    if not path.exists():
        create_workbook(path)
        # Neue Abrechnungsrunde -> offene Belege im State zurücksetzen
        state = load_state()
        state["open_receipts"] = []
        save_state(state)

    wb = load_workbook(path)
    datum = data.get("datum") or date.today().isoformat()
    markt = data.get("markt") or "Unbekannt"
    sheet = wb.create_sheet(unique_sheet_name(wb, sanitize_sheet_name(f"{datum} {markt}")))

    heads = ["Artikel", "Menge", "Einzelpreis", "Zeilensumme (Bon)",
             "Anteil abrechnen", "Betrag", "Hinweis"]
    for col, head in enumerate(heads, start=1):
        c = sheet.cell(row=1, column=col, value=head)
        c.font = HEADER_FONT
    sheet.column_dimensions["A"].width = 34
    for col in "BCDEF":
        sheet.column_dimensions[col].width = 15
    sheet.column_dimensions["G"].width = 26

    betrag_se = 0.0
    inkonsistente = 0
    row = 2
    for pos in data["positionen"]:
        menge = float(pos.get("menge") or 1)
        preis = float(pos.get("einzelpreis") or 0)
        zsum = float(pos.get("zeilensumme") or round(menge * preis, 2))
        anteil = float(pos.get("anteil") or 0)
        sheet.cell(row=row, column=1, value=pos.get("artikel", "?"))
        sheet.cell(row=row, column=2, value=menge)
        c = sheet.cell(row=row, column=3, value=preis); c.number_format = EUR_FMT
        c = sheet.cell(row=row, column=4, value=zsum); c.number_format = EUR_FMT
        ac = sheet.cell(row=row, column=5, value=anteil)
        ac.fill = CHECK_FILL  # editierbare Zelle hervorheben
        fc = sheet.cell(row=row, column=6, value=f"=E{row}*C{row}")
        fc.number_format = EUR_FMT
        # Zeilen-Konsistenz: Menge × Einzelpreis muss zur Bon-Zeilensumme passen.
        # Abweichung = Fehllesung von Menge oder Preis -> Betrag SE wäre falsch!
        hinweise = []
        if abs(menge * preis - zsum) > 0.015:
            hinweise.append("⚠ Menge×Preis ≠ Zeilensumme")
            inkonsistente += 1
        if pos.get("unsicher"):
            hinweise.append("⚠ Markierung unsicher")
        if hinweise:
            hc = sheet.cell(row=row, column=7, value=", ".join(hinweise))
            hc.fill = WARN_FILL
        betrag_se += round(anteil * preis, 2)
        row += 1

    last_data_row = row - 1
    total_row = row + 1
    sheet.cell(row=total_row, column=1, value="Summe Abrechnung").font = HEADER_FONT
    tc = sheet.cell(row=total_row, column=6, value=f"=SUM(F2:F{last_data_row})")
    tc.font = HEADER_FONT
    tc.number_format = EUR_FMT

    # ------------------------------------------------------------------
    # Kontrollen – als LEBENDE Excel-Formeln, damit sie sich bei manuellen
    # Korrekturen im Blatt automatisch mitaktualisieren.
    # ------------------------------------------------------------------
    bon_summe = data.get("bon_summe")
    zeilen_summe = sum(
        float(p.get("zeilensumme") or
              round(float(p.get("menge") or 1) * float(p.get("einzelpreis") or 0), 2))
        for p in data["positionen"])

    check_row = total_row + 2
    sheet.cell(row=check_row, column=1, value="Kontrolle: Positionen vs. Bon-Summe")
    c = sheet.cell(row=check_row, column=4, value=f"=SUM(D2:D{last_data_row})")
    c.number_format = EUR_FMT
    plausibel = None
    if bon_summe is not None:
        c2 = sheet.cell(row=check_row, column=5, value=float(bon_summe))
        c2.number_format = EUR_FMT
        status = sheet.cell(
            row=check_row, column=7,
            value=(f'=IF(ABS(D{check_row}-E{check_row})<0.02,'
                   f'"✓ stimmt überein","⚠ Differenz – Blatt prüfen!")'))
        plausibel = abs(zeilen_summe - float(bon_summe)) < 0.02
        status.fill = OK_FILL if (plausibel and inkonsistente == 0) else WARN_FILL

    # Bei Komplett-Belegen muss zusätzlich Betrag SE == Bon-Summe gelten
    if full and bon_summe is not None:
        check_row += 1
        sheet.cell(row=check_row, column=1,
                   value="Kontrolle (Komplett): Abrechnungsbetrag vs. Bon-Summe")
        c = sheet.cell(row=check_row, column=6, value=f"=F{total_row}")
        c.number_format = EUR_FMT
        sheet.cell(
            row=check_row, column=7,
            value=(f'=IF(ABS(F{total_row}-E{check_row - 1})<0.02,'
                   f'"✓ stimmt überein","⚠ Differenz – Blatt prüfen!")'))
        if abs(betrag_se - float(bon_summe)) >= 0.02:
            plausibel = False

    if plausibel is False or inkonsistente > 0:
        sheet.sheet_properties.tabColor = "FF6666"
    sheet.cell(row=check_row + 1, column=1,
               value=f"Quelle: Paperless-Dokument #{doc_id} – {doc_title}")

    # Summenblatt ergänzen
    ws = wb[SUM_SHEET]
    srow = 4
    while ws.cell(row=srow, column=1).value is not None:
        srow += 1
    ws.cell(row=srow, column=1, value=datum)
    ws.cell(row=srow, column=2, value=markt)
    ws.cell(row=srow, column=3, value=sheet.title)
    sc = ws.cell(row=srow, column=4, value=f"='{sheet.title}'!F{total_row}")
    sc.number_format = EUR_FMT

    wb.save(path)
    return round(betrag_se, 2), plausibel, inkonsistente, sheet.title


# ----------------------------------------------------------------------------
# Benachrichtigung (optional)
# ----------------------------------------------------------------------------
def notify(payload):
    if not HA_WEBHOOK_URL:
        return
    try:
        requests.post(HA_WEBHOOK_URL, json=payload, timeout=15)
    except Exception as e:  # noqa: BLE001
        log.warning("HA-Notification fehlgeschlagen: %s", e)


# ----------------------------------------------------------------------------
# Verarbeitung
# ----------------------------------------------------------------------------
def process_document(doc_id, full=False):
    with _excel_lock:
        state = load_state()
        if doc_id in state["processed_ids"]:
            log.info("Dokument %s bereits verarbeitet – übersprungen.", doc_id)
            return

    log.info("Verarbeite Paperless-Dokument %s (Modus: %s) ...",
             doc_id, "kompletter Beleg" if full else "markierte Positionen")
    try:
        meta, images = fetch_document_images(doc_id)
        raw = call_llm(images_to_base64_png(images), PROMPT_FULL if full else PROMPT)
        data = parse_llm_json(raw)

        if full:
            # Komplett-Tag gesetzt: alle Positionen vollständig übernehmen
            for pos in data["positionen"]:
                pos["anteil"] = pos.get("menge") or 1
                pos["unsicher"] = False

        with _excel_lock:
            betrag, plausibel, inkonsistente, sheet = add_receipt_sheet(
                data, doc_id, meta.get("title", ""), full=full)
            state = load_state()
            state["processed_ids"].append(doc_id)
            state["open_receipts"].append(
                {"doc_id": doc_id, "sheet": sheet, "betrag_se": betrag})
            offen = round(sum(r["betrag_se"] for r in state["open_receipts"]), 2)
            save_state(state)

        unsichere = sum(1 for p in data["positionen"] if p.get("unsicher"))
        log.info("Fertig: %s | Abrechnungsbetrag %.2f € | offen gesamt %.2f € | "
                 "plausibel=%s | inkonsistente Zeilen=%d",
                 sheet, betrag, offen, plausibel, inkonsistente)
        modus = " (kompletter Beleg)" if full else ""
        msg = (f"Beleg {data.get('markt', '?')}{modus} erfasst: {betrag:.2f} € zur "
               f"Abrechnung. Offen gesamt: {offen:.2f} €.")
        if plausibel is False:
            msg += " ⚠ Summen-Differenz, bitte Blatt prüfen!"
        if inkonsistente:
            msg += f" ⚠ {inkonsistente} Zeile(n) rechnerisch inkonsistent."
        if unsichere:
            msg += f" ⚠ {unsichere} Markierung(en) unsicher."
        notify({"event": "receipt_processed", "markt": data.get("markt"),
                "betrag_se": betrag, "offen_gesamt": offen,
                "plausibel": plausibel, "unsichere_positionen": unsichere,
                "inkonsistente_zeilen": inkonsistente,
                "komplett": full,
                "sheet": sheet, "doc_id": doc_id, "message": msg})

    except Exception as e:  # noqa: BLE001
        log.exception("Fehler bei Dokument %s", doc_id)
        notify({"event": "receipt_error", "doc_id": doc_id,
                "message": f"Beleg-Verarbeitung fehlgeschlagen (Dokument {doc_id}): {e}"})


# ----------------------------------------------------------------------------
# HTTP-Endpunkte
# ----------------------------------------------------------------------------
@app.route("/process", methods=["POST"])
def process_endpoint():
    body = request.get_json(silent=True) or {}
    doc_id = body.get("document_id")
    full = bool(body.get("full", False))
    if not isinstance(doc_id, int):
        return jsonify({"error": "document_id (int) fehlt"}), 400
    threading.Thread(target=process_document, args=(doc_id, full), daemon=True).start()
    return jsonify({"status": "accepted", "document_id": doc_id, "full": full}), 202


@app.route("/health", methods=["GET"])
def health():
    return jsonify({"status": "ok", "provider": PROVIDER})


if __name__ == "__main__":
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    app.run(host="0.0.0.0", port=5000)
