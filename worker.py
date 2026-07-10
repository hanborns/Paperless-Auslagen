#!/usr/bin/env python3
"""
Receipt-Worker für Paperless-ngx
--------------------------------
Nimmt per HTTP eine Paperless-Dokument-ID entgegen (Post-Consume-Hook) oder
holt sich unverarbeitete Dokumente per periodischem Scan selbst, lädt den
Beleg, extrahiert alle Kaufpositionen per Azure Document Intelligence
(Default) oder per Vision-LLM (Anthropic, OpenAI-kompatibel/Ollama), und
pflegt pro Person eine Excel-Datei mit einem Blatt pro Beleg und einem
Summenblatt.

Tags in Paperless steuern den Ablauf (Namen per Environment konfigurierbar):
  - Trigger-Tag (Default "Auslagen")     -> Dokument wird verarbeitet
  - Personen-Tags (Allowlist, z. B. "Rose,Anna") -> je Treffer ein eigenes
    Excel-Blatt/eine eigene Datei; kein Treffer -> gemeinsame Standard-Datei
  - Verarbeitet-Tag (Default "Verarbeitet") -> wird nach erfolgreicher
    Verarbeitung gesetzt. Dient als Duplikatschutz UND als Reprocessing-
    Schalter: Tag in Paperless entfernen lässt das Dokument beim nächsten
    Scan (oder erneuten Consume) noch einmal verarbeiten.

Jede Position wird vollständig übernommen (Anteil = Menge); falls nur ein
Teil geteilt wurde, wird das direkt in der editierbaren Spalte "Anteil
abrechnen" im Excel korrigiert.
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
from pdf2image import convert_from_bytes
from PIL import Image

# ----------------------------------------------------------------------------
# Konfiguration (per Environment)
# ----------------------------------------------------------------------------
PAPERLESS_URL = os.environ["PAPERLESS_URL"].rstrip("/")   # z. B. http://paperless-webserver:8000
PAPERLESS_TOKEN = os.environ["PAPERLESS_TOKEN"]

PROVIDER = os.environ.get("PROVIDER", "azure_di")         # "azure_di" | "anthropic" | "openai"
ANTHROPIC_API_KEY = os.environ.get("ANTHROPIC_API_KEY", "")
ANTHROPIC_MODEL = os.environ.get("ANTHROPIC_MODEL", "claude-sonnet-4-6")
OPENAI_BASE_URL = os.environ.get("OPENAI_BASE_URL", "http://ollama:11434").rstrip("/")
OPENAI_MODEL = os.environ.get("OPENAI_MODEL", "qwen2.5vl:7b")
AZURE_DI_ENDPOINT = os.environ.get("AZURE_DI_ENDPOINT", "").rstrip("/")
AZURE_DI_KEY = os.environ.get("AZURE_DI_KEY", "")
AZURE_DI_MODEL = os.environ.get("AZURE_DI_MODEL", "prebuilt-receipt")
AZURE_DI_API_VERSION = os.environ.get("AZURE_DI_API_VERSION", "2024-11-30")
AZURE_DI_CONFIDENCE_THRESHOLD = float(os.environ.get("AZURE_DI_CONFIDENCE_THRESHOLD", "0.7"))

OUTPUT_DIR = Path(os.environ.get("OUTPUT_DIR", "/output"))
OPEN_FILENAME = os.environ.get("OPEN_FILENAME", "Auslagen_offen.xlsx")
LOCK_FILE = OUTPUT_DIR / ".receipt_worker.lock"

RECEIPT_TRIGGER_TAG = os.environ.get("RECEIPT_TRIGGER_TAG", "Auslagen")
RECEIPT_PERSON_TAGS = [t.strip() for t in os.environ.get("RECEIPT_PERSON_TAGS", "").split(",") if t.strip()]
RECEIPT_PROCESSED_TAG = os.environ.get("RECEIPT_PROCESSED_TAG", "Verarbeitet")
SCAN_INTERVAL_MINUTES = int(os.environ.get("SCAN_INTERVAL_MINUTES", "30"))

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

# ----------------------------------------------------------------------------
# Extraktions-Prompt (immer vollständig – keine Handschrift-Erkennung mehr,
# Korrekturen passieren direkt im Excel)
# ----------------------------------------------------------------------------
PROMPT = """Du siehst einen eingescannten deutschen Supermarkt-Kassenbon. \
Extrahiere ALLE Kaufpositionen vollständig.

Wichtige Regeln:
- Mengenzeilen wie "6 x 1,15" gehören zur darüber/darunter stehenden Artikelzeile: \
Menge 6, Einzelpreis 1,15.
- Positionen ohne explizite Menge haben Menge 1 und Einzelpreis = Zeilenpreis.
- Rabatt-/Nachlasszeilen als eigene Position mit negativem Preis aufnehmen.
- Pfandzeilen als normale Positionen behandeln.
- Preise als Dezimalzahl mit Punkt (1.15), niemals mit Komma.
- Lies auch den Markt-/Filialnamen, das Kaufdatum und die gedruckte GESAMTSUMME des Bons.
- Wenn du dir bei einer Position (Artikel, Menge oder Preis) nicht sicher bist, \
setze "unsicher": true bei dieser Position.

Antworte AUSSCHLIESSLICH mit einem JSON-Objekt, ohne Markdown, ohne Erklärung:
{
  "markt": "REWE",
  "datum": "2026-07-04",
  "bon_summe": 47.23,
  "positionen": [
    {"artikel": "Frische Vollmilch 1L", "menge": 6, "einzelpreis": 1.15,
     "zeilensumme": 6.90, "unsicher": false}
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


def paperless_post(path, json_body, timeout=30):
    resp = requests.post(
        f"{PAPERLESS_URL}{path}",
        headers={"Authorization": f"Token {PAPERLESS_TOKEN}"},
        json=json_body,
        timeout=timeout,
    )
    if resp.status_code >= 400:
        raise RuntimeError(f"Paperless API {resp.status_code} bei {path}: {resp.text[:500]}")
    return resp


_tag_id_cache = {}


def resolve_tag_id(name):
    """Löst einen Tag-Namen in seine Paperless-Tag-ID auf, legt ihn bei Bedarf an."""
    if name in _tag_id_cache:
        return _tag_id_cache[name]
    resp = paperless_get("/api/tags/", params={"name__iexact": name})
    resp.raise_for_status()
    results = resp.json().get("results", [])
    if results:
        tag_id = results[0]["id"]
    else:
        tag_id = paperless_post("/api/tags/", {"name": name}).json()["id"]
        log.info("Paperless-Tag angelegt: %s (id=%s)", name, tag_id)
    _tag_id_cache[name] = tag_id
    return tag_id


def resolve_document_tags(doc_meta):
    """Liefert (hat_trigger_tag, passende_personen) anhand der aktuellen Tags des Dokuments."""
    tag_ids = set(doc_meta.get("tags", []))
    has_trigger = resolve_tag_id(RECEIPT_TRIGGER_TAG) in tag_ids
    persons = [p for p in RECEIPT_PERSON_TAGS if resolve_tag_id(p) in tag_ids]
    return has_trigger, persons


def is_processed(doc_meta):
    return resolve_tag_id(RECEIPT_PROCESSED_TAG) in set(doc_meta.get("tags", []))


def mark_processed(doc_id):
    paperless_post("/api/documents/bulk_edit/", {
        "documents": [doc_id],
        "method": "modify_tags",
        "parameters": {"add_tags": [resolve_tag_id(RECEIPT_PROCESSED_TAG)], "remove_tags": []},
    })


def download_document(doc_id):
    resp = paperless_get(f"/api/documents/{doc_id}/download/")
    resp.raise_for_status()
    return resp.content, resp.headers.get("Content-Type", "")


def rasterize(content, content_type):
    """PDF/Bild -> Liste von PIL-Images für Vision-LLM-Provider (Anthropic/OpenAI)."""
    if "pdf" in content_type:
        pages = convert_from_bytes(content, dpi=200)[:MAX_PAGES]
    else:
        pages = [Image.open(io.BytesIO(content))]

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
    return images


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

    # OpenAI-kompatibel (Ollama & Co.)
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
# Azure Document Intelligence (prebuilt-receipt)
# ----------------------------------------------------------------------------
def call_azure_di(doc_bytes):
    resp = requests.post(
        f"{AZURE_DI_ENDPOINT}/documentintelligence/documentModels/"
        f"{AZURE_DI_MODEL}:analyze?api-version={AZURE_DI_API_VERSION}",
        headers={
            "Ocp-Apim-Subscription-Key": AZURE_DI_KEY,
            "Content-Type": "application/json",
        },
        json={"base64Source": base64.standard_b64encode(doc_bytes).decode("ascii")},
        timeout=60,
    )
    if resp.status_code >= 400:
        raise RuntimeError(f"Azure DI API {resp.status_code}: {resp.text[:500]}")
    operation_location = resp.headers["Operation-Location"]

    deadline = time.time() + 120
    while time.time() < deadline:
        poll = requests.get(
            operation_location,
            headers={"Ocp-Apim-Subscription-Key": AZURE_DI_KEY},
            timeout=30,
        )
        if poll.status_code >= 400:
            raise RuntimeError(f"Azure DI API {poll.status_code}: {poll.text[:500]}")
        body = poll.json()
        status = body.get("status")
        if status == "succeeded":
            return body["analyzeResult"]
        if status == "failed":
            raise RuntimeError(f"Azure DI Analyse fehlgeschlagen: {json.dumps(body)[:500]}")
        time.sleep(2)
    raise RuntimeError("Azure DI Analyse: Timeout beim Warten auf Ergebnis")


def normalize_azure_result(analyze_result):
    """Bringt das Azure-DI-Ergebnis in dieselbe Form wie parse_llm_json()."""
    documents = analyze_result.get("documents") or [{}]
    fields = documents[0].get("fields", {})

    def currency(field):
        value = (field or {}).get("valueCurrency")
        return value.get("amount") if value else None

    positionen = []
    for item in (fields.get("Items") or {}).get("valueArray", []):
        obj = item.get("valueObject", {})
        menge = (obj.get("Quantity") or {}).get("valueNumber") or 1
        preis = currency(obj.get("Price"))
        zsum = currency(obj.get("TotalPrice"))
        if preis is None:
            preis = round(zsum / menge, 2) if zsum is not None and menge else 0.0
        if zsum is None:
            zsum = round(menge * preis, 2)
        confidences = [
            (obj.get(k) or {}).get("confidence", 1.0)
            for k in ("Description", "Quantity", "Price", "TotalPrice")
            if obj.get(k)
        ]
        positionen.append({
            "artikel": (obj.get("Description") or {}).get("valueString", "?"),
            "menge": menge,
            "einzelpreis": preis,
            "zeilensumme": zsum,
            "unsicher": any(c < AZURE_DI_CONFIDENCE_THRESHOLD for c in confidences),
        })

    return {
        "markt": (fields.get("MerchantName") or {}).get("valueString"),
        "datum": (fields.get("TransactionDate") or {}).get("valueDate"),
        "bon_summe": currency(fields.get("Total")),
        "positionen": positionen,
    }


# ----------------------------------------------------------------------------
# Nebenläufigkeits-Sperre (eine Verarbeitung zur Zeit, prozessübergreifend)
# ----------------------------------------------------------------------------
def acquire_lock():
    try:
        fd = os.open(LOCK_FILE, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
    except FileExistsError:
        return False
    os.write(fd, f"{os.getpid()} {datetime.now().isoformat()}".encode("utf-8"))
    os.close(fd)
    return True


def release_lock():
    try:
        LOCK_FILE.unlink()
    except FileNotFoundError:
        pass


# ----------------------------------------------------------------------------
# Excel
# ----------------------------------------------------------------------------
HEADER_FONT = Font(bold=True)
CHECK_FILL = PatternFill("solid", fgColor="FFF3B0")   # gelb: bitte prüfen
OK_FILL = PatternFill("solid", fgColor="C6EFCE")      # grün
WARN_FILL = PatternFill("solid", fgColor="FFC7CE")    # rot

SUM_SHEET = "Summe"
EUR_FMT = '#,##0.00 "€"'


def excel_path_for_person(person):
    """Leere Person ('') -> Standarddatei; sonst Personen-Suffix vor der Endung."""
    if not person:
        return OUTPUT_DIR / OPEN_FILENAME
    stem, suffix = Path(OPEN_FILENAME).stem, Path(OPEN_FILENAME).suffix
    return OUTPUT_DIR / f"{stem}_{person}{suffix}"


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


def add_receipt_sheet(data, doc_id, doc_title, person):
    """Fügt ein Beleg-Blatt in der Excel-Datei der angegebenen Person hinzu
    (leerer String = gemeinsame Standarddatei ohne Personen-Tag-Treffer).
    Gibt (betrag_auslagen, plausibel, inkonsistente_zeilen, blattname) zurück."""
    path = excel_path_for_person(person)
    if not path.exists():
        create_workbook(path)

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

    betrag_auslagen = 0.0
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
        # Abweichung = Fehllesung von Menge oder Preis -> Auslagenbetrag wäre falsch!
        hinweise = []
        if abs(menge * preis - zsum) > 0.015:
            hinweise.append("⚠ Menge×Preis ≠ Zeilensumme")
            inkonsistente += 1
        if pos.get("unsicher"):
            hinweise.append("⚠ Position unsicher gelesen")
        if hinweise:
            hc = sheet.cell(row=row, column=7, value=", ".join(hinweise))
            hc.fill = WARN_FILL
        betrag_auslagen += round(anteil * preis, 2)
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

        # Zweite Kontrolle: Abrechnungsbetrag muss (bei voller Übernahme
        # jeder Position) ebenfalls der Bon-Summe entsprechen.
        check_row += 1
        sheet.cell(row=check_row, column=1,
                   value="Kontrolle: Abrechnungsbetrag vs. Bon-Summe")
        c = sheet.cell(row=check_row, column=6, value=f"=F{total_row}")
        c.number_format = EUR_FMT
        sheet.cell(
            row=check_row, column=7,
            value=(f'=IF(ABS(F{total_row}-E{check_row - 1})<0.02,'
                   f'"✓ stimmt überein","⚠ Differenz – Blatt prüfen!")'))
        if abs(betrag_auslagen - float(bon_summe)) >= 0.02:
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
    return round(betrag_auslagen, 2), plausibel, inkonsistente, sheet.title


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
def process_document(doc_id):
    if not acquire_lock():
        log.info("Verarbeitung läuft bereits – Dokument %s übersprungen "
                 "(nächster Lauf holt es nach).", doc_id)
        return

    try:
        meta = paperless_get(f"/api/documents/{doc_id}/").json()

        if is_processed(meta):
            log.info("Dokument %s bereits verarbeitet (Tag '%s' gesetzt) – übersprungen.",
                     doc_id, RECEIPT_PROCESSED_TAG)
            return

        has_trigger, persons = resolve_document_tags(meta)
        if not has_trigger:
            log.info("Dokument %s ohne Trigger-Tag '%s' – ignoriert.",
                     doc_id, RECEIPT_TRIGGER_TAG)
            return

        log.info("Verarbeite Paperless-Dokument %s (Personen: %s) ...",
                 doc_id, ", ".join(persons) if persons else "gemeinsam")

        content, content_type = download_document(doc_id)
        if PROVIDER == "azure_di":
            data = normalize_azure_result(call_azure_di(content))
        else:
            images = rasterize(content, content_type)
            raw = call_llm(images_to_base64_png(images), PROMPT)
            data = parse_llm_json(raw)

        for pos in data["positionen"]:
            pos["anteil"] = pos.get("menge") or 1

        unsichere = sum(1 for p in data["positionen"] if p.get("unsicher"))
        targets = persons or [""]
        for person in targets:
            betrag, plausibel, inkonsistente, sheet = add_receipt_sheet(
                data, doc_id, meta.get("title", ""), person)

            label = person or "gemeinsam"
            log.info("Fertig: %s (%s) | Auslagenbetrag %.2f € | plausibel=%s | "
                     "inkonsistente Zeilen=%d", sheet, label, betrag, plausibel, inkonsistente)

            msg = f"Beleg {data.get('markt', '?')} erfasst ({label}): {betrag:.2f} € zur Abrechnung."
            if plausibel is False:
                msg += " ⚠ Summen-Differenz, bitte Blatt prüfen!"
            if inkonsistente:
                msg += f" ⚠ {inkonsistente} Zeile(n) rechnerisch inkonsistent."
            if unsichere:
                msg += f" ⚠ {unsichere} Position(en) unsicher gelesen."
            notify({"event": "receipt_processed", "person": person, "markt": data.get("markt"),
                    "betrag_auslagen": betrag, "plausibel": plausibel,
                    "unsichere_positionen": unsichere, "inkonsistente_zeilen": inkonsistente,
                    "sheet": sheet, "doc_id": doc_id, "message": msg})

        mark_processed(doc_id)

    except Exception as e:  # noqa: BLE001
        log.exception("Fehler bei Dokument %s", doc_id)
        notify({"event": "receipt_error", "doc_id": doc_id,
                "message": f"Beleg-Verarbeitung fehlgeschlagen (Dokument {doc_id}): {e}"})
    finally:
        release_lock()


# ----------------------------------------------------------------------------
# Backlog-Scan: findet Dokumente mit Trigger-Tag, die (noch) nicht den
# Verarbeitet-Tag tragen – fängt z. B. per E-Mail importierte und erst
# nachträglich getaggte Belege ab, die keinen Post-Consume-Hook ausgelöst haben.
# ----------------------------------------------------------------------------
def scan_paperless():
    trigger_id = resolve_tag_id(RECEIPT_TRIGGER_TAG)
    processed_id = resolve_tag_id(RECEIPT_PROCESSED_TAG)

    found = 0
    resp = paperless_get("/api/documents/", params={
        "tags__id__all": trigger_id, "page_size": 100,
    })
    resp.raise_for_status()
    body = resp.json()
    while True:
        for doc in body.get("results", []):
            if processed_id not in doc.get("tags", []):
                found += 1
                process_document(doc["id"])
        next_url = body.get("next")
        if not next_url:
            break
        resp = requests.get(next_url,
                             headers={"Authorization": f"Token {PAPERLESS_TOKEN}"}, timeout=60)
        resp.raise_for_status()
        body = resp.json()

    log.info("Scan abgeschlossen: %d unverarbeitete(s) Dokument(e) gefunden.", found)


def _scan_loop():
    while True:
        time.sleep(SCAN_INTERVAL_MINUTES * 60)
        try:
            scan_paperless()
        except Exception:  # noqa: BLE001
            log.exception("Periodischer Scan fehlgeschlagen")


# ----------------------------------------------------------------------------
# HTTP-Endpunkte
# ----------------------------------------------------------------------------
@app.route("/process", methods=["POST"])
def process_endpoint():
    body = request.get_json(silent=True) or {}
    doc_id = body.get("document_id")
    if not isinstance(doc_id, int):
        return jsonify({"error": "document_id (int) fehlt"}), 400
    threading.Thread(target=process_document, args=(doc_id,), daemon=True).start()
    return jsonify({"status": "accepted", "document_id": doc_id}), 202


@app.route("/scan", methods=["POST"])
def scan_endpoint():
    threading.Thread(target=scan_paperless, daemon=True).start()
    return jsonify({"status": "accepted"}), 202


@app.route("/health", methods=["GET"])
def health():
    return jsonify({"status": "ok", "provider": PROVIDER})


if __name__ == "__main__":
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    LOCK_FILE.unlink(missing_ok=True)  # Neustart: keine Verarbeitung kann noch laufen
    threading.Thread(target=_scan_loop, daemon=True).start()
    app.run(host="0.0.0.0", port=5000)
