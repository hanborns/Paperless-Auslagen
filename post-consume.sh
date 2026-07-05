#!/usr/bin/env bash
# Paperless-ngx Post-Consume-Hook
# Prüft die Tags des Dokuments und stößt den Receipt-Worker an.
# Läuft IM Paperless-Container – braucht nur curl (ist im Image enthalten).
#
# Zwei Modi:
#   RECEIPT_FULL_TAG    -> kompletter Beleg geht in die Abrechnung
#   RECEIPT_TRIGGER_TAG -> nur handschriftlich markierte Positionen
#                          ("R" = ganze Menge, Zahl = Anteil)
# Sind beide Tags gesetzt, gewinnt der Komplett-Tag.

TRIGGER_TAG="${RECEIPT_TRIGGER_TAG:-abrechnung}"
FULL_TAG="${RECEIPT_FULL_TAG:-abrechnung-komplett}"
WORKER_URL="${RECEIPT_WORKER_URL:-http://receipt-worker:5000/process}"

# DOCUMENT_TAGS ist eine kommaseparierte Liste der Tag-Namen
if echo ",${DOCUMENT_TAGS}," | grep -qi ",${FULL_TAG},"; then
    PAYLOAD="{\"document_id\": ${DOCUMENT_ID}, \"full\": true}"
elif echo ",${DOCUMENT_TAGS}," | grep -qi ",${TRIGGER_TAG},"; then
    PAYLOAD="{\"document_id\": ${DOCUMENT_ID}, \"full\": false}"
else
    exit 0
fi

# Kurzer Timeout, Worker arbeitet asynchron weiter – Konsumierung nicht blockieren
curl -m 10 -s -X POST "${WORKER_URL}" \
    -H "Content-Type: application/json" \
    -d "${PAYLOAD}" \
    >> /tmp/receipt-post-consume.log 2>&1 || true

exit 0
