#!/usr/bin/env bash
# Paperless-ngx Post-Consume-Hook
# Prüft günstig (ohne API-Call) den Trigger-Tag und stößt den Receipt-Worker
# an. Läuft IM Paperless-Container – braucht nur curl (ist im Image enthalten).
#
# Personen-Zuordnung, Verarbeitet-Status und Duplikatschutz entscheidet der
# Worker selbst anhand der aktuellen Paperless-Tags des Dokuments – hier wird
# nur grob vorgefiltert, damit unbeteiligte Dokumente den Worker gar nicht
# erst wecken.

TRIGGER_TAG="${RECEIPT_TRIGGER_TAG:-Auslagen}"
WORKER_URL="${RECEIPT_WORKER_URL:-http://receipt-worker:5000/process}"

# DOCUMENT_TAGS ist eine kommaseparierte Liste der Tag-Namen
if ! echo ",${DOCUMENT_TAGS}," | grep -qi ",${TRIGGER_TAG},"; then
    exit 0
fi

# Kurzer Timeout, Worker arbeitet asynchron weiter – Konsumierung nicht blockieren
curl -m 10 -s -X POST "${WORKER_URL}" \
    -H "Content-Type: application/json" \
    -d "{\"document_id\": ${DOCUMENT_ID}}" \
    >> /tmp/receipt-post-consume.log 2>&1 || true

exit 0
