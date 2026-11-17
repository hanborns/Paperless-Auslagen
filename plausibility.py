"""
Plausibilitätskontrolle und automatische Korrektur für extrahierte Belege
-----------------------------------------------------------------------
Reine Funktionen ohne Netzwerk-/Dateizugriff und ohne Abhängigkeiten außer
der Standardbibliothek – dadurch direkt testbar (tests/) und gegen
gespeicherte Azure-Rohdaten wiederholbar (tools/replay.py).

Ablauf:
  1. normalize_azure_result(): Azure-DI-Ergebnis -> interne Form
     (dieselbe Form, die parse_llm_json() für die LLM-Provider liefert).
  2. reconcile(): prüft, ob Positionen und Bon-Summe zusammenpassen, und
     probiert sonst nacheinander Korrekturregeln. Eine Korrektur wird NUR
     übernommen, wenn danach alles exakt aufgeht (Summe der Positionen =
     Bon-Summe und jede Zeile Menge × Einzelpreis = Zeilensumme). Sonst
     bleibt das Original stehen (Status "unklar") – der Worker kann dann eine
     Zweitmeinung einholen bzw. das Blatt wie bisher als prüfbedürftig
     markieren.

Jede Korrektur wird an der Position unter "korrektur" vermerkt und landet
sichtbar in der Hinweis-Spalte des Excel-Blatts – nichts wird still geändert.

Neue Regeln: als Funktion rule_xxx(data, ctx) -> int (Anzahl Änderungen)
schreiben, in RULES eintragen und mit einem Testfall absichern.
"""

import copy
import re

ROW_TOL = 0.015   # Toleranz Menge × Einzelpreis vs. Zeilensumme (wie im Excel-Check)
SUM_TOL = 0.02    # Toleranz Summe der Positionen vs. Bon-Summe (wie im Excel-Check)
MAX_COMPONENTS = 8  # max. Anzahl Komponenten-Zeilen unter einer Menü-Zeile
DEFAULT_VAT_RATES = (0.07, 0.19)


# ----------------------------------------------------------------------------
# Azure DI -> interne Form
# ----------------------------------------------------------------------------
def _currency(field):
    value = (field or {}).get("valueCurrency")
    return value.get("amount") if value else None


def _number(field):
    field = field or {}
    if "valueNumber" in field:
        return field["valueNumber"]
    return _currency(field)


def normalize_azure_result(analyze_result, confidence_threshold=0.7):
    """Bringt das Azure-DI-Ergebnis in dieselbe Form wie parse_llm_json().

    Zusätzlich zu den Standardfeldern werden Zwischensumme (netto), Steuer
    und Steuersätze übernommen – die braucht die Netto/Brutto-Korrektur."""
    documents = analyze_result.get("documents") or [{}]
    fields = documents[0].get("fields", {})

    positionen = []
    for item in (fields.get("Items") or {}).get("valueArray", []):
        obj = item.get("valueObject", {})
        menge = (obj.get("Quantity") or {}).get("valueNumber") or 1
        preis = _currency(obj.get("Price"))
        zsum = _currency(obj.get("TotalPrice"))
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
            "unsicher": any(c < confidence_threshold for c in confidences),
        })

    steuersaetze = []
    for td in (fields.get("TaxDetails") or {}).get("valueArray", []):
        tobj = td.get("valueObject", {})
        rate = _number(tobj.get("Rate"))
        if rate is None:
            continue
        if rate > 1:          # 19 statt 0.19
            rate = rate / 100
        steuersaetze.append({
            "satz": round(rate, 4),
            "netto": _currency(tobj.get("NetAmount")),
            "steuer": _currency(tobj.get("Amount")),
        })

    return {
        "markt": (fields.get("MerchantName") or {}).get("valueString"),
        "datum": (fields.get("TransactionDate") or {}).get("valueDate"),
        "bon_summe": _currency(fields.get("Total")),
        "zwischensumme": _currency(fields.get("Subtotal")),
        "steuer": _currency(fields.get("TotalTax")),
        "steuersaetze": steuersaetze,
        "positionen": positionen,
    }


# ----------------------------------------------------------------------------
# Bewertung
# ----------------------------------------------------------------------------
def _f(value, default=0.0):
    try:
        return float(value)
    except (TypeError, ValueError):
        return default


def zeilensumme(pos):
    if pos.get("zeilensumme") is not None:
        return _f(pos["zeilensumme"])
    return round(_f(pos.get("menge"), 1) * _f(pos.get("einzelpreis")), 2)


def row_consistent(pos):
    menge = _f(pos.get("menge"), 1) or 1
    return abs(menge * _f(pos.get("einzelpreis")) - zeilensumme(pos)) <= ROW_TOL


def assess(data):
    """Bewertet die Plausibilität. ok = Bon-Summe bekannt, Summe passt und
    alle Zeilen rechnerisch konsistent."""
    positionen = data.get("positionen") or []
    summe = round(sum(zeilensumme(p) for p in positionen), 2)
    bon = data.get("bon_summe")
    diff = None if bon is None else round(summe - _f(bon), 2)
    bad_rows = sum(1 for p in positionen if not row_consistent(p))
    sum_ok = diff is not None and abs(diff) < SUM_TOL
    return {
        "summe_positionen": summe,
        "differenz": diff,
        "inkonsistente_zeilen": bad_rows,
        "summe_ok": sum_ok,
        "ok": sum_ok and bad_rows == 0,
    }


def _score(data):
    """Kleiner = besser. Summen-Treffer zählt vor Zeilenkonsistenz."""
    a = assess(data)
    return (0 if a["summe_ok"] else 1,
            a["inkonsistente_zeilen"],
            abs(a["differenz"]) if a["differenz"] is not None else 0.0)


def _note(pos, text):
    pos.setdefault("korrektur", []).append(text)


def _eur(value):
    return f"{_f(value):.2f}".replace(".", ",")


# ----------------------------------------------------------------------------
# Regeln
# ----------------------------------------------------------------------------
_QTY_LINE = re.compile(
    r"^\s*(\d+(?:[.,]\d+)?)\s*(?:kg|stk|st)?\.?\s*[x×*]\s*(\d+[.,]\d{2})\s*"
    r"(?:€|eur)?\s*(?:/\s*kg)?\s*$",
    re.IGNORECASE)


def _parse_qty_line(text):
    m = _QTY_LINE.match(text or "")
    if not m:
        return None
    return float(m.group(1).replace(",", ".")), float(m.group(2).replace(",", "."))


_ADDON = re.compile(r"^\s*(extra|plus|\+)\b", re.IGNORECASE)
_AT_PRICE = re.compile(r"\s*@\D{0,2}?(\d+[.,]\d{2})\s*")


def _is_addon(text):
    return bool(_ADDON.match(text or ""))


def rule_at_price_quantity(data, ctx):
    """Burger-King-Muster: "Whopper@7,29 … EUR 14,58" bzw. "Extra Bacon@1,20 …
    EUR 2,40" – der Stückpreis steht hinter dem @ im Artikeltext, die Menge
    wird bei Extras NICHT gedruckt (sie erbt die Menge der Hauptzeile).
    Menge = Zeilensumme / Stückpreis, wenn das ganzzahlig aufgeht. Ändert
    keine Zeilensumme – wird daher immer angewendet."""
    changes = 0
    for pos in data["positionen"]:
        artikel = pos.get("artikel") or ""
        m = _AT_PRICE.search(artikel)
        if not m:
            continue
        stueck = float(m.group(1).replace(",", "."))
        zsum = zeilensumme(pos)
        if stueck <= 0 or zsum <= 0:
            continue
        menge = zsum / stueck
        if abs(menge - round(menge)) > 0.01 or round(menge) < 1:
            continue
        menge = int(round(menge))
        pos["artikel"] = _AT_PRICE.sub(" ", artikel, count=1).strip()
        if menge != _f(pos.get("menge"), 1) or abs(stueck - _f(pos.get("einzelpreis"))) > 0.005:
            _note(pos, f"Menge {menge} × {_eur(stueck)} aus „@{_eur(stueck)}“ abgeleitet "
                       f"(vorher Menge {_f(pos.get('menge'), 1):g}, Preis {_eur(pos.get('einzelpreis'))})")
            changes += 1
        pos["menge"] = menge
        pos["einzelpreis"] = stueck
        pos["zeilensumme"] = zsum
    return changes


def rule_fold_zero_components(data, ctx):
    """Burger-King-Muster: Komponenten eines Menüs ohne eigenen Preis (0,00)
    werden in den Artikelnamen der darüberstehenden Menü-Zeile übernommen.
    Ändert keine Beträge – wird daher immer angewendet."""
    out, changes = [], 0
    positionen = data["positionen"]
    for idx, pos in enumerate(positionen):
        is_zero = abs(zeilensumme(pos)) < 0.005 and abs(_f(pos.get("einzelpreis"))) < 0.005
        # Artikel direkt vor/nach einer Mengenzeile nicht anfassen – der Betrag
        # ist dort evtl. nur verrutscht (rule_merge_quantity_lines)
        near_qty = any(
            0 <= j < len(positionen) and _parse_qty_line(positionen[j].get("artikel"))
            for j in (idx - 1, idx + 1))
        if is_zero and out and not near_qty and _parse_qty_line(pos.get("artikel")) is None:
            # Elternzeile = letzte Position, die kein Aufpreis-Extra ist
            # (BK: "Extra Käse EUR 1,20" steht mitten in den Menü-Komponenten)
            parent = next((p for p in reversed(out) if not _is_addon(p.get("artikel"))), out[-1])
            parent.setdefault("_komponenten", []).append(pos.get("artikel", "?"))
            if pos.get("unsicher"):
                parent["unsicher"] = True
            changes += 1
            continue
        out.append(pos)
    for pos in out:
        comps = pos.pop("_komponenten", None)
        if comps:
            pos.setdefault("artikel_basis", pos.get("artikel", "?"))
            pos["artikel"] = f"{pos.get('artikel', '?')}: {', '.join(comps)}"
            _note(pos, f"{len(comps)} Menü-Komponente(n) ohne Preis zusammengefasst")
    data["positionen"] = out
    return changes


def _matches_any(text, terms):
    text = (text or "").lower()
    return any(t.strip().lower() in text for t in terms if t and t.strip())


def rule_fold_priced_addons(data, ctx):
    """Burger-King-Muster: Aufpreis-Extras ("Extra Käse EUR 1,20",
    "Extra Bacon@1,20 EUR 2,40") stehen mit eigenem Betrag UNTER der
    Hauptposition, sind in deren Zeilensumme aber NICHT enthalten. Fachlich
    gehören sie zur Hauptposition: ihr Betrag wird zur Zeilensumme der
    letzten Nicht-Extra-Position davor addiert, der Einzelpreis daraus neu
    berechnet (Menge bleibt). Die Gesamtsumme ändert sich nicht – die Regel
    läuft daher immer, und zwar NACH rule_fold_zero_components (sonst wäre
    im Menü die preislose Komponente direkt über dem Extra die Elternzeile).

    Extras, die auf ctx["keep_separate"] passen (RECEIPT_EXCLUDE_ITEMS),
    bleiben eigene Zeilen, damit sie weiter Anteil 0 bekommen können."""
    keep = ctx.get("keep_separate") or ()
    out, changes = [], 0
    for pos in data["positionen"]:
        artikel = pos.get("artikel") or ""
        zsum = zeilensumme(pos)
        parent = next((p for p in reversed(out) if not _is_addon(p.get("artikel"))), None)
        if (_is_addon(artikel) and zsum > 0 and parent is not None
                and zeilensumme(parent) > 0 and not _matches_any(artikel, keep)
                and _parse_qty_line(artikel) is None):
            parent.setdefault("_extras", []).append(pos)
            if pos.get("unsicher"):
                parent["unsicher"] = True
            changes += 1
            continue
        out.append(pos)
    for pos in out:
        extras = pos.pop("_extras", None)
        if not extras:
            continue
        menge = _f(pos.get("menge"), 1) or 1
        alt_zsum, alt_preis = zeilensumme(pos), _f(pos.get("einzelpreis"))
        neu_zsum = round(alt_zsum + sum(zeilensumme(e) for e in extras), 2)
        basis = pos.get("artikel_basis") or pos.get("artikel") or "?"
        artikel = pos.get("artikel") or basis
        rest = artikel[len(basis):] if artikel.startswith(basis) else ""
        namen = " + ".join((e.get("artikel") or "?").strip() for e in extras)
        pos["artikel_basis"] = basis
        pos["artikel"] = f"{basis} + {namen}{rest}"
        pos["zeilensumme"] = neu_zsum
        pos["einzelpreis"] = round(neu_zsum / menge, 2)
        details = ", ".join(f"{(e.get('artikel') or '?').strip()} {_eur(zeilensumme(e))}" for e in extras)
        _note(pos, f"Extras eingerechnet: {details} (Bon-Zeile {_eur(alt_zsum)}, "
                   f"Einzelpreis vorher {_eur(alt_preis)})")
    data["positionen"] = out
    return changes


def fold_priced_addons(data, keep_separate=()):
    """Öffentlicher Einstieg für Provider ohne reconcile() (LLM-Pfad)."""
    return rule_fold_priced_addons(data, {"keep_separate": tuple(keep_separate)})


def rule_merge_quantity_lines(data, ctx):
    """Aldi-/Discounter-Muster: Mengenzeile wie "2 x 0,99" oder
    "0,524 kg x 2,99 EUR/kg" wurde als eigene Position erkannt. Sie wird der
    benachbarten Artikelzeile zugeordnet, deren Betrag dazu passt."""
    pos_list = data["positionen"]
    changes = 0
    i = 0
    while i < len(pos_list):
        parsed = _parse_qty_line(pos_list[i].get("artikel"))
        if not parsed:
            i += 1
            continue
        menge, preis = parsed
        betrag = round(menge * preis, 2)
        qty_has_amount = abs(zeilensumme(pos_list[i]) - betrag) <= ROW_TOL
        target = None
        for j in (i - 1, i + 1):
            if 0 <= j < len(pos_list) and not _parse_qty_line(pos_list[j].get("artikel")):
                z = zeilensumme(pos_list[j])
                # Nachbar trägt den passenden Betrag – oder Nachbar hat keinen
                # Betrag und die Mengenzeile trägt ihn (Betrag verrutscht)
                if abs(z - betrag) <= ROW_TOL or (abs(z) < 0.005 and qty_has_amount):
                    target = j
                    break
        if target is None:
            i += 1
            continue
        tpos = pos_list[target]
        _note(tpos, f"Mengenzeile „{pos_list[i].get('artikel', '').strip()}“ zugeordnet "
                    f"(vorher Menge {tpos.get('menge')}, Preis {_eur(tpos.get('einzelpreis'))})")
        tpos["menge"] = menge
        tpos["einzelpreis"] = preis
        tpos["zeilensumme"] = betrag if abs(zeilensumme(tpos)) < 0.005 else zeilensumme(tpos)
        del pos_list[i]
        changes += 1
        if target < i:
            continue  # Index i zeigt jetzt schon auf das nächste Element
    return changes


def _vat_ratio_match(small, big, rates):
    for r in rates:
        if small > 0 and abs(small * (1 + r) - big) <= 0.011:
            return r
    return None


def rule_quantity_from_ratio(data, ctx):
    """Menge fehlt/falsch gelesen: Einzelpreis stimmt, Zeilensumme ist ein
    ganzzahliges Vielfaches davon (BK-Extras ohne gedruckte Menge:
    "Extra Bacon 1,20 … 2,40")."""
    changes = 0
    for pos in data["positionen"]:
        if row_consistent(pos):
            continue
        preis, zsum = _f(pos.get("einzelpreis")), zeilensumme(pos)
        if preis <= 0 or zsum <= 0:
            continue
        menge = zsum / preis
        if round(menge) >= 2 and abs(menge - round(menge)) <= 0.01:
            _note(pos, f"Menge {int(round(menge))} aus Zeilensumme/Einzelpreis abgeleitet "
                       f"(vorher {_f(pos.get('menge'), 1):g})")
            pos["menge"] = int(round(menge))
            changes += 1
    return changes


def rule_row_net_gross(data, ctx):
    """Netto und Brutto in einer Zeile vermischt: Einzelpreis × Menge und
    Zeilensumme stehen genau im Verhältnis eines MwSt-Satzes zueinander.
    Es wird der größere (Brutto-)Wert für beide übernommen."""
    rates = ctx["rates"]
    changes = 0
    for pos in data["positionen"]:
        if row_consistent(pos):
            continue
        menge = _f(pos.get("menge"), 1) or 1
        preis, zsum = _f(pos.get("einzelpreis")), zeilensumme(pos)
        brutto_menge = menge * preis
        if _vat_ratio_match(zsum, brutto_menge, rates):
            _note(pos, f"Netto/Brutto vermischt → Brutto übernommen "
                       f"(Zeilensumme vorher {_eur(zsum)})")
            pos["zeilensumme"] = round(brutto_menge, 2)
            changes += 1
        elif _vat_ratio_match(brutto_menge, zsum, rates):
            _note(pos, f"Netto/Brutto vermischt → Brutto übernommen "
                       f"(Einzelpreis vorher {_eur(preis)})")
            pos["einzelpreis"] = round(zsum / menge, 2)
            changes += 1
    return changes


def rule_fold_priced_components(data, ctx):
    """Burger-King-Muster: unter einer Menü-Zeile stehen Komponenten MIT
    Preisen, die zusammen genau den Menüpreis ergeben (sonst doppelt
    gezählt). Das Menü bleibt als Pauschalbetrag, die Komponenten wandern in
    den Artikelnamen."""
    pos_list = data["positionen"]
    changes = 0
    i = 0
    while i < len(pos_list):
        parent = zeilensumme(pos_list[i])
        if parent <= 0:
            i += 1
            continue
        found = None
        running = 0.0
        for k in range(1, MAX_COMPONENTS + 1):
            if i + k >= len(pos_list):
                break
            z = zeilensumme(pos_list[i + k])
            if z <= 0:
                break
            running += z
            if k >= 2 and abs(running - parent) <= ROW_TOL:
                found = k
                break
            if running > parent + ROW_TOL:
                break
        if found:
            comps = pos_list[i + 1:i + 1 + found]
            names = ", ".join(p.get("artikel", "?") for p in comps)
            pos_list[i].setdefault("artikel_basis", pos_list[i].get("artikel", "?"))
            pos_list[i]["artikel"] = f"{pos_list[i].get('artikel', '?')}: {names}"
            _note(pos_list[i], f"{found} Menü-Komponenten mit Einzelpreisen "
                               f"(zusammen {_eur(parent)}) zusammengefasst")
            del pos_list[i + 1:i + 1 + found]
            changes += 1
        i += 1
    return changes


def rule_global_net_to_gross(data, ctx):
    """Alle Positionen wurden netto gelesen: Summe der Positionen trifft die
    Netto-Zwischensumme bzw. Bon-Summe/(1+Satz). Bei genau einem plausiblen
    Steuersatz werden alle Positionen auf Brutto hochgerechnet."""
    bon = data.get("bon_summe")
    if bon is None:
        return 0
    summe = sum(zeilensumme(p) for p in data["positionen"])
    if summe <= 0:
        return 0
    zw = data.get("zwischensumme")
    looks_net = zw is not None and abs(summe - _f(zw)) < SUM_TOL
    candidates = [r for r in ctx["rates"] if abs(summe * (1 + r) - _f(bon)) <= 0.01 * len(data["positionen"]) + SUM_TOL]
    if not (looks_net or candidates):
        return 0
    rates = candidates or ctx["rates"]
    if len(rates) != 1:
        return 0  # gemischte Sätze: ohne Beispiele keine sichere Zuordnung
    r = rates[0]
    for pos in data["positionen"]:
        old = zeilensumme(pos)
        pos["zeilensumme"] = round(old * (1 + r), 2)
        pos["einzelpreis"] = round(_f(pos.get("einzelpreis")) * (1 + r), 2)
        _note(pos, f"Netto → Brutto ({int(round(r * 100))} % MwSt, vorher {_eur(old)})")
    return len(data["positionen"])


# Reihenfolge = Anwendungsreihenfolge. Jede Regel wird nur behalten, wenn sie
# die Bewertung verbessert; am Ende zählt nur, ob alles exakt aufgeht.
RULES = [
    rule_quantity_from_ratio,
    rule_merge_quantity_lines,
    rule_row_net_gross,
    rule_fold_priced_components,
    rule_global_net_to_gross,
]


def reconcile(data, keep_separate=()):
    """Gibt (data, info) zurück. keep_separate: Begriffe (RECEIPT_EXCLUDE_ITEMS),
    deren Aufpreis-Extras NICHT in die Hauptposition gefaltet werden.
    info["status"]:
      "ok"         – passte bereits (ggf. nur strukturell aufgeräumt)
      "korrigiert" – erst nach automatischer Korrektur exakt plausibel
      "unklar"     – auch mit Korrekturen nicht eindeutig; Original bleibt
    """
    rates = sorted({s["satz"] for s in data.get("steuersaetze") or [] if s.get("satz")}) \
        or list(DEFAULT_VAT_RATES)
    ctx = {"rates": rates, "keep_separate": tuple(keep_separate)}

    # Rein strukturelle Aufräumschritte, die die Summe nicht verändern, werden
    # immer angewendet (Mengenzeilen ohne eigenen/mit verrutschtem Betrag
    # zuordnen, Menge aus "@Stückpreis" ableiten, preislose Menü-Komponenten
    # zusammenfassen, bezahlte Extras in die Hauptposition einrechnen).
    work = copy.deepcopy(data)
    candidate = copy.deepcopy(work)
    if rule_merge_quantity_lines(candidate, ctx):
        before, after = assess(work), assess(candidate)
        if abs(before["summe_positionen"] - after["summe_positionen"]) < 0.005 \
                and after["inkonsistente_zeilen"] <= before["inkonsistente_zeilen"]:
            work = candidate
    rule_at_price_quantity(work, ctx)
    rule_fold_zero_components(work, ctx)
    rule_fold_priced_addons(work, ctx)
    base = copy.deepcopy(work)

    if assess(work)["ok"]:
        return work, _info("ok", work)

    applied = []
    for rule in RULES:
        candidate = copy.deepcopy(work)
        if rule(candidate, ctx) and _score(candidate) < _score(work):
            work = candidate
            applied.append(rule.__name__)
            if assess(work)["ok"]:
                break

    if assess(work)["ok"]:
        info = _info("korrigiert", work)
        info["regeln"] = applied
        return work, info
    return base, _info("unklar", base)


def apply_shares(data, exclude_items=()):
    """Setzt den abzurechnenden Anteil: standardmäßig die volle Menge, für
    Positionen mit einem Ausschluss-Begriff (z. B. "Extra Bacon") 0.
    Vergleich: Groß-/Kleinschreibung egal, Begriff als Teiltext des Artikels.
    Gibt die Anzahl ausgeschlossener Positionen zurück."""
    terms = [t.strip().lower() for t in exclude_items if t and t.strip()]
    excluded = 0
    for pos in data["positionen"]:
        # nur der eigene Name zählt, nicht hineingefaltete Menü-Komponenten
        artikel = (pos.get("artikel_basis") or pos.get("artikel") or "").lower()
        hit = next((t for t in terms if t in artikel), None)
        if hit:
            pos["anteil"] = 0
            pos["ausschluss"] = hit
            excluded += 1
        else:
            pos["anteil"] = pos.get("menge") or 1
    return excluded


def _info(status, data):
    info = assess(data)
    info["status"] = status
    info["korrekturen"] = sum(1 for p in data["positionen"] if p.get("korrektur"))
    return info
