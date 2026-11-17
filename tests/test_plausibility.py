"""Synthetische Fälle für plausibility.py (keine echten Belegdaten im Repo)."""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
import plausibility  # noqa: E402
from plausibility import normalize_azure_result, reconcile  # noqa: E402


def pos(artikel, zsum, menge=1, preis=None):
    return {"artikel": artikel, "menge": menge,
            "einzelpreis": round(zsum / menge, 2) if preis is None else preis,
            "zeilensumme": zsum, "unsicher": False}


def receipt(positionen, bon, **extra):
    return {"markt": "Test", "datum": "2026-11-17", "bon_summe": bon,
            "positionen": positionen, **extra}


def amounts(data):
    return [round(plausibility.zeilensumme(p), 2) for p in data["positionen"]]


# --- Grundverhalten ----------------------------------------------------------
def test_plausible_receipt_unchanged():
    data = receipt([pos("Milch", 6.90, 6, 1.15), pos("Brot", 2.49)], 9.39)
    out, info = reconcile(data)
    assert info["status"] == "ok"
    assert out["positionen"] == data["positionen"]
    assert info["korrekturen"] == 0


def test_unfixable_receipt_stays_original():
    data = receipt([pos("A", 1.00), pos("B", 2.00)], 5.55)
    out, info = reconcile(data)
    assert info["status"] == "unklar"
    assert amounts(out) == [1.00, 2.00]
    assert not any(p.get("korrektur") for p in out["positionen"])


def test_missing_bon_summe_is_unklar():
    _, info = reconcile(receipt([pos("A", 1.00)], None))
    assert info["status"] == "unklar"


def test_input_not_mutated():
    data = receipt([pos("Menü", 9.99), pos("Burger", 5.99), pos("Pommes", 4.00)], 9.99)
    snapshot = repr(data)
    reconcile(data)
    assert repr(data) == snapshot


# --- Burger King: Menü-Komponenten ------------------------------------------
def test_zero_price_components_folded_into_menu():
    data = receipt([pos("Whopper Menü", 9.99), pos("Whopper", 0.0, preis=0.0),
                    pos("Pommes M", 0.0, preis=0.0), pos("Cola M", 0.0, preis=0.0),
                    pos("Extra Bacon", 0.80)], 10.79)
    out, info = reconcile(data)
    assert info["status"] == "ok"
    # bezahltes Extra wird in die Menü-Zeile eingerechnet
    assert amounts(out) == [10.79]
    assert out["positionen"][0]["artikel"] == "Whopper Menü + Extra Bacon: Whopper, Pommes M, Cola M"
    assert out["positionen"][0]["artikel_basis"] == "Whopper Menü"
    assert len(out["positionen"][0]["korrektur"]) == 2


def test_priced_components_summing_to_menu_are_folded():
    data = receipt([pos("King Menü", 10.49), pos("Burger", 5.49), pos("Pommes", 2.50),
                    pos("Getränk", 2.50), pos("Sundae", 1.99)], 12.48)
    out, info = reconcile(data)
    assert info["status"] == "korrigiert"
    assert amounts(out) == [10.49, 1.99]
    assert "Burger, Pommes, Getränk" in out["positionen"][0]["artikel"]


def test_priced_components_not_folded_when_sum_already_fits():
    # Zufällig passende Folgepositionen dürfen nicht angefasst werden
    data = receipt([pos("A", 3.00), pos("B", 1.00), pos("C", 2.00)], 6.00)
    out, info = reconcile(data)
    assert info["status"] == "ok"
    assert amounts(out) == [3.00, 1.00, 2.00]


# --- Netto/Brutto -------------------------------------------------------------
def test_row_net_gross_mix_takes_gross():
    # Einzelpreis brutto 5.99, Zeilensumme netto 5.03 (19 %)
    data = receipt([pos("Burger", 5.03, 1, 5.99), pos("Wasser", 1.50)], 7.49)
    out, info = reconcile(data)
    assert info["status"] == "korrigiert"
    assert amounts(out) == [5.99, 1.50]
    assert "Netto/Brutto" in out["positionen"][0]["korrektur"][0]


def test_row_net_gross_mix_net_unit_price():
    # Einzelpreis netto 2.80 (7 %), Zeilensumme brutto 3.00
    data = receipt([pos("Pommes", 3.00, 1, 2.80)], 3.00)
    out, info = reconcile(data)
    assert info["status"] == "korrigiert"
    assert out["positionen"][0]["einzelpreis"] == 3.00


def test_all_items_net_single_rate():
    data = receipt([pos("A", 10.00), pos("B", 5.00)], 17.85,
                   zwischensumme=15.00, steuersaetze=[{"satz": 0.19}])
    out, info = reconcile(data)
    assert info["status"] == "korrigiert"
    assert amounts(out) == [11.90, 5.95]


def test_all_items_net_mixed_rates_not_guessed():
    data = receipt([pos("A", 10.00), pos("B", 5.00)], 16.25,
                   zwischensumme=15.00, steuersaetze=[{"satz": 0.07}, {"satz": 0.19}])
    _, info = reconcile(data)
    assert info["status"] == "unklar"


# --- Aldi: Mengenzeilen ---------------------------------------------------------
def test_quantity_line_merged_into_article_above():
    data = receipt([pos("Joghurt", 1.98), pos("2 x 0,99", 0.0, preis=0.0), pos("Brot", 1.29)], 3.27)
    out, info = reconcile(data)
    assert info["status"] in ("ok", "korrigiert")
    assert len(out["positionen"]) == 2
    assert out["positionen"][0]["menge"] == 2
    assert out["positionen"][0]["einzelpreis"] == 0.99


def test_quantity_line_with_duplicated_amount():
    data = receipt([pos("Joghurt", 1.98), pos("2 x 0,99", 1.98, 1, 1.98), pos("Brot", 1.29)], 3.27)
    out, info = reconcile(data)
    assert info["status"] == "korrigiert"
    assert amounts(out) == [1.98, 1.29]


def test_shifted_amount_on_quantity_line():
    data = receipt([pos("Joghurt", 0.0, preis=0.0), pos("2 x 0,99", 1.98, 1, 1.98), pos("Brot", 1.29)], 3.27)
    out, info = reconcile(data)
    assert info["status"] in ("ok", "korrigiert")
    assert out["positionen"][0]["artikel"] == "Joghurt"
    assert amounts(out) == [1.98, 1.29]


def test_weight_line():
    data = receipt([pos("Bananen", 1.57), pos("0,524 kg x 2,99 EUR/kg", 1.57, 1, 1.57)], 1.57)
    out, info = reconcile(data)
    assert info["status"] == "korrigiert"
    assert out["positionen"][0]["menge"] == 0.524


# --- Azure-Normalisierung ------------------------------------------------------
def _cur(v, conf=0.99):
    return {"valueCurrency": {"amount": v}, "confidence": conf}


def test_normalize_azure_tax_details():
    raw = {"documents": [{"fields": {
        "MerchantName": {"valueString": "Burger King"},
        "Total": _cur(10.70), "Subtotal": _cur(10.00), "TotalTax": _cur(0.70),
        "TaxDetails": {"valueArray": [{"valueObject": {
            "Rate": {"valueNumber": 7}, "NetAmount": _cur(10.00), "Amount": _cur(0.70)}}]},
        "Items": {"valueArray": [{"valueObject": {
            "Description": {"valueString": "Menü", "confidence": 0.95},
            "TotalPrice": _cur(10.70)}}]},
    }}]}
    data = normalize_azure_result(raw)
    assert data["zwischensumme"] == 10.00
    assert data["steuersaetze"] == [{"satz": 0.07, "netto": 10.00, "steuer": 0.70}]
    assert data["positionen"][0]["einzelpreis"] == 10.70
    assert data["positionen"][0]["unsicher"] is False


# --- Burger King: Struktur echter Bons (Preise öffentlich, keine Belegdaten) ---
def _bk_azure_like(extra_desc_with_at):
    """Bon-Aufbau wie bei BK: Hauptzeile mit '@Stückpreis', Extras ohne
    gedruckte Menge, Menü mit preislosen Komponenten und bezahltem Extra."""
    at = (lambda name, p: f"{name}@{p}") if extra_desc_with_at else (lambda name, p: name)
    return receipt([
        pos("King Fusion Oreo", 3.49),
        pos(at("Whopper", "7,29"), 14.58, 1, 14.58),
        pos(at("Extra Käse", "1,20"), 2.40, 1, 1.20),
        pos(at("Extra Bacon", "1,20"), 2.40, 1, 1.20),
        pos("Whopper King Menü", 12.19),
        pos("Whopper", 0.0, preis=0.0),
        pos("Extra Käse", 1.20),
        pos("Curly Fries + Dip", 0.0, preis=0.0),
        pos("Extra Sour Cream Dip", 0.0, preis=0.0),
        pos("0,5 l Fuze Tea Lemon", 0.0, preis=0.0),
        pos(at("Plant-based Whopper King Menü", "12,09"), 24.18, 1, 24.18),
        pos("Plant-based Whopper", 0.0, preis=0.0),
        pos(at("Extra Käse", "1,20"), 2.40, 1, 1.20),
    ], 62.84)


def test_bk_extras_added_to_main_position():
    out, info = reconcile(_bk_azure_like(True))
    assert info["status"] == "ok"
    assert info["inkonsistente_zeilen"] == 0
    got = [(p.get("artikel_basis", p["artikel"]), p["menge"], p["einzelpreis"],
            round(plausibility.zeilensumme(p), 2)) for p in out["positionen"]]
    assert got == [("King Fusion Oreo", 1, 3.49, 3.49),
                   ("Whopper", 2, 9.69, 19.38),
                   ("Whopper King Menü", 1, 13.39, 13.39),
                   ("Plant-based Whopper King Menü", 2, 13.29, 26.58)]
    whopper, menu = out["positionen"][1], out["positionen"][2]
    assert whopper["artikel"] == "Whopper + Extra Käse + Extra Bacon"
    assert "Extra Käse 2,40, Extra Bacon 2,40 (Bon-Zeile 14,58" in whopper["korrektur"][-1]
    # Komponenten hängen am Menü, nicht am bezahlten "Extra Käse" dazwischen
    assert menu["artikel"] == ("Whopper King Menü + Extra Käse: Whopper, Curly Fries + Dip, "
                               "Extra Sour Cream Dip, 0,5 l Fuze Tea Lemon")


def test_bk_extras_without_at_price():
    # ohne "@" bleibt die Hauptzeile 1 × 14,58; die inkonsistenten
    # Extra-Zeilen (1 × 1,20 = 2,40) gehen im Einrechnen auf
    out, info = reconcile(_bk_azure_like(False))
    assert info["status"] == "ok"
    assert info["inkonsistente_zeilen"] == 0
    assert amounts(out) == [3.49, 19.38, 13.39, 26.58]


def test_bk_real_receipt_structure_menu_first():
    # Aufbau wie ein echter Bon mit Gesamtsumme 46,06 (Menü zuerst)
    data = receipt([
        pos("Whopper King Menü", 12.19), pos("Whopper", 0.0, preis=0.0),
        pos("Extra Käse", 1.20), pos("große King Pommes + 1 Dip", 0.0, preis=0.0),
        pos("Extra Pommes-Sauce", 0.0, preis=0.0), pos("0,5 l Fuze Tea Lemon", 0.0, preis=0.0),
        pos("Whopper@7,29", 14.58, 2, 7.29), pos("Extra Käse@1,20", 2.40, 1, 1.20),
        pos("Extra Bacon@1,20", 2.40, 1, 1.20),
        pos("Plant-based Whopper King Menü", 12.09), pos("Plant-based Whopper", 0.0, preis=0.0),
        pos("Extra Käse", 1.20), pos("Curly Fries + Dip", 0.0, preis=0.0),
    ], 46.06)
    out, info = reconcile(data)
    assert info["status"] == "ok"
    got = [(p["menge"], p["einzelpreis"], p["zeilensumme"]) for p in out["positionen"]]
    assert got == [(1, 13.39, 13.39), (2, 9.69, 19.38), (1, 13.29, 13.29)]


def test_excluded_extra_stays_separate():
    out, info = reconcile(_bk_azure_like(True), keep_separate=["Extra Bacon"])
    assert info["status"] == "ok"
    names = [p.get("artikel_basis", p["artikel"]) for p in out["positionen"]]
    assert names[:3] == ["King Fusion Oreo", "Whopper", "Extra Bacon"]
    assert amounts(out)[:3] == [3.49, 16.98, 2.40]
    out["positionen"].append(pos("Bacon King", 8.99))
    assert plausibility.apply_shares(out, ["Extra Bacon"]) == 1
    assert [p["anteil"] for p in out["positionen"]] == [1, 2, 0, 1, 2, 1]


def test_addon_without_parent_is_kept():
    data = receipt([pos("Extra Käse", 1.20), pos("Cola", 2.50)], 3.70)
    out, _ = reconcile(data)
    assert amounts(out) == [1.20, 2.50]


def test_fold_priced_addons_for_llm_path():
    data = receipt([pos("Whopper", 7.29), pos("Extra Käse", 1.20)], 8.49)
    assert plausibility.fold_priced_addons(data) == 1
    assert amounts(data) == [8.49]
    assert data["positionen"][0]["einzelpreis"] == 8.49


def test_shares_default_full_quantity():
    data = receipt([pos("Milch", 6.90, 6, 1.15)], 6.90)
    assert plausibility.apply_shares(data, []) == 0
    assert data["positionen"][0]["anteil"] == 6


# --- ALDI: zweispaltig gelesene Bons (Beträge/Mengen beim Nachbarn) -----------
def _nb(artikel, zsum, menge=1, preis=None, ohne_text=False, offset=None):
    p = pos(artikel, zsum, menge, preis)
    p["_ohne_text"], p["_offset"] = ohne_text, offset
    return p


def test_aldi_shifted_unit_price_moves_to_neighbour():
    # "Salat 0,89 … 2,79" / "Sahne 5,34": 0,89 gehört zur Sahne (6 × 0,89)
    data = receipt([_nb("Salat", 2.79), _nb("Salat", 2.79, 1, 0.89), _nb("Sahne", 5.34)], 10.92)
    out, info = reconcile(data)
    assert info["status"] == "ok" and info["inkonsistente_zeilen"] == 0
    got = [(p["menge"], p["einzelpreis"], p["zeilensumme"]) for p in out["positionen"]]
    assert got == [(1, 2.79, 2.79), (1, 2.79, 2.79), (6, 0.89, 5.34)]


def test_aldi_amount_without_text_and_shifted_weight():
    content = "Käse\n2,59 A\nObst\n0,600 kg × 2,00 EUR/kg\n1,20 A\n5 x\n0,99\nBrot\n4,95 A\n"
    data = receipt([
        _nb("Käse", 2.59, offset=0),
        _nb("Obst", 0.0, preis=0.0, offset=12),
        _nb("?", 1.20, ohne_text=True, offset=40),
        _nb("Brot", 4.95, 0.6, 2.00, offset=content.index("Brot")),
    ], 8.74, _content=content)
    out, info = reconcile(data)
    assert info["status"] == "ok" and info["inkonsistente_zeilen"] == 0
    got = [(p["artikel"], p["menge"], p["einzelpreis"], p["zeilensumme"]) for p in out["positionen"]]
    # Obst nicht als "Menü-Komponente" in den Käse gefaltet
    assert got == [("Käse", 1, 2.59, 2.59), ("Obst", 0.6, 2.00, 1.20), ("Brot", 5, 0.99, 4.95)]


def test_content_quantity_with_faint_count():
    # Stückzahl blass -> nur "x 0,99" im Text; Menge aus Zeilensumme/Preis
    content = "Saft\n1,49 A\nx\n0,99\nBrot\n2,97 A\n"
    data = receipt([_nb("Saft", 1.49, offset=0), _nb("Brot", 2.97, offset=content.index("Brot"))],
                   4.46, _content=content)
    out, _ = reconcile(data)
    assert [(p["menge"], p["einzelpreis"]) for p in out["positionen"]] == [(1, 1.49), (3, 0.99)]


def test_content_quantity_ignores_weight_lines_and_pack_sizes():
    content = "Tee 10x200\n4,85 A\nTomaten\n0,500 kg x 9,90 EUR/kg\n4,95 A\n"
    data = receipt([_nb("Tee 10x200", 4.85, offset=0),
                    _nb("Tomaten", 4.95, 0.5, 9.90, offset=content.index("Tomaten"))],
                   9.80, _content=content)
    out, info = reconcile(data)
    assert info["status"] == "ok"
    assert [p["menge"] for p in out["positionen"]] == [1, 0.5]


def test_split_merged_description_takes_amount_from_content():
    content = "Bananen\n1,75 A\n0,880 kg × 1,99 EUR/kg\nOrangen 750g\n2,19 A\nSalat\n2,29 A\n"
    data = receipt([_nb("Bananen\nOrangen 750g", 1.75, 0.88, 1.99), _nb("Salat", 2.29)], 6.23,
                   _content=content)
    out, info = reconcile(data)
    assert info["status"] == "korrigiert"
    assert [(p["artikel"], p["zeilensumme"]) for p in out["positionen"]] == [
        ("Bananen", 1.75), ("Orangen 750g", 2.19), ("Salat", 2.29)]


def test_split_merged_description_rejected_if_not_exact():
    content = "Bananen\n1,75 A\nOrangen 750g\n9,99 A\n"
    data = receipt([_nb("Bananen\nOrangen 750g", 1.75)], 3.94, _content=content)
    out, info = reconcile(data)
    assert info["status"] == "unklar"
    assert len(out["positionen"]) == 1


# --- Kaufdatum ------------------------------------------------------------------
TSE_BON = """Bezahlung MASTERCARD
Betrag 54,90 EUR
07,11.2026 10:04 T-ID 123
x1726 BJ12/001/018 07,11.26,10:03
Unsere Öffnungszeiten: ab 01.01.2026 Mo-Sa
Kasse: International
- TSE nicht erreichbar -
Letzte erfolgreiche Signatur:
2026-11-02 12:47:33
TSE-Start: 2026-11-07T10:03:01.000Z
"""


def test_purchase_date_topmost_ignores_tse_and_ocr_commas():
    datum, hinweis, sicher = plausibility.pick_purchase_date(TSE_BON, "2026-02-11", "2026-11-16")
    assert (datum, sicher) == ("2026-11-07", True)
    assert "11.02.2026" in hinweis
    found = [d.isoformat() for d, _ in plausibility.find_dates(TSE_BON)]
    assert "2026-11-02" not in found and "2026-02-11" not in found


def test_purchase_date_iso_is_year_month_day():
    datum, hinweis, _ = plausibility.pick_purchase_date("Datum: 2026-11-02 12:00", None, "2026-11-17")
    assert datum == "2026-11-02" and hinweis is None


def test_purchase_date_skips_future_and_flags_old():
    content = "Gutschein bis 31.12.2027\nKauf 05.11.2026\n"
    assert plausibility.pick_purchase_date(content, None, "2026-11-17")[0] == "2026-11-05"
    datum, hinweis, sicher = plausibility.pick_purchase_date("03.01.2024", None, "2026-11-17")
    assert datum == "2024-01-03" and sicher is False and hinweis.startswith("⚠")


def test_purchase_date_falls_back_to_extracted():
    assert plausibility.pick_purchase_date("", "2026-11-05", "2026-11-17") == ("2026-11-05", None, True)
    datum, _, sicher = plausibility.pick_purchase_date("", "2016-11-26", "2026-11-17")
    assert datum == "2016-11-26" and sicher is False
