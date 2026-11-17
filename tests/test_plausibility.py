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
    assert amounts(out) == [9.99, 0.80]
    assert out["positionen"][0]["artikel"] == "Whopper Menü: Whopper, Pommes M, Cola M"
    assert out["positionen"][0]["korrektur"]


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


def test_bk_extras_inherit_quantity_from_at_price():
    out, info = reconcile(_bk_azure_like(True))
    assert info["status"] in ("ok", "korrigiert")
    got = [(p["artikel"].split(":")[0], p["menge"], round(plausibility.zeilensumme(p), 2))
           for p in out["positionen"]]
    assert got == [("King Fusion Oreo", 1, 3.49), ("Whopper", 2, 14.58), ("Extra Käse", 2, 2.40),
                   ("Extra Bacon", 2, 2.40), ("Whopper King Menü", 1, 12.19), ("Extra Käse", 1, 1.20),
                   ("Plant-based Whopper King Menü", 2, 24.18), ("Extra Käse", 2, 2.40)]
    menu = out["positionen"][4]
    # Komponenten hängen am Menü, nicht am bezahlten "Extra Käse" dazwischen
    assert menu["artikel"] == ("Whopper King Menü: Whopper, Curly Fries + Dip, "
                               "Extra Sour Cream Dip, 0,5 l Fuze Tea Lemon")


def test_bk_extras_without_at_price_fixed_by_ratio():
    out, info = reconcile(_bk_azure_like(False))
    assert info["status"] == "korrigiert"
    assert [p["menge"] for p in out["positionen"]][:4] == [1, 1, 2, 2]  # Whopper: 1×14,58 bleibt (kein @)
    assert info["inkonsistente_zeilen"] == 0


def test_exclusion_only_matches_own_name():
    out, _ = reconcile(_bk_azure_like(True))
    out["positionen"].append(pos("Bacon King", 8.99))
    out["positionen"][4]["artikel"] += ", Extra Bacon"   # als Komponente im Menü
    n = plausibility.apply_shares(out, ["Extra Bacon"])
    assert n == 1
    assert [p["anteil"] for p in out["positionen"]] == [1, 2, 2, 0, 1, 1, 2, 2, 1]


def test_shares_default_full_quantity():
    data = receipt([pos("Milch", 6.90, 6, 1.15)], 6.90)
    assert plausibility.apply_shares(data, []) == 0
    assert data["positionen"][0]["anteil"] == 6
