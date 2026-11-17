"""Regressionstests gegen ECHTE Belege aus dem lokalen, gitignorierten Ordner
testdata/ (nicht im Repo – in CI daher übersprungen).

Pro Fall:  testdata/<name>.azure.json    Rohantwort von Azure DI
           testdata/<name>.expected.json erwartetes Ergebnis
(anlegen mit: python tools/replay.py testdata/<name>.azure.json --write-expected,
 danach von Hand korrigieren)
"""

import json
import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
import plausibility  # noqa: E402

CASES = sorted((ROOT / "testdata").glob("*.expected.json")) if (ROOT / "testdata").exists() else []


@pytest.mark.skipif(not CASES, reason="keine lokalen Testdaten in testdata/")
@pytest.mark.parametrize("expected_path", CASES, ids=lambda p: p.name.replace(".expected.json", ""))
def test_real_receipt(expected_path):
    raw_path = expected_path.with_name(expected_path.name.replace(".expected.json", ".azure.json"))
    if not raw_path.exists():
        pytest.skip(f"{raw_path.name} fehlt (Azure-Rohantwort noch nicht abgelegt)")
    raw = json.loads(raw_path.read_text(encoding="utf-8"))
    expected = json.loads(expected_path.read_text(encoding="utf-8"))

    out, info = plausibility.reconcile(plausibility.normalize_azure_result(raw.get("analyzeResult", raw)))

    if "status" in expected:
        assert info["status"] == expected["status"]
    got = [(round(float(p.get("menge") or 1), 3), round(plausibility.zeilensumme(p), 2))
           for p in out["positionen"]]
    want = [(round(float(p.get("menge") or 1), 3), round(float(p["zeilensumme"]), 2))
            for p in expected["positionen"]]
    assert got == want
