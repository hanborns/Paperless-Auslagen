#!/usr/bin/env python3
"""
Spielt die Plausibilitätskontrolle gegen gespeicherte Belege durch – ohne
Paperless und ohne Excel.

    python tools/replay.py testdata/bk_2026-11-17.azure.json
    python tools/replay.py testdata/                 # alle *.azure.json im Ordner
    python tools/replay.py testdata/aldi.pdf         # ruft Azure DI auf (braucht
                                                     # AZURE_DI_ENDPOINT/AZURE_DI_KEY
                                                     # in der Umgebung oder .env)
                                                     # und speichert aldi.azure.json
    python tools/replay.py X.azure.json --write-expected
        -> legt X.expected.json mit dem aktuellen Ergebnis an (zum Korrigieren
           von Hand; danach prüft tests/test_testdata.py darauf)
"""

import argparse
import base64
import json
import os
import sys
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
import plausibility  # noqa: E402

DOC_SUFFIXES = {".pdf", ".jpg", ".jpeg", ".png", ".tif", ".tiff"}


def load_env_file():
    env = Path(__file__).resolve().parent.parent / ".env"
    if env.exists():
        for line in env.read_text(encoding="utf-8").splitlines():
            line = line.strip()
            if line and not line.startswith("#") and "=" in line:
                k, v = line.split("=", 1)
                os.environ.setdefault(k.strip(), v.strip().strip('"').strip("'"))


def analyze_with_azure(path):
    import requests

    load_env_file()
    endpoint = os.environ.get("AZURE_DI_ENDPOINT", "").rstrip("/")
    key = os.environ.get("AZURE_DI_KEY", "")
    if not endpoint or not key:
        sys.exit("AZURE_DI_ENDPOINT und AZURE_DI_KEY müssen gesetzt sein (Umgebung oder .env).")
    params = {"api-version": os.environ.get("AZURE_DI_API_VERSION", "2024-11-30")}
    locale = os.environ.get("AZURE_DI_LOCALE", "de-DE")
    if locale:
        params["locale"] = locale
    resp = requests.post(
        f"{endpoint}/documentintelligence/documentModels/"
        f"{os.environ.get('AZURE_DI_MODEL', 'prebuilt-receipt')}:analyze",
        params=params,
        headers={"Ocp-Apim-Subscription-Key": key, "Content-Type": "application/json"},
        json={"base64Source": base64.standard_b64encode(path.read_bytes()).decode("ascii")},
        timeout=60)
    if resp.status_code >= 400:
        sys.exit(f"Azure DI API {resp.status_code}: {resp.text[:500]}")
    op = resp.headers["Operation-Location"]
    for _ in range(60):
        body = requests.get(op, headers={"Ocp-Apim-Subscription-Key": key}, timeout=30).json()
        if body.get("status") == "succeeded":
            return body["analyzeResult"]
        if body.get("status") == "failed":
            sys.exit(f"Azure DI Analyse fehlgeschlagen: {json.dumps(body)[:500]}")
        time.sleep(2)
    sys.exit("Azure DI: Timeout")


def print_table(title, data):
    print(f"\n  {title}")
    print(f"  {'Artikel':<44} {'Menge':>7} {'Preis':>8} {'Summe':>8}  Hinweis")
    for p in data["positionen"]:
        flags = []
        if not plausibility.row_consistent(p):
            flags.append("⚠ Zeile")
        if p.get("unsicher"):
            flags.append("⚠ unsicher")
        flags += ["🔧 " + k for k in p.get("korrektur") or []]
        print(f"  {str(p.get('artikel', '?'))[:44]:<44} {plausibility._f(p.get('menge'), 1):>7g} "
              f"{plausibility._f(p.get('einzelpreis')):>8.2f} {plausibility.zeilensumme(p):>8.2f}  "
              f"{'; '.join(flags)}")


def replay(json_path, write_expected):
    raw = json.loads(json_path.read_text(encoding="utf-8"))
    raw = raw.get("analyzeResult", raw)
    data = plausibility.normalize_azure_result(raw)
    result, info = plausibility.reconcile(data)
    print("=" * 100)
    print(f"{json_path.name}: {data.get('markt')} {data.get('datum')} | Bon-Summe {data.get('bon_summe')} "
          f"| Zwischensumme {data.get('zwischensumme')} | Steuer {data.get('steuer')} "
          f"| Sätze {[s['satz'] for s in data.get('steuersaetze', [])]}")
    a = plausibility.assess(data)
    print_table(f"Azure roh: Summe {a['summe_positionen']:.2f}, Differenz {a['differenz']}, "
                f"inkonsistent {a['inkonsistente_zeilen']}", data)
    print_table(f"Nach Kontrolle: Status {info['status'].upper()}, Summe {info['summe_positionen']:.2f}, "
                f"Differenz {info['differenz']}, Regeln {info.get('regeln', [])}", result)
    if write_expected:
        exp = json_path.with_name(json_path.name.replace(".azure.json", ".expected.json"))
        exp.write_text(json.dumps({
            "bon_summe": result.get("bon_summe"),
            "status": info["status"],
            "positionen": [{"artikel": p.get("artikel"), "menge": p.get("menge"),
                            "einzelpreis": p.get("einzelpreis"),
                            "zeilensumme": plausibility.zeilensumme(p)}
                           for p in result["positionen"]],
        }, ensure_ascii=False, indent=1), encoding="utf-8")
        print(f"\n  -> {exp} geschrieben (bitte von Hand prüfen/korrigieren)")


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("paths", nargs="+", type=Path)
    ap.add_argument("--write-expected", action="store_true")
    args = ap.parse_args()

    files = []
    for p in args.paths:
        if p.is_dir():
            files += sorted(p.glob("*.azure.json"))
            files += [d for d in sorted(p.iterdir()) if d.suffix.lower() in DOC_SUFFIXES
                      and not d.with_name(d.stem + ".azure.json").exists()]
        else:
            files.append(p)

    for f in files:
        if f.suffix.lower() in DOC_SUFFIXES:
            target = f.with_name(f.stem + ".azure.json")
            if not target.exists():
                print(f"Analysiere {f.name} mit Azure DI ...")
                target.write_text(json.dumps(analyze_with_azure(f), ensure_ascii=False, indent=1),
                                  encoding="utf-8")
            f = target
        replay(f, args.write_expected)


if __name__ == "__main__":
    main()
