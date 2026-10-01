"""Procesa de una vez todos los PDF nuevos de denuncias/: OCR (ocr_agy) y registro JSON (sistematizar_agy).

Uso:  python scripts/procesar_denuncias.py [--forzar] [--modelo-ocr ...] [--modelo-sis ...]
Salta los PDF que ya tienen denuncias/json/<nombre>.json (usa --forzar para rehacerlos).
Cada corrida es segura de repetir: el PDF original nunca se modifica.
"""
from __future__ import annotations

import argparse
import subprocess
import sys
import time
from pathlib import Path

RAIZ = Path(__file__).resolve().parent.parent
PDFS = RAIZ / "denuncias"
JSON_DIR = PDFS / "json"
AQUI = Path(__file__).resolve().parent


def correr(script: str, pdf: Path, extra: list[str]) -> int:
    return subprocess.run([sys.executable, str(AQUI / script), str(pdf), *extra]).returncode


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--forzar", action="store_true")
    ap.add_argument("--modelo-ocr")
    ap.add_argument("--modelo-sis")
    a = ap.parse_args()

    pdfs = sorted(PDFS.glob("*.pdf"))
    pendientes = [p for p in pdfs if a.forzar or not (JSON_DIR / f"{p.stem}.json").exists()]
    print(f"{len(pdfs)} PDF en denuncias/, {len(pendientes)} por procesar", flush=True)
    resultados = []
    for p in pendientes:
        t0 = time.time()
        print(f"\n=== {p.name}", flush=True)
        ocr = correr("ocr_agy.py", p, ["--modelo", a.modelo_ocr] if a.modelo_ocr else [])
        if ocr != 0:
            resultados.append((p.name, "FALLO EL OCR"))
            continue
        sis = correr("sistematizar_agy.py", p, ["--modelo", a.modelo_sis] if a.modelo_sis else [])
        resultados.append((p.name, "listo" if sis == 0 else "FALLO LA SISTEMATIZACION"))
        print(f"    {time.time() - t0:.0f} s", flush=True)
    print("\nResumen:")
    for n, r in resultados:
        print(f"  {r}: {n}")


if __name__ == "__main__":
    main()
