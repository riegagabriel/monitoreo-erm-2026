"""Genera formulario/distritos.gs con los distritos agrupados por departamento (UBIGEO INEI).

Fuente (solo lectura): DENUNCIAS_REA/data/geodata/distritos_lookup.parquet.
Cada opcion del formulario se escribe «PROVINCIA · DISTRITO (ubigeo)»; el ubigeo va al final para extraerlo en la hoja.
"""
from __future__ import annotations

import json
from pathlib import Path

import pandas as pd

RAIZ = Path(__file__).resolve().parent.parent
LOOKUP = RAIZ.parent / "DENUNCIAS_REA" / "data" / "geodata" / "distritos_lookup.parquet"
SALIDA = RAIZ / "formulario" / "distritos.gs"


def construir() -> dict[str, list[list[str]]]:
    df = pd.read_parquet(LOOKUP, columns=["UBIGEO_INEI", "DEPARTAMEN", "PROVINCIA", "DISTRITO"]).copy()
    for c in df.columns:
        df[c] = df[c].astype(str).str.strip()
    assert df["UBIGEO_INEI"].str.fullmatch(r"\d{6}").all(), "ubigeo no es texto de 6 digitos"
    assert not df["UBIGEO_INEI"].duplicated().any(), "ubigeo duplicado"
    out: dict[str, list[list[str]]] = {}
    for dep, g in df.sort_values(["DEPARTAMEN", "PROVINCIA", "DISTRITO"]).groupby("DEPARTAMEN", sort=True):
        out[dep] = [[r.UBIGEO_INEI, r.PROVINCIA, r.DISTRITO] for r in g.itertuples()]
    return out


def main() -> None:
    datos = construir()
    SALIDA.parent.mkdir(parents=True, exist_ok=True)
    cuerpo = json.dumps(datos, ensure_ascii=False, separators=(",", ":"))
    SALIDA.write_text(
        "// Generado por scripts/generar_distritos_apps_script.py. No editar a mano.\n"
        "// Estructura: { DEPARTAMENTO: [[ubigeo_inei, provincia, distrito], ...] }\n"
        f"const DISTRITOS = {cuerpo};\n",
        encoding="utf-8",
    )
    n = sum(len(v) for v in datos.values())
    print(f"{len(datos)} departamentos, {n} distritos, max por departamento {max(len(v) for v in datos.values())} -> {SALIDA}")


if __name__ == "__main__":
    main()
