"""Actualiza data/orientacion_locales.json desde la base de asignacion vigente, conservando los ids L-nn.

Uso: python scripts/actualizar_locales.py [ruta_base.xlsx] [hoja]
Defecto: MONITOREO_REPORTE 041026_VF.xlsx (version final del 04/10/2026), hoja «04102026».

Llave de emparejamiento: distrito + nombre del local (sin tildes ni mayusculas) contra el archivo actual. Un local de la base que no
exista en el archivo actual aborta (no se inventa ubigeo ni coordenadas). Los locales que ya no estan en la base se retiran.
Solo salen datos de ubicacion y el conteo de orientadores: ni nombres, ni DNI, ni contactos, ni monitores.
"""
from __future__ import annotations

import csv
import json
import re
import sys
import unicodedata
from collections import Counter
from pathlib import Path

import openpyxl

RAIZ = Path(__file__).resolve().parent.parent
SALIDA = RAIZ / "data" / "orientacion_locales.json"
BASE = RAIZ.parent / "MONITOREO_REPORTE 041026_VF.xlsx"
HOJA = "04102026"


def nz(s) -> str:
    s = unicodedata.normalize("NFD", str(s or "")).encode("ascii", "ignore").decode().upper()
    return re.sub(r"\s+", " ", re.sub(r"[^A-Z0-9 ]", " ", s)).strip()


def leer_base(ruta: Path, hoja: str) -> list[tuple[str, str]]:
    """(distrito, local) por orientador; una tupla por fila de la base."""
    ws = openpyxl.load_workbook(ruta, data_only=True)[hoja]
    cab = [str(c).strip() if c else "" for c in next(ws.iter_rows(values_only=True))]
    ix = {h: i for i, h in enumerate(cab)}
    out = []
    for r in ws.iter_rows(min_row=2, values_only=True):
        if r[ix["N°"]] not in (None, "") and not str(r[ix["N°"]]).strip().isdigit():  # N° vacio = 2.a fila de un orientador en dos locales
            continue
        if r[ix["NOMBRES Y APELLIDOS"]]:
            out.append((str(r[ix["DISTRITO"]]).strip(), str(r[ix["NOMBRE DEL LOCAL"]]).strip()))
    return out


ADICIONALES = RAIZ / "data" / "entrada" / "asignaciones_adicionales.csv"


def leer_adicionales(ruta: Path = ADICIONALES) -> list[dict]:
    """Orientadores que atienden un local ademas del de la base: columnas nombre, distrito, local (ignorado por Git: lleva nombres)."""
    if not ruta.exists():
        return []
    with ruta.open(encoding="utf-8", newline="") as f:
        return [{"nombre": r["nombre"].strip(), "distrito": r["distrito"].strip(), "local": r["local"].strip()} for r in csv.DictReader(f)]


def construir(base: list[tuple[str, str]], actual: dict, adicionales: list[dict] = ()) -> dict:
    viejo = {(nz(r["distrito"]), nz(r["local"])): r for r in actual["registros"]}
    extra = [(nz(a["distrito"]), nz(a["local"])) for a in adicionales]
    conteo = Counter([(nz(d), nz(l)) for d, l in base] + extra)
    sig = max(int(r["id"][2:]) for r in actual["registros"]) + 1
    nuevos = {(nz(a["distrito"]), nz(a["local"])): a["local"] for a in adicionales}
    nuevos.update({(nz(d), nz(l)): l.strip() for d, l in base})
    for k, nombre in nuevos.items():
        if k in viejo:
            continue
        hermano = next((r for r in viejo.values() if nz(r["distrito"]) == k[0]), None)
        if hermano is None:
            raise SystemExit(f"Local sin ningun local del mismo distrito para tomar ubigeo y oficina (sin ubigeo ni coordenadas): {k}")
        print(f"AVISO: local nuevo {nombre} ({hermano['distrito']}) = {f'L-{sig:02d}'}; toma ubigeo y coordenadas de {hermano['local']} (afinar el punto si hace falta)")
        viejo[k] = {**{c: hermano[c] for c in ("ubigeo_inei", "departamento", "provincia", "distrito", "oficina_regional", "lon", "lat")},
                    "local": nombre, "id": f"L-{sig:02d}"}
        sig += 1
    regs = []
    for k, n in conteo.items():
        r = dict(viejo[k])
        r["n_orientadores"] = n
        regs.append(r)
    regs.sort(key=lambda r: int(r["id"][2:]))
    return {"fuente": "Base de asignacion de orientadores", "corte": None, "orientadores": sum(conteo.values()), "locales": len(regs),
            "distritos": len({r["ubigeo_inei"] for r in regs}), "registros": regs}


def main() -> None:
    ruta = Path(sys.argv[1]) if len(sys.argv) > 1 else BASE
    hoja = sys.argv[2] if len(sys.argv) > 2 else HOJA
    actual = json.loads(SALIDA.read_text(encoding="utf-8"))
    nuevo = construir(leer_base(ruta, hoja), actual, leer_adicionales())
    m = re.search(r"(\d{2})\.?(\d{2})\.?(\d{2})(\d{2})?$", hoja)
    nuevo["corte"] = (f"{m.group(1)}/{m.group(2)}/{m.group(3)}{m.group(4)}" if m.group(4) else f"{m.group(1)}/{m.group(2)}/20{m.group(3)}") if m else hoja
    antes = {r["id"]: r["n_orientadores"] for r in actual["registros"]}
    SALIDA.write_text(json.dumps(nuevo, ensure_ascii=False, indent=1), encoding="utf-8")
    retirados = sorted(set(antes) - {r["id"] for r in nuevo["registros"]})
    cambian = [(r["id"], antes.get(r["id"]), r["n_orientadores"]) for r in nuevo["registros"] if antes.get(r["id"]) != r["n_orientadores"]]
    print(f"{nuevo['orientadores']} orientadores en {nuevo['locales']} locales ({nuevo['distritos']} distritos), corte {nuevo['corte']}")
    print("retirados:", retirados, "| cambia el conteo (id, antes, ahora):", cambian)


if __name__ == "__main__":
    main()
