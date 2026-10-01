"""Precarga del contexto REA para el dashboard de monitoreo ERM 2026.

Lee (solo lectura) los productos ya publicables de DENUNCIAS_REA y escribe en
dashboard/public/data/precarga/:
  - riesgo_previo.json   por distrito: alerta REA, verificacion domiciliaria, restituidos
  - distritos_pais.geojson  1 891 distritos INEI simplificados (para ver cobertura)
  - departamentos.geojson

Solo entra lo acordado: denuncias CON bandera, distritos con verificacion y restituidos.
Nada con datos personales. UBIGEO siempre INEI.
"""
from __future__ import annotations

import json
import sys
from pathlib import Path

RAIZ = Path(__file__).resolve().parent.parent
REA = RAIZ.parent / "DENUNCIAS_REA"
DESTINO = RAIZ / "dashboard" / "public" / "data" / "precarga"

CASOS = REA / "sitio_web" / "public" / "data" / "casos.json"
RESTITUIDOS = REA / "data" / "processed" / "restituidos_Q20260925.json"
VERIFICACIONES = REA / "data" / "processed" / "verificaciones_Q20260921.json"
LOOKUP = REA / "data" / "geodata" / "distritos_lookup.parquet"
DEPARTAMENTOS = REA / "data" / "geodata" / "departamentos.geojson"

CLAVES_PII = {"dni", "apellido paterno", "apellido materno", "prenombres", "documento"}


def construir(casos: list[dict], restituidos: dict, verificaciones: dict) -> dict:
    """Une las tres fuentes por UBIGEO INEI. Funcion pura (probada en tests/)."""
    por_ubigeo: dict[str, dict] = {}

    def fila(u: str, dep: str, prov: str, dist: str) -> dict:
        return por_ubigeo.setdefault(
            u,
            {"ubigeo_inei": u, "departamento": dep, "provincia": prov, "distrito": dist,
             "alertas_rea": [], "verificacion": None, "restituidos": None},
        )

    for c in casos:
        if not c["alerta"] or not c.get("ubigeo_inei"):
            continue
        f = fila(c["ubigeo_inei"], c["departamento"], c["provincia"], c["distrito"])
        f["alertas_rea"].append({"item": c["item"], "motivo": c["alerta_motivo"]})

    for u, v in verificaciones["distritos"].items() if isinstance(verificaciones["distritos"], dict) else (
        (x["ubigeo_inei"], x) for x in verificaciones["distritos"]
    ):
        f = fila(u, v["departamento"], v["provincia"], v["distrito"])
        f["verificacion"] = {"tipo": v["tipo"], "domicilios": v.get("domicilios")}

    for u, r in restituidos["distritos"].items():
        f = fila(u, r["departamento"], r["provincia"], r["distrito"])
        f["restituidos"] = {"total": r["total"], "reniec": r["reniec"], "jne": r["jne"]}

    return por_ubigeo


def validar(por_ubigeo: dict, restituidos: dict, verificaciones: dict, casos: list[dict]) -> None:
    n_alertas = sum(len(f["alertas_rea"]) for f in por_ubigeo.values())
    esperado = sum(1 for c in casos if c["alerta"] and c.get("ubigeo_inei"))
    assert n_alertas == esperado, f"alertas {n_alertas} != {esperado}"
    n_ver = sum(1 for f in por_ubigeo.values() if f["verificacion"])
    assert n_ver == len(verificaciones["distritos"]), f"verificados {n_ver}"
    n_res = sum(1 for f in por_ubigeo.values() if f["restituidos"])
    assert n_res == len(restituidos["distritos"]), f"restituidos {n_res}"
    for f in por_ubigeo.values():
        assert len(f["ubigeo_inei"]) == 6 and f["ubigeo_inei"].isdigit(), f["ubigeo_inei"]
        assert not (CLAVES_PII & {k.lower() for k in f}), "PII en riesgo_previo"


def geometrias(destino: Path) -> int:
    import geopandas as gpd

    g = gpd.read_parquet(LOOKUP)
    g = g[["UBIGEO_INEI", "DEPARTAMEN", "PROVINCIA", "DISTRITO", "lon", "lat", "geometry"]].copy()
    g["geometry"] = g.geometry.simplify(0.01, preserve_topology=True)
    g = g.rename(columns={"UBIGEO_INEI": "u", "DEPARTAMEN": "dep", "PROVINCIA": "prov", "DISTRITO": "dist"})
    g["lon"] = g["lon"].round(4)
    g["lat"] = g["lat"].round(4)
    out = destino / "distritos_pais.geojson"
    out.write_text(g.to_json(drop_id=True, na="null", to_wgs84=True), encoding="utf-8")
    (destino / "departamentos.geojson").write_bytes(DEPARTAMENTOS.read_bytes())
    return len(g)


def main() -> None:
    casos = json.loads(CASOS.read_text(encoding="utf-8"))
    restituidos = json.loads(RESTITUIDOS.read_text(encoding="utf-8"))
    verificaciones = json.loads(VERIFICACIONES.read_text(encoding="utf-8"))
    por_ubigeo = construir(casos, restituidos, verificaciones)
    validar(por_ubigeo, restituidos, verificaciones, casos)
    DESTINO.mkdir(parents=True, exist_ok=True)
    salida = {
        "fuentes": {
            "denuncias_con_alerta": "casos.json (corte 30/09/2026)",
            "verificaciones": f"verificaciones_Q20260921 ({verificaciones['meta']['corte']})",
            "restituidos": f"restituidos_Q20260925 ({restituidos['corte']})",
        },
        "totales": {
            "alertas_rea": sum(len(f["alertas_rea"]) for f in por_ubigeo.values()),
            "distritos_verificados": sum(1 for f in por_ubigeo.values() if f["verificacion"]),
            "distritos_restituidos": sum(1 for f in por_ubigeo.values() if f["restituidos"]),
            "restituidos": restituidos["totales"],
        },
        "distritos": sorted(por_ubigeo.values(), key=lambda f: f["ubigeo_inei"]),
    }
    (DESTINO / "riesgo_previo.json").write_text(
        json.dumps(salida, ensure_ascii=False, separators=(",", ":")), encoding="utf-8"
    )
    n = geometrias(DESTINO)
    print(f"riesgo_previo.json: {len(por_ubigeo)} distritos | {salida['totales']}")
    print(f"distritos_pais.geojson: {n} distritos")


if __name__ == "__main__":
    sys.exit(main())
