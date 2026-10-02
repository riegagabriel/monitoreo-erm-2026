"""Coordenadas de los locales de orientacion: empareja cada local con su institucion educativa en el padron de IIEE.

Entrada : data/orientacion_locales.json (publico) y el padron D:\\BIENESTAR_DOCENTE\\scraper\\padron_full.csv (MINEDU/ESCALE).
Salida  : data/locales_coords.json (publico). Solo datos del servicio educativo: codigo de local, nombre, direccion,
          centro poblado, area, niveles y coordenadas. NO se copian director, telefono, correo ni RUC.
Regla   : mismo UBIGEO INEI; el nombre del local debe contener su numero (IE 16273) o todas sus palabras. Si hay varios
          codigos de local se elige el que tiene primaria o secundaria y mas niveles, y queda en confianza «revisar».
Lo que no empareja sale sin coordenadas: el dashboard lo deja en el centro del distrito y lo dice.
"""
from __future__ import annotations

import json
import re
import sys
import unicodedata
from pathlib import Path

import pandas as pd

RAIZ = Path(__file__).resolve().parent.parent
PADRON = Path(r"D:\BIENESTAR_DOCENTE\scraper\padron_full.csv")
LOCALES = RAIZ / "data" / "orientacion_locales.json"
SALIDA = RAIZ / "data" / "locales_coords.json"
COLS = ["codlocal", "nombreSE", "nivelModalidad", "estado", "direccion", "centroPoblado", "areaGeografica", "ubigeo", "latitud", "longitud"]
STOP = {"IE", "IEP", "I", "E", "DE", "DEL", "LA", "LOS", "LAS", "EL", "Y", "N"}


def norm(x) -> str:
    s = unicodedata.normalize("NFKD", str(x)).encode("ascii", "ignore").decode().upper()
    return re.sub(r"\s+", " ", re.sub(r"[^A-Z0-9 ]", " ", s)).strip()


def buscar(sub: pd.DataFrame, local: str) -> pd.DataFrame:
    toks = [t for t in norm(local).split() if t not in STOP]
    nums = [t.lstrip("0") for t in toks if t.isdigit()]
    if nums:
        return sub[sub.nn.apply(lambda s: all(n in [w.lstrip("0") for w in s.split()] for n in nums))]
    return sub[sub.nn.apply(lambda s: all(t in s.split() for t in toks))]


def main() -> int:
    df = pd.read_csv(PADRON, dtype=str, usecols=COLS)
    df = df[(df.estado == "Activo") & df.latitud.notna() & df.longitud.notna()].copy()
    df["nn"] = df.nombreSE.map(norm)
    locales = json.loads(LOCALES.read_text(encoding="utf-8"))["registros"]
    salida, sin = [], []
    for l in locales:
        sub = df[df.ubigeo == l["ubigeo_inei"]]
        m = buscar(sub, l["local"])
        confianza = "exacta"
        if m.empty:  # nombre con palabras distintas (abreviaturas, parroquial, etc.): coincide alguna palabra distintiva
            toks = [t for t in norm(l["local"]).split() if t not in STOP and not t.isdigit() and len(t) > 3]
            m = sub[sub.nn.apply(lambda s: sum(t in s.split() for t in toks) >= max(2, len(toks) // 2))]
            confianza = "revisar"
        if m.empty:
            sin.append(l["id"])
            salida.append({"id": l["id"], "lat": None, "lon": None, "confianza": "sin dato"})
            continue
        g = m.groupby("codlocal").agg(nombre=("nombreSE", "first"), direccion=("direccion", "first"), cp=("centroPoblado", "first"),
                                      area=("areaGeografica", "first"), lat=("latitud", "first"), lon=("longitud", "first"),
                                      niveles=("nivelModalidad", lambda x: sorted(set(x)))).reset_index()
        g["puntaje"] = g.niveles.map(lambda n: (any(k in " ".join(n) for k in ("Primaria", "Secundaria")), len(n)))
        g = g.sort_values("puntaje", ascending=False)
        if len(g) > 1:
            confianza = "revisar"
        r = g.iloc[0]
        salida.append({"id": l["id"], "codlocal": r.codlocal, "nombre_padron": r.nombre, "direccion": r.direccion, "centro_poblado": r.cp,
                       "area": r.area, "niveles": r.niveles, "lat": round(float(r.lat), 6), "lon": round(float(r.lon), 6), "confianza": confianza})
    SALIDA.write_text(json.dumps({"fuente": "Padrón de instituciones educativas (MINEDU, ESCALE)", "descarga": "28/08/2026", "registros": salida},
                                 ensure_ascii=False, indent=1), encoding="utf-8")
    con = [s for s in salida if s["lat"] is not None]
    print(f"{len(con)} de {len(salida)} locales con coordenadas ({sum(s['confianza'] == 'exacta' for s in con)} exactos, "
          f"{sum(s['confianza'] == 'revisar' for s in con)} por revisar); sin dato: {sin} -> {SALIDA.name}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
