"""Corte de la jornada: de la hoja de respuestas del formulario de orientadores a data/jornada.json (solo avance).

Uso: python scripts/corte.py [hoja_respuestas.xlsx]      (defecto: data/entrada/hoja_respuestas.xlsx)
Lee la pestaña RESP_ORIENTADORES por TITULO de columna (une los bloques repetidos con sufijo « 2»), cruza cada orientador con la base de
asignacion y cuenta, por local, cuantos orientadores reportaron llegada (corte 1) y cierre (corte 3). Las filas anteriores al inicio de la
jornada son ensayos y se excluyen. Un nombre que no este en la base aborta: nunca se descarta en silencio.
Incidencias: las del formulario entran solo si estan validadas en data/entrada/incidencias_forms.csv (id_envio, hora, local_id, tipo, resumen_publicable,
validado = SI) y las de la matriz desde data/entrada/incidencias_matriz.json (lo escribe matriz_a_dashboard.py --aplicar). Una incidencia del formulario
que no esta en el CSV se lista como pendiente de validar y NO se publica.
Salida: data/jornada.json, sin nombres ni monitores.
"""
from __future__ import annotations

import csv
import datetime as dt
import json
import re
import sys
from collections import defaultdict
from pathlib import Path

import openpyxl

sys.path.insert(0, str(Path(__file__).resolve().parent))
from actualizar_locales import BASE, HOJA, leer_adicionales, nz  # noqa: E402
from matriz_a_dashboard import revisar_texto  # noqa: E402

RAIZ = Path(__file__).resolve().parent.parent
ENTRADA = RAIZ / "data" / "entrada" / "hoja_respuestas.xlsx"
LOCALES = RAIZ / "data" / "orientacion_locales.json"
SALIDA = RAIZ / "data" / "jornada.json"
INC_FORMS = RAIZ / "data" / "entrada" / "incidencias_forms.csv"
INC_MATRIZ = RAIZ / "data" / "entrada" / "incidencias_matriz.json"
INICIO_JORNADA = dt.datetime(2026, 10, 4)
T_INC = r"¿Reporta alguna incidencia en el corte {c}\?(?: 2)?"
T_DET = r"Describa brevemente la incidencia del corte {c}(?: 2)?"
MAX_INCIDENCIA = 280
T_CONSULTA = re.compile(r"Consultas (?:hasta el corte 2|de todo el día) · (.+?)(?: \d)?")
TIPO_CONSULTA = {"Restitución de domicilio — dashboard": "Restitución de domicilio", "Otros tipos de consultas": "Otras consultas",
                 "Consulta sobre ciudadanos fallecidos": "Ciudadanos fallecidos"}
PESTANA = "RESP_ORIENTADORES"
T_NOMBRE = "Seleccione su nombre y local de votación."
T_CORTE = "¿Qué corte va a registrar?"
T_LLEGADA = "Hora de llegada al local de votación"
T_TERMINO = "Hora de término de la orientación electoral"


def _base_titulo(t) -> str:
    return re.sub(r"\s+\d+$", "", str(t or "").strip())


def _hora(v, fila: int, campo: str):
    if v in (None, ""):
        return None
    if isinstance(v, dt.datetime):
        return v.time()
    if isinstance(v, dt.time):
        return v
    m = re.fullmatch(r"(\d{1,2}):(\d{2})(?::\d{2})?", str(v).strip())
    if not m or int(m.group(1)) > 23:
        raise SystemExit(f"Fila {fila}: {campo} no es una hora HH:mm ({v!r})")
    return dt.time(int(m.group(1)), int(m.group(2)))


def _primero(r, cabecera, patron):
    """Primer valor no vacio entre las columnas cuyo titulo completo coincide con el patron (los bloques A y B repiten el titulo con sufijo « 2»)."""
    return next((r[i] for i, t in enumerate(cabecera) if re.fullmatch(patron, str(t or "").strip()) and r[i] not in (None, "")), None)


def _consultas(r, cabecera) -> dict[str, int]:
    """Cantidades por tipo de consulta de una fila (solo numeros; el DNI opcional de fallecidos nunca se lee)."""
    out: dict[str, int] = {}
    for i, t in enumerate(cabecera):
        m = T_CONSULTA.fullmatch(str(t or "").strip())
        v = r[i]
        if m and isinstance(v, (int, float)) and not isinstance(v, bool) and v >= 0:
            tipo = TIPO_CONSULTA.get(m.group(1), m.group(1))
            out[tipo] = out.get(tipo, 0) + int(v)
    return out


def normalizar(cabecera: list, filas: list) -> list[dict]:
    """Una fila de la hoja -> {fila, ts, k (nombre normalizado), corte (1-3), llegada, termino}. Une las columnas de titulo repetido."""
    cols = defaultdict(list)
    for i, t in enumerate(cabecera):
        cols[_base_titulo(t)].append(i)
    faltan = [t for t in (T_NOMBRE, T_CORTE, T_LLEGADA, T_TERMINO) if t not in cols]
    if faltan:
        raise SystemExit(f"No encuentro estos titulos en {PESTANA}: {faltan}")

    def val(r, t):
        return next((r[i] for i in cols[t] if r[i] not in (None, "")), None)

    out = []
    for n, r in enumerate(filas, start=2):
        if not any(c not in (None, "") for c in r):
            continue
        etiqueta, corte_t = val(r, T_NOMBRE), val(r, T_CORTE)
        m = re.search(r"Corte (\d)", str(corte_t or ""))
        if not etiqueta or not m:
            raise SystemExit(f"Fila {n}: falta el nombre o el corte")
        partes, c = [x for x in re.split(r"\s*·\s*", str(etiqueta)) if x], int(m.group(1))
        out.append({"fila": n, "ts": r[0], "k": nz(" ".join(partes[2:])), "dist": nz(partes[0]), "local": nz(partes[1]) if len(partes) > 1 else "", "corte": c,
                    "llegada": _hora(val(r, T_LLEGADA), n, "llegada"), "termino": _hora(val(r, T_TERMINO), n, "termino"),
                    "inc": _primero(r, cabecera, T_INC.format(c=c)), "det": _primero(r, cabecera, T_DET.format(c=c)), "consultas": _consultas(r, cabecera)})
    return out


def resumen_consultas(reg: list[dict], inicio: dt.datetime) -> dict:
    """Suma de todos los registros de consultas enviados en los cortes 2 y 3 (criterio del usuario, 04/10/2026)."""
    por_tipo: dict[str, int] = defaultdict(int)
    orientadores: set[str] = set()
    for r in reg:
        if r["ts"] >= inicio and r["corte"] in (2, 3) and r.get("consultas"):
            orientadores.add(r["k"])
            for t, n in r["consultas"].items():
                por_tipo[t] += n
    return {"total": sum(por_tipo.values()), "orientadores": len(orientadores), "por_tipo": dict(sorted(por_tipo.items(), key=lambda a: -a[1]))}


def _lids(r: dict, base: dict[str, list[str]], por_local: dict | None) -> list[str]:
    lids = base[r["k"]]
    if len(lids) > 1 and por_local:  # orientador de varios locales: cuenta en el que eligio en el formulario
        elegido = por_local.get((r.get("dist"), r.get("local")))
        lids = [elegido] if elegido in lids else lids
    return lids


def resumen_corte2(reg: list[dict], base: dict[str, list[str]], n_local: dict[str, int], locales: dict[str, dict], inicio: dt.datetime, por_local: dict | None = None) -> dict:
    """Orientadores, locales y distritos con al menos un registro del corte 2, y locales con todos sus puestos registrados."""
    reales = [r for r in reg if r["ts"] >= inicio and r["corte"] == 2]
    personas = {r["k"] for r in reales}
    por_lid = defaultdict(set)
    for r in reales:
        for lid in _lids(r, base, por_local):
            por_lid[lid].add(r["k"])
    completos = sum(1 for lid, ks in por_lid.items() if len(ks) >= n_local.get(lid, 0) > 0)
    return {"orientadores": len(personas), "locales": len(por_lid), "distritos": len({locales[l]["ubigeo_inei"] for l in por_lid}), "locales_completos": completos}


def avance_por_local(reg: list[dict], base: dict[str, list[str]], n_local: dict[str, int], inicio: dt.datetime, por_local: dict | None = None):
    """base: nombre normalizado -> ids de sus locales (uno o mas); n_local: id -> orientadores en la base. Devuelve (avance, info)."""
    reales = [r for r in reg if r["ts"] >= inicio]
    desconocidos = sorted({r["k"] for r in reales if r["k"] not in base})
    if desconocidos:
        raise SystemExit(f"Orientadores que no estan en la base de asignacion: {desconocidos}")
    lleg, cier = defaultdict(set), defaultdict(set)
    for r in reales:
        for lid in _lids(r, base, por_local):
            if r["corte"] == 1 and r["llegada"]:
                lleg[lid].add(r["k"])
            if r["corte"] == 3 and r["termino"]:
                cier[lid].add(r["k"])
    av = {}
    for lid, n in n_local.items():
        a, c = len(lleg[lid]), len(cier[lid])
        if a > n or c > n:
            raise SystemExit(f"Local {lid}: {a} llegadas y {c} cierres para {n} orientadores en la base")
        av[lid] = {"orientadores_base": n, "llegaron": a, "cierre": c}
    info = {"envios_reales": len(reales), "pruebas_excluidas": len(reg) - len(reales), "orientadores_con_llegada": len(set().union(*lleg.values())),
            "orientadores_con_cierre": len(set().union(*cier.values())), "ultimo_envio": max(r["ts"] for r in reales) if reales else None}
    return av, info


def incidencias_forms(reg: list[dict], por_local: dict, locales: dict[str, dict], ruta: Path, inicio: dt.datetime) -> tuple[list[dict], list[dict]]:
    """(incidencias validadas en formato del dashboard, incidencias del formulario aun sin revisar). Aborta ante un registro invalido."""
    hojas = {}
    if ruta.exists():
        with ruta.open(encoding="utf-8", newline="") as f:
            hojas = {x["id_envio"].strip(): x for x in csv.DictReader(f)}
    inc, pend, vistos = [], [], set()
    for r in reg:
        if r["ts"] < inicio or r["inc"] != "Sí":
            continue
        lid = por_local.get((r["dist"], r["local"]))
        if lid is None:
            raise SystemExit(f"Fila {r['fila']}: la incidencia viene de un local que no esta en orientacion_locales.json ({r['dist']} · {r['local']})")
        ide = f"ORI-{r['ts']:%Y%m%d%H%M%S}-{lid}"
        vistos.add(ide)
        x = hojas.get(ide)
        if x is None:
            pend.append({"id_envio": ide, "fila": r["fila"], "local_id": lid, "texto": r["det"]})
            continue
        if x["validado"].strip().upper() != "SI":
            continue
        texto, tipo, h = x["resumen_publicable"].strip(), x["tipo"].strip().upper(), _hora(x["hora"].strip(), r["fila"], "hora de la incidencia")
        errores = revisar_texto(texto, set())
        if not texto or len(texto) > MAX_INCIDENCIA:
            errores.append(f"el resumen debe tener entre 1 y {MAX_INCIDENCIA} caracteres")
        if tipo not in ("A", "B"):
            errores.append("tipo debe ser A o B")
        if x["local_id"].strip() != lid:
            errores.append(f"local_id {x['local_id']} no coincide con el del envio ({lid})")
        if errores:
            raise SystemExit(f"incidencias_forms.csv, {ide}: {errores}")
        inc.append({"id": ide, "h": round(h.hour + h.minute / 60, 2), "g": 2 if tipo == "A" else 1, "c": "ori", "u": locales[lid]["ubigeo_inei"], "l": lid, "t": texto})
    sobran = sorted(set(hojas) - vistos)
    if sobran:
        raise SystemExit(f"incidencias_forms.csv trae ids que no existen en la hoja de respuestas: {sobran}")
    return inc, pend


def incidencias_matriz(ruta: Path, locales: dict[str, dict]) -> list[dict]:
    if not ruta.exists():
        return []
    inc = json.loads(ruta.read_text(encoding="utf-8"))
    for x in inc:
        if set(x) != {"id", "h", "g", "c", "u", "l", "t"} or (x["l"] and x["l"] not in locales):
            raise SystemExit(f"incidencias_matriz.json: registro invalido {x.get('id')}")
    return inc


def cargar_base(locales: dict) -> dict[str, list[str]]:
    """nombre normalizado -> ids de local, desde la base de asignacion mas las asignaciones adicionales (llave del local: distrito + nombre)."""
    ws = openpyxl.load_workbook(BASE, data_only=True)[HOJA]
    cab = [str(c).strip() if c else "" for c in next(ws.iter_rows(values_only=True))]
    ix = {h: i for i, h in enumerate(cab)}
    por_local = {(nz(r["distrito"]), nz(r["local"])): r["id"] for r in locales["registros"]}
    out = {}
    for r in ws.iter_rows(min_row=2, values_only=True):
        try:
            int(str(r[ix["N°"]]).strip())
        except ValueError:
            continue
        nombre = r[ix["NOMBRES Y APELLIDOS"]]
        if not nombre:
            continue
        k = nz(nombre)
        lid = por_local.get((nz(r[ix["DISTRITO"]]), nz(r[ix["NOMBRE DEL LOCAL"]])))
        if lid is None:
            raise SystemExit(f"Local de la base sin id en orientacion_locales.json: {r[ix['DISTRITO']]} · {r[ix['NOMBRE DEL LOCAL']]}")
        if k in out:
            raise SystemExit(f"Nombre repetido en la base: {k}")
        out[k] = [lid]
    for a in leer_adicionales():
        k, lid = nz(a["nombre"]), por_local.get((nz(a["distrito"]), nz(a["local"])))
        if k not in out or lid is None:
            raise SystemExit(f"Asignacion adicional sin orientador en la base o sin local en orientacion_locales.json: {a}")
        if lid not in out[k]:
            out[k].append(lid)
    return out


def main() -> None:
    ruta = Path(sys.argv[1]) if len(sys.argv) > 1 else ENTRADA
    locales = json.loads(LOCALES.read_text(encoding="utf-8"))
    base = cargar_base(locales)
    n_local = {r["id"]: r["n_orientadores"] for r in locales["registros"]}
    puestos = sum(len(v) for v in base.values())
    if sum(n_local.values()) != puestos:
        raise SystemExit(f"orientacion_locales.json suma {sum(n_local.values())} puestos y la base {puestos}: ejecute actualizar_locales.py")
    filas = list(openpyxl.load_workbook(ruta, data_only=True)[PESTANA].iter_rows(values_only=True))
    reg = normalizar(list(filas[0]), filas[1:])
    por_id = {r["id"]: r for r in locales["registros"]}
    por_local = {(nz(r["distrito"]), nz(r["local"])): r["id"] for r in locales["registros"]}
    av, info = avance_por_local(reg, base, n_local, INICIO_JORNADA, por_local)
    inc_f, pend = incidencias_forms(reg, por_local, por_id, INC_FORMS, INICIO_JORNADA)
    incidencias = sorted(inc_f + incidencias_matriz(INC_MATRIZ, por_id), key=lambda x: x["h"])
    if len({x["id"] for x in incidencias}) != len(incidencias):
        raise SystemExit("Hay incidencias con el mismo id")
    ultimo = info["ultimo_envio"]
    completos = sum(1 for a in av.values() if a["llegaron"] == a["orientadores_base"])
    jornada = {"corte": ultimo.strftime("%d/%m/%Y %H:%M") if ultimo else "", "es_ejemplo": False,
               "fuente": "Formulario de orientadores (corte 1: llegada; corte 3: cierre)", "orientadores_base": len(base),
               "orientadores_con_llegada": info["orientadores_con_llegada"], "orientadores_con_cierre": info["orientadores_con_cierre"],
               "locales_completos": completos, "corte2": resumen_corte2(reg, base, n_local, por_id, INICIO_JORNADA, por_local), "consultas": resumen_consultas(reg, INICIO_JORNADA), "avance": av, "incidencias": incidencias}
    SALIDA.write_text(json.dumps(jornada, ensure_ascii=False, indent=1), encoding="utf-8")
    print(f"corte {jornada['corte']} | envios reales {info['envios_reales']} (pruebas excluidas: {info['pruebas_excluidas']})")
    print(f"llegada reportada: {info['orientadores_con_llegada']} de {len(base)} orientadores | cierre: {info['orientadores_con_cierre']}")
    print(f"locales con todos sus orientadores: {completos} de {len(av)} -> {SALIDA}")
    c2 = jornada["corte2"]
    print(f"corte 2: {c2['orientadores']} orientadores en {c2['locales']} locales y {c2['distritos']} distritos ({c2['locales_completos']} locales completos)")
    cs = jornada["consultas"]
    print(f"consultas atendidas (total estimado de la jornada): {cs['total']} de {cs['orientadores']} orientadores | " + ", ".join(f"{t} {n}" for t, n in cs["por_tipo"].items()))
    print(f"incidencias publicadas: {len(incidencias)} ({len(inc_f)} del formulario, {len(incidencias) - len(inc_f)} de la matriz, {sum(1 for x in incidencias if x['g'] == 2)} graves)")
    for p in pend:
        print(f"  PENDIENTE DE VALIDAR (no se publica) fila {p['fila']} {p['id_envio']}: {p['texto']}")


if __name__ == "__main__":
    main()
