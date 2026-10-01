"""Paso 3: del texto extraido a UN registro de denuncia (JSON), con IA (agy) y controles locales.

Uso:
  python scripts/sistematizar_agy.py "denuncias/ARCHIVO.pdf" [--modelo gemini-3.8-flash-medium]

Entrada : denuncias/texto/<stem>__pNNN.txt (salida de ocr_agy.py)
Salida  : denuncias/json/<stem>.json  (campos de la denuncia + controles; sin nombres en el resumen)

Los conteos (ciudadanos listados, actas) se calculan LOCALMENTE sobre las paginas de tabla y de actas;
a la IA solo se le pasa el cuerpo de la denuncia. El resumen publicable se rechaza si repite un nombre
propio tomado de las tablas o trae numeros largos (DNI).
"""
from __future__ import annotations

import argparse
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

RAIZ = Path(__file__).resolve().parent.parent
TEXTO = RAIZ / "denuncias" / "texto"
JSON_DIR = RAIZ / "denuncias" / "json"
sys.path.insert(0, str(Path(__file__).resolve().parent))
from ocr_agy import AGY  # noqa: E402  (reutiliza la busqueda de agy)

MODELO = os.environ.get("AGY_MODEL_SIS", "gemini-3.8-flash-medium")
DNI = re.compile(r"\b\d{8}\b")

CATEGORIAS = (
    "1 Impugnacion o depuracion de electores concretos; 2 Verificacion o fiscalizacion de domicilios de una localidad o grupo; "
    "3 Solicitud de padron o informacion; 4 Denuncia o alerta de golondrinaje, traslado fraudulento o riesgo de conflicto, "
    "sin pedir accion concreta sobre el padron; 5 Reclamo posterior a una verificacion o restitucion ya hecha por RENIEC"
)

ESQUEMA = {
    "type": "object",
    "properties": {
        "fecha_recepcion": {"type": "string"}, "hora_recepcion": {"type": "string"},
        "entidad_emisora": {"type": "string"}, "tipo_documento": {"type": "string"},
        "departamento": {"type": "string"}, "provincia": {"type": "string"}, "distrito": {"type": "string"},
        "n_ciudadanos_declarado": {"type": ["integer", "null"]},
        "categoria_rea": {"type": "integer"}, "conflictividad": {"type": "boolean"},
        "motivo_conflictividad": {"type": "string"}, "resumen_publicable": {"type": "string"},
        "pide_a_reniec": {"type": "string"}, "confianza": {"type": "string"}, "paginas_citadas": {"type": "string"},
    },
    "required": ["fecha_recepcion", "entidad_emisora", "tipo_documento", "departamento", "provincia", "distrito",
                 "categoria_rea", "conflictividad", "resumen_publicable", "pide_a_reniec", "confianza"],
}

PROMPT = (
    "Lee el archivo cuerpo.txt de esta carpeta: es el texto de una denuncia recibida por el RENIEC durante las "
    "Elecciones Regionales y Municipales 2026 (Peru). Devuelve SOLO un JSON con los campos del esquema. Reglas: "
    "(1) No inventes: si un dato no esta en el texto, deja cadena vacia o null. (2) fecha_recepcion en dd/mm/aaaa: la del "
    "sello de recibido si se lee; si no, la fecha del documento; hora_recepcion solo si se lee con claridad. "
    "(3) departamento, provincia y distrito del hecho denunciado (no del remitente), en MAYUSCULAS y sin tildes. "
    "(4) n_ciudadanos_declarado: la cantidad que el propio texto declara, no la que cuentes. "
    "(5) categoria_rea: " + CATEGORIAS + ". "
    "(6) conflictividad = true solo si el texto habla de violencia, amenazas, riesgo de conflicto o enfrentamientos; "
    "motivo_conflictividad en una frase. (7) resumen_publicable: maximo 280 caracteres, tercera persona, con la formula "
    "'presunto golondrinaje' cuando corresponda, mencionando el distrito, SIN nombres de personas, SIN DNI y SIN cifras de "
    "identidad. (8) pide_a_reniec: que solicita a RENIEC, en una frase. (9) confianza: alta, media o baja. "
    "(10) paginas_citadas: de donde sale lo esencial. No ejecutes codigo ni crees archivos."
)


def paginas(stem: str) -> dict[int, str]:
    out = {}
    for f in sorted(TEXTO.glob(f"{stem}__p*.txt")):
        out[int(f.stem.rsplit("__p", 1)[1])] = f.read_text(encoding="utf-8")
    return out


def es_tabla(t: str) -> bool:
    return len(set(DNI.findall(t))) >= 8


def es_acta(t: str) -> bool:
    return "ACTA DE CONSTATACI" in t.upper()


def nombres_de_tablas(pags: dict[int, str]) -> set[str]:
    """Palabras de 4+ letras que preceden al DNI en las filas de las tablas: nombres de ciudadanos."""
    palabras: set[str] = set()
    for t in pags.values():
        if not es_tabla(t):
            continue
        for linea in t.splitlines():
            m = DNI.search(linea)
            if m:
                previo = re.sub(r"^\W*\d+\W+", "", linea[: m.start()])
                palabras |= {w for w in re.findall(r"[A-ZÁÉÍÓÚÑ]{4,}", previo.upper())}
    return palabras


def llamar_agy(cuerpo: str, modelo: str) -> dict:
    tmp = Path(tempfile.mkdtemp(prefix="sis_agy_"))
    try:
        (tmp / "cuerpo.txt").write_text(cuerpo, encoding="utf-8")
        (tmp / "esquema.json").write_text(json.dumps(ESQUEMA), encoding="utf-8")
        cmd = [AGY, "-p", PROMPT, "--model", modelo, "--output-format", "json",
               "--json-schema", str(tmp / "esquema.json"), "--add-dir", str(tmp),
               "--dangerously-skip-permissions", "--print-timeout", "5m"]
        r = subprocess.run(cmd, cwd=tmp, capture_output=True, text=True, encoding="utf-8")
        if r.returncode != 0:
            raise RuntimeError(f"agy salio con codigo {r.returncode}: {r.stderr.strip()[-300:]}")
        env = json.loads(r.stdout)
        resp = (env.get("response") or "").strip()
        i, j = resp.find("{"), resp.rfind("}")
        return json.loads(resp[i : j + 1])
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("pdf")
    ap.add_argument("--modelo", default=MODELO)
    a = ap.parse_args()
    stem = Path(a.pdf).stem
    pags = paginas(stem)
    if not pags:
        raise SystemExit(f"No hay texto en {TEXTO} para '{stem}'. Corre antes ocr_agy.py.")

    tablas = [n for n, t in pags.items() if es_tabla(t)]
    actas = [n for n, t in pags.items() if es_acta(t)]
    cuerpo_pags = [n for n in sorted(pags) if n not in actas]

    def antes_de_tabla(t: str) -> str:
        """De una pagina de tabla solo se toma lo que va antes de la primera fila con DNI (sello, sumilla, petitorio)."""
        previas = []
        for linea in t.splitlines():
            if DNI.search(linea):
                break
            previas.append(linea)
        return "\n".join(previas[:-1] if len(previas) > 1 else previas)  # la ultima suele ser el nombre de la 1.a fila

    cuerpo = "\n\n".join(
        f"[pagina {n}]\n{antes_de_tabla(pags[n]) if n in tablas else pags[n]}" for n in cuerpo_pags
    ).strip()
    if not cuerpo.strip():
        raise SystemExit("No quedo texto de cuerpo tras separar tablas y actas.")

    dni_tabla = set()
    for n in tablas:
        dni_tabla |= set(DNI.findall(pags[n]))
    sin_ubicar = sum(1 for n in actas if "NO FUE POSIBLE" in pags[n].upper())

    rec = llamar_agy(cuerpo, a.modelo)

    # --- controles locales sobre el resumen publicable ---
    res = rec.get("resumen_publicable", "")
    problemas = []
    if len(res) > 280:
        problemas.append("resumen de mas de 280 caracteres")
    if re.search(r"\d{7,}", res):
        problemas.append("el resumen trae un numero largo (posible DNI)")
    repetidos = sorted({w for w in re.findall(r"[A-ZÁÉÍÓÚÑ]{4,}", res.upper())} & nombres_de_tablas(pags))
    if repetidos:
        problemas.append(f"el resumen repite palabras de nombres de las tablas: {repetidos}")
    declarado = rec.get("n_ciudadanos_declarado")
    if declarado and dni_tabla and abs(declarado - len(dni_tabla)) > 0.1 * declarado:
        problemas.append(f"n declarado ({declarado}) difiere de DNI unicos en tablas ({len(dni_tabla)})")

    rec["controles"] = {
        "archivo_origen": Path(a.pdf).name,
        "paginas_cuerpo": cuerpo_pags, "paginas_tabla": tablas, "paginas_acta": actas,
        "n_ciudadanos_listados": len(dni_tabla), "n_actas": len(actas), "n_actas_no_ubicado": sin_ubicar,
        "modelo": a.modelo, "problemas": problemas,
        "estado": "BORRADOR_CON_ALERTAS" if problemas else "BORRADOR",
    }
    JSON_DIR.mkdir(parents=True, exist_ok=True)
    (JSON_DIR / f"{stem}.json").write_text(json.dumps(rec, ensure_ascii=False, indent=1), encoding="utf-8")
    c = rec["controles"]
    print(f"{stem[:40]}: cat {rec.get('categoria_rea')} | {rec.get('distrito')}, {rec.get('provincia')}, {rec.get('departamento')} "
          f"| declarados {declarado} | listados {c['n_ciudadanos_listados']} | actas {c['n_actas']} ({sin_ubicar} sin ubicar) "
          f"| conflictividad {rec.get('conflictividad')} | estado {c['estado']}")
    for p in problemas:
        print("  ALERTA:", p)


if __name__ == "__main__":
    main()
