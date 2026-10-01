"""Extrae el texto de un PDF de denuncia: capa digital con PyMuPDF y OCR con IA (agy) en las paginas escaneadas.

Uso:
  python scripts/ocr_agy.py "denuncias/ARCHIVO.pdf" [--paginas 6-10,21] [--lote 6]
        [--modelo gemini-3.8-flash-high] [--solo-cuerpo N]

Salida (en denuncias/texto/):
  <stem>__pNN.txt   texto de cada pagina (digital u ocr)
  <stem>__manifiesto.json   metodo por pagina, caracteres, sha256 del PDF, modelo, duracion
El PDF original nunca se modifica. Los .txt contienen datos personales: no van a Git.
"""
from __future__ import annotations

import argparse
import hashlib
import json
import os
import shutil
import subprocess
import sys
import tempfile
import time
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path

import pymupdf

RAIZ = Path(__file__).resolve().parent.parent
SALIDA = RAIZ / "denuncias" / "texto"
def _buscar_agy() -> str:
    """AGY_BIN, luego el PATH, luego las ubicaciones habituales del instalador."""
    if os.environ.get("AGY_BIN"):
        return os.environ["AGY_BIN"]
    en_path = shutil.which("agy")
    if en_path:
        return en_path
    candidatos = [Path(os.environ.get("LOCALAPPDATA", "")) / "agy" / "bin" / "agy.exe",
                  Path.home() / "AppData" / "Local" / "agy" / "bin" / "agy.exe",
                  Path.home() / ".local" / "bin" / "agy.exe"]
    for c in candidatos:
        if c.is_file():
            return str(c)
    return str(candidatos[0])


AGY = _buscar_agy()
# El esfuerzo va en el slug del modelo (-low, -medium, -high); no se puede pasar --effort aparte.
MODELO = os.environ.get("AGY_MODEL", "gemini-3.8-flash-low")
MIN_TEXTO = 40  # caracteres: menos que esto = pagina escaneada

ESQUEMA = {
    "type": "object",
    "properties": {
        "paginas": {
            "type": "array",
            "items": {
                "type": "object",
                "properties": {
                    "archivo": {"type": "string"},
                    "texto": {"type": "string"},
                },
                "required": ["archivo", "texto"],
            },
        }
    },
    "required": ["paginas"],
}

PROMPT = (
    "Haz OCR fiel de las imagenes de paginas escaneadas que estan en esta carpeta. "
    "Archivos: {archivos}. Abre cada imagen y transcribe TODO su texto en espanol, respetando "
    "tildes y enes, en orden de lectura; incluye sellos, firmas legibles, fechas y numeros. "
    "Las tablas van fila por fila, con las celdas separadas por ' | '. Si algo es ilegible escribe [ilegible]. "
    "No resumas, no interpretes, no agregues nada que no este en la imagen. "
    "IMPORTANTE: lee cada imagen tal como esta, en una sola pasada. NO recortes, NO amplies, NO ajustes contraste, "
    "NO crees archivos, NO ejecutes codigo ni otras herramientas: solo abre la imagen y transcribe. "
    "Un sello o texto tenue que no se lea bien va como [ilegible] y sigues adelante. "
    "Devuelve solo JSON con la forma {{\"paginas\":[{{\"archivo\":\"p006.png\",\"texto\":\"...\"}}]}}."
)


def parsear_paginas(txt: str | None, total: int) -> list[int]:
    if not txt:
        return list(range(1, total + 1))
    out: list[int] = []
    for parte in txt.split(","):
        if "-" in parte:
            a, b = parte.split("-")
            out += range(int(a), int(b) + 1)
        else:
            out.append(int(parte))
    bad = [p for p in out if not 1 <= p <= total]
    if bad:
        raise SystemExit(f"Paginas fuera de rango (1 a {total}): {bad}")
    return out


def llamar_agy(carpeta: Path, archivos: list[str], modelo: str, timeout: str) -> dict:
    esq = carpeta / "esquema.json"
    esq.write_text(json.dumps(ESQUEMA), encoding="utf-8")
    cmd = [
        AGY, "-p", PROMPT.format(archivos=", ".join(archivos)),
        "--model", modelo, "--output-format", "json", "--json-schema", str(esq),
        "--add-dir", str(carpeta), "--dangerously-skip-permissions", "--print-timeout", timeout,
    ]
    r = subprocess.run(cmd, cwd=carpeta, capture_output=True, text=True, encoding="utf-8")
    if r.returncode != 0:
        raise RuntimeError(f"agy salio con codigo {r.returncode}: {r.stderr.strip()[-400:]}")
    envoltura = json.loads(r.stdout)
    if envoltura.get("status") != "SUCCESS":
        raise RuntimeError(f"agy status {envoltura.get('status')}")
    datos = _extraer_json(envoltura)
    if datos is None:
        SALIDA.mkdir(parents=True, exist_ok=True)
        dbg = SALIDA / f"_debug_respuesta_{carpeta.name}.json"
        dbg.write_text(json.dumps(envoltura, ensure_ascii=False, indent=1), encoding="utf-8")
        raise RuntimeError(
            f"La respuesta de agy no es el JSON esperado (claves: {sorted(envoltura)}, "
            f"largo de 'response': {len(envoltura.get('response') or '')}). Sobre guardado en {dbg}"
        )
    return {"datos": datos, "uso": envoltura.get("usage", {}), "seg": envoltura.get("duration_seconds")}


def _extraer_json(envoltura: dict) -> dict | None:
    """Prefiere 'structured_output' (lo que valida el esquema); si no, el primer objeto JSON de 'response'.
    raw_decode ignora el texto que el modelo pueda agregar despues del JSON."""
    so = envoltura.get("structured_output")
    if isinstance(so, str):
        try:
            so = json.loads(so)
        except ValueError:
            so = None
    if isinstance(so, dict) and "paginas" in so:
        return so
    cuerpo = (envoltura.get("response") or "").strip()
    i = cuerpo.find("{")
    if i < 0:
        return None
    try:
        obj, _ = json.JSONDecoder().raw_decode(cuerpo[i:])
    except ValueError:
        return None
    return obj if isinstance(obj, dict) and "paginas" in obj else None


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("pdf")
    ap.add_argument("--paginas")
    ap.add_argument("--lote", type=int, default=6)
    ap.add_argument("--modelo", default=MODELO)
    ap.add_argument("--timeout", default="10m")
    ap.add_argument("--paralelo", type=int, default=3, help="lotes simultaneos (1 = en serie)")
    ap.add_argument("--solo-cuerpo", type=int, metavar="N", help="OCR solo de las N primeras paginas escaneadas")
    a = ap.parse_args()

    pdf = Path(a.pdf)
    if not pdf.is_absolute():
        pdf = RAIZ / pdf
    if not Path(AGY).exists():
        raise SystemExit(f"No encuentro agy en {AGY}. Instalalo o define AGY_BIN.")
    doc = pymupdf.open(pdf)
    paginas = parsear_paginas(a.paginas, len(doc))
    sha = hashlib.sha256(pdf.read_bytes()).hexdigest()
    SALIDA.mkdir(parents=True, exist_ok=True)

    manifiesto = {"pdf": pdf.name, "sha256": sha, "modelo": a.modelo, "paginas": {}}
    escaneadas: list[int] = []
    for n in paginas:
        t = doc[n - 1].get_text().strip()
        if len(t) >= MIN_TEXTO:
            (SALIDA / f"{pdf.stem}__p{n:03d}.txt").write_text(t, encoding="utf-8")
            manifiesto["paginas"][n] = {"metodo": "digital", "caracteres": len(t)}
        else:
            escaneadas.append(n)
    if a.solo_cuerpo:
        omitidas = escaneadas[a.solo_cuerpo:]
        escaneadas = escaneadas[: a.solo_cuerpo]
        for n in omitidas:
            manifiesto["paginas"][n] = {"metodo": "omitida", "caracteres": 0}

    t0 = time.time()
    grupos = [escaneadas[i : i + a.lote] for i in range(0, len(escaneadas), a.lote)]

    def procesar(grupo: list[int]) -> tuple[list[int], dict, int, float]:
        """Un lote = una llamada a agy en su propia carpeta temporal (seguro para correr en paralelo)."""
        tmp = Path(tempfile.mkdtemp(prefix="ocr_agy_"))
        try:
            d = pymupdf.open(pdf)  # un documento por hilo: PyMuPDF no es seguro entre hilos
            for n in grupo:
                d[n - 1].get_pixmap(matrix=pymupdf.Matrix(1.6, 1.6)).save(tmp / f"p{n:03d}.png")
            for intento in (1, 2):
                try:
                    res = llamar_agy(tmp, [f"p{n:03d}.png" for n in grupo], a.modelo, a.timeout)
                    break
                except RuntimeError:
                    if intento == 2:
                        raise
            return grupo, {p["archivo"]: p["texto"] for p in res["datos"]["paginas"]}, res["uso"].get("total_tokens", 0), res["seg"]
        finally:
            shutil.rmtree(tmp, ignore_errors=True)

    tokens = 0
    with ThreadPoolExecutor(max_workers=max(1, a.paralelo)) as ex:
        for grupo, por_archivo, tk, seg in ex.map(procesar, grupos):
            tokens += tk
            for n in grupo:
                texto = por_archivo.get(f"p{n:03d}.png", "").strip()
                if texto:
                    (SALIDA / f"{pdf.stem}__p{n:03d}.txt").write_text(texto, encoding="utf-8")
                manifiesto["paginas"][n] = {"metodo": "ocr", "caracteres": len(texto)}
            print(f"lote {grupo[0]}-{grupo[-1]}: {sum(manifiesto['paginas'][n]['caracteres'] for n in grupo)} caracteres, {seg:.0f} s", flush=True)

    manifiesto["segundos_ocr"] = round(time.time() - t0, 1)
    manifiesto["tokens_ocr"] = tokens
    (SALIDA / f"{pdf.stem}__manifiesto.json").write_text(json.dumps(manifiesto, ensure_ascii=False, indent=1), encoding="utf-8")
    vacias = [n for n, v in manifiesto["paginas"].items() if v["metodo"] == "ocr" and v["caracteres"] == 0]
    print(f"listo: {len(paginas)} paginas, {len(escaneadas)} por OCR, {len(vacias)} vacias {vacias}")


if __name__ == "__main__":
    sys.exit(main())
