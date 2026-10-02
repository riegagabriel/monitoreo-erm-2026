"""Arma publicar/monitoreo-erm-2026/ (carpeta publicable en Vercel) a partir del ultimo boceto y sus datos.

Uso: python scripts/construir_web.py [bocetos/dashboard_erm_vN.html]   (sin argumento: el N mayor)
Salida: publicar/monitoreo-erm-2026/{index.html,data,logo,vercel.json} (noindex, sin cache de datos).
Solo entran datos sin informacion personal: precarga REA, alertas del JNE, banderas REA, locales de orientacion
(sin nombres, DNI ni contactos) y logos.
"""
from __future__ import annotations

import re
import shutil
import sys
from pathlib import Path

RAIZ = Path(__file__).resolve().parent.parent
WEB = RAIZ / "publicar" / "monitoreo-erm-2026"  # el nombre de la carpeta es el nombre del proyecto en Vercel
PRECARGA = RAIZ / "dashboard" / "public" / "data" / "precarga"
# Cambios obligatorios: si el boceto ya no trae el texto, el constructor aborta en vez de publicar rutas rotas.
CAMBIOS = [
    ("const P='../dashboard/public/data/precarga/'", "const P='data/precarga/'"),
    ("JNE_='../data/alertas_jne_Q20261001.json'", "JNE_='data/alertas_jne.json'"),
    ("BAN_='../data/banderas_rea_Q20261001.json'", "BAN_='data/banderas_rea.json'"),
    ("ORIE='../data/orientacion_locales.json'", "ORIE='data/orientacion_locales.json'"),
    ("LOGO='../logo/JNE_marker.png'", "LOGO='logo/JNE_marker.png'"),
    ('src="../dashboard/public/reniec-logo.png"', 'src="reniec-logo.png"'),
    ("const SELECTOR=true;", "const SELECTOR=false;"),  # sin selector de estilos ni modo de prueba en la web
]
# Cambios de rotulado (regex): se aplican solo si el boceto todavia trae el texto de boceto.
OPCIONALES = [
    (r"<title>Boceto v\d+ · ", "<title>"),
    (r"BOCETO v\d+ · solo datos reales cargados", "VISTA PREVIA · datos reales cargados"),
]
VERCEL = """{
  "headers": [
    { "source": "/(.*)", "headers": [{ "key": "X-Robots-Tag", "value": "noindex, nofollow" }] },
    { "source": "/data/(.*)", "headers": [{ "key": "Cache-Control", "value": "public, max-age=0, must-revalidate" }] }
  ]
}
"""
DATOS = [  # (origen, destino dentro de publicar/.../data)
    (RAIZ / "data" / "alertas_jne_Q20261001.json", "alertas_jne.json"),
    (RAIZ / "data" / "banderas_rea_Q20261001.json", "banderas_rea.json"),
    (RAIZ / "data" / "orientacion_locales.json", "orientacion_locales.json"),
]


def ultimo_boceto() -> Path:
    cand = [(int(m.group(1)), p) for p in (RAIZ / "bocetos").glob("dashboard_erm_v*.html") if (m := re.search(r"_v(\d+)\.html$", p.name))]
    return max(cand)[1]


def main() -> None:
    origen = Path(sys.argv[1]) if len(sys.argv) > 1 else ultimo_boceto()
    html = origen.read_text(encoding="utf-8")
    for a, b in CAMBIOS:
        if a not in html:
            raise SystemExit(f"No encuentro en el boceto: {a}")
        html = html.replace(a, b)
    for patron, nuevo in OPCIONALES:
        html = re.sub(patron, nuevo, html)
    restantes = re.findall(r"""['"(]\.\./[^'")]*""", html)
    if restantes:
        raise SystemExit(f"Quedan rutas relativas fuera de la carpeta publicable: {restantes[:5]}")
    html = html.replace("<head>", '<head><meta name="robots" content="noindex, nofollow">', 1)
    if WEB.exists():
        shutil.rmtree(WEB)
    (WEB / "data" / "precarga").mkdir(parents=True)
    (WEB / "logo").mkdir()
    (WEB / "index.html").write_text(html, encoding="utf-8")
    for f in ("distritos_pais.geojson", "riesgo_previo.json"):
        shutil.copy(PRECARGA / f, WEB / "data" / "precarga" / f)
    for src, dst in DATOS:
        shutil.copy(src, WEB / "data" / dst)
    shutil.copy(RAIZ / "logo" / "JNE_marker.png", WEB / "logo" / "JNE_marker.png")
    shutil.copy(RAIZ / "dashboard" / "public" / "reniec-logo.png", WEB / "reniec-logo.png")
    (WEB / "vercel.json").write_text(VERCEL, encoding="utf-8")
    peso = sum(p.stat().st_size for p in WEB.rglob("*") if p.is_file())
    print(f"{origen.name} -> publicar/monitoreo-erm-2026 listo: {sum(1 for p in WEB.rglob('*') if p.is_file())} archivos, {peso/1e6:.1f} MB")


if __name__ == "__main__":
    main()
