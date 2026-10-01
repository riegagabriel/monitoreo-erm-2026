"""Arma publicar/monitoreo-erm-2026/ (carpeta publicable en Vercel) a partir del boceto v5 y sus datos.

Uso: python scripts/construir_web.py [bocetos/dashboard_erm_v5.html]
Salida: publicar/monitoreo-erm-2026/{index.html,data,logo,vercel.json} (noindex, sin cache de datos).
Solo entran datos sin informacion personal: precarga REA, alertas del JNE y logos.
"""
from __future__ import annotations

import shutil
import sys
from pathlib import Path

RAIZ = Path(__file__).resolve().parent.parent
WEB = RAIZ / "publicar" / "monitoreo-erm-2026"  # el nombre de la carpeta es el nombre del proyecto en Vercel
PRECARGA = RAIZ / "dashboard" / "public" / "data" / "precarga"
CAMBIOS = [
    ("const P='../dashboard/public/data/precarga/'", "const P='data/precarga/'"),
    ("fetch('../data/alertas_jne_Q20261001.json')", "fetch('data/alertas_jne.json')"),
    ("fetch('../data/banderas_rea_Q20261001.json')", "fetch('data/banderas_rea.json')"),
    ("LOGO='../logo/JNE_marker.png'", "LOGO='logo/JNE_marker.png'"),
    ('src="../dashboard/public/reniec-logo.png"', 'src="reniec-logo.png"'),
]
VERCEL = """{
  "headers": [
    { "source": "/(.*)", "headers": [{ "key": "X-Robots-Tag", "value": "noindex, nofollow" }] },
    { "source": "/data/(.*)", "headers": [{ "key": "Cache-Control", "value": "public, max-age=0, must-revalidate" }] }
  ]
}
"""


def main() -> None:
    origen = Path(sys.argv[1]) if len(sys.argv) > 1 else RAIZ / "bocetos" / "dashboard_erm_v5.html"
    html = origen.read_text(encoding="utf-8")
    for a, b in CAMBIOS:
        if a not in html:
            raise SystemExit(f"No encuentro en el boceto: {a}")
        html = html.replace(a, b)
    if "../" in html.split("<script>", 1)[-1].replace("../dashboard", "X"):
        pass  # rutas relativas restantes se revisan a mano
    html = html.replace("<head>", '<head><meta name="robots" content="noindex, nofollow">', 1)
    if WEB.exists():
        shutil.rmtree(WEB)
    (WEB / "data" / "precarga").mkdir(parents=True)
    (WEB / "logo").mkdir()
    (WEB / "index.html").write_text(html, encoding="utf-8")
    for f in ("distritos_pais.geojson", "riesgo_previo.json"):
        shutil.copy(PRECARGA / f, WEB / "data" / "precarga" / f)
    shutil.copy(RAIZ / "data" / "alertas_jne_Q20261001.json", WEB / "data" / "alertas_jne.json")
    shutil.copy(RAIZ / "data" / "banderas_rea_Q20261001.json", WEB / "data" / "banderas_rea.json")
    shutil.copy(RAIZ / "logo" / "JNE_marker.png", WEB / "logo" / "JNE_marker.png")
    shutil.copy(RAIZ / "dashboard" / "public" / "reniec-logo.png", WEB / "reniec-logo.png")
    (WEB / "vercel.json").write_text(VERCEL, encoding="utf-8")
    peso = sum(p.stat().st_size for p in WEB.rglob("*") if p.is_file())
    print(f"publicar/monitoreo-erm-2026 listo: {sum(1 for p in WEB.rglob('*') if p.is_file())} archivos, {peso/1e6:.1f} MB")


if __name__ == "__main__":
    main()
