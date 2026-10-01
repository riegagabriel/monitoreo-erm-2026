import json
import re
import sys
import unittest
from pathlib import Path

RAIZ = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(RAIZ / "scripts"))
import generar_distritos_apps_script as g  # noqa: E402


class DistritosAppsScript(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.d = g.construir()

    def test_totales(self):
        self.assertEqual(sum(len(v) for v in self.d.values()), 1891)
        self.assertEqual(len(self.d), 25)  # 24 departamentos y el Callao

    def test_ubigeo_texto_y_unico(self):
        ubi = [x[0] for v in self.d.values() for x in v]
        self.assertTrue(all(re.fullmatch(r"\d{6}", u) for u in ubi))
        self.assertEqual(len(ubi), len(set(ubi)))

    def test_ceros_a_la_izquierda(self):
        self.assertIn("010101", {x[0] for x in self.d["AMAZONAS"]})

    def test_limite_de_opciones_por_seccion(self):
        self.assertLess(max(len(v) for v in self.d.values()), 200)  # Lima tiene 171

    def test_archivo_generado_es_js_valido(self):
        txt = g.SALIDA.read_text(encoding="utf-8")
        m = re.search(r"const DISTRITOS = (\{.*\});", txt, re.S)
        self.assertIsNotNone(m)
        self.assertEqual(json.loads(m.group(1)), self.d)


if __name__ == "__main__":
    unittest.main()
