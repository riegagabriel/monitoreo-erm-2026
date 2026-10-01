import json
import sys
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "scripts"))
import pandas as pd  # noqa: E402
import procesar_alertas_jne as p  # noqa: E402


class AlertasJNE(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.casos = p.leer_casos()
        cat = pd.read_parquet(p.LOOKUP, columns=["UBIGEO_INEI", "DEPARTAMEN", "PROVINCIA", "DISTRITO", "lon", "lat"])
        cls.cat = cat
        cls.alertas = p.construir(cls.casos, cat)

    def test_cifras_de_la_base(self):
        self.assertEqual(len(self.casos), 53)
        self.assertEqual(len(self.alertas), 32)

    def test_ubigeo_texto_de_seis_digitos_y_existente(self):
        existentes = set(self.cat["UBIGEO_INEI"].astype(str))
        for a in self.alertas:
            self.assertRegex(a["ubigeo_inei"], r"^\d{6}$")
            self.assertIn(a["ubigeo_inei"], existentes)

    def test_un_caso_se_ubica_por_departamento_y_distrito(self):
        aj = [a for a in self.alertas if a["resolucion_territorial"] != "exacta"]
        self.assertEqual([a["n_jne"] for a in aj], [52])  # Ite figura bajo Tarata; es de Jorge Basadre
        self.assertEqual(aj[0]["provincia"], "JORGE BASADRE")

    def test_solo_casos_con_si(self):
        con_si = {c["n"] for c in self.casos if c["violencia"] == "SI"}
        self.assertEqual({a["n_jne"] for a in self.alertas}, con_si)

    def test_salida_sin_datos_personales(self):
        txt = json.dumps(self.alertas, ensure_ascii=False)
        self.assertNotIn("DNI", txt)


if __name__ == "__main__":
    unittest.main()
