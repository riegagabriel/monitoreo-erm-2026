import json
import re
import sys
import unittest
from collections import Counter
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "scripts"))
import pandas as pd  # noqa: E402
import procesar_orientadores as p  # noqa: E402


class Orientadores(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.filas = p.leer()
        cat = pd.read_parquet(p.LOOKUP, columns=["UBIGEO_INEI", "DEPARTAMEN", "PROVINCIA", "DISTRITO", "lon", "lat"])
        cls.cat = cat
        cls.pub, cls.priv = p.construir(cls.filas, cat, p.leer_equivalencias())

    def test_cifras_de_la_lista(self):
        self.assertEqual(len(self.filas), 52)
        self.assertEqual(len(self.pub), 31)
        self.assertEqual(len({l["ubigeo_inei"] for l in self.pub}), 29)
        self.assertEqual(sum(l["n_orientadores"] for l in self.pub), 52)
        self.assertEqual(Counter(l["n_orientadores"] for l in self.pub), Counter({2: 21, 1: 10}))

    def test_la_punta_es_callao(self):
        lp = [l for l in self.pub if l["distrito"] == "LA PUNTA"]
        self.assertEqual(len(lp), 1)
        self.assertEqual((lp[0]["departamento"], lp[0]["provincia"]), ("CALLAO", "CALLAO"))

    def test_ubigeo_texto_y_existente(self):
        existentes = set(self.cat["UBIGEO_INEI"].astype(str))
        for l in self.pub:
            self.assertRegex(l["ubigeo_inei"], r"^\d{6}$")
            self.assertIn(l["ubigeo_inei"], existentes)

    def test_ids_unicos(self):
        ids = [l["id"] for l in self.pub]
        self.assertEqual(len(ids), len(set(ids)))
        self.assertEqual(ids[0], "L-01")

    def test_json_publico_sin_datos_personales(self):
        txt = json.dumps(self.pub, ensure_ascii=False)
        self.assertIsNone(re.search(r"\d{8,}", txt))
        self.assertNotIn("@", txt)
        # Un nombre de orientador solo cuenta como fuga si dos o mas de sus palabras aparecen como palabras
        # completas en el JSON. Una sola coincidencia es un nombre comun que tambien esta en un topónimo o en el
        # nombre de un colegio (ALBERTO, CLARA, JORGE, HUANUCO); un apellido que sea subcadena de un topónimo
        # (HUANCA en HUANCARAYLLA) tampoco lo es.
        palabras = set(re.findall(r"[A-ZÁÉÍÓÚÑ]{4,}", txt.upper()))
        for f in self.filas:
            suyas = set(re.findall(r"[A-ZÁÉÍÓÚÑ]{4,}", f["NOMBRES Y APELLIDOS"].upper()))
            self.assertLess(len(suyas & palabras), 2, f"un orientador (fila de {f['DISTRITO']}) aparece en el JSON publico")

    def test_lista_privada_une_cada_orientador_a_su_local(self):
        self.assertEqual(len(self.priv), 52)
        self.assertEqual(Counter(x["local_id"] for x in self.priv)[self.pub[0]["id"]], self.pub[0]["n_orientadores"])

    def test_sin_equivalencia_aborta(self):
        with self.assertRaises(SystemExit):
            p.construir(self.filas, self.cat, {})


if __name__ == "__main__":
    unittest.main()
