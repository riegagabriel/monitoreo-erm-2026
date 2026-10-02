import sys
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "scripts"))
import pandas as pd  # noqa: E402
import procesar_banderas_rea as p  # noqa: E402


class BanderasREA(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cat = pd.read_parquet(p.LOOKUP, columns=["UBIGEO_INEI", "DEPARTAMEN", "PROVINCIA", "DISTRITO", "lon", "lat"])
        cls.reg = p.construir(p.leer(), cat)
        cls.con = [r for r in cls.reg if r["conflictividad"]]

    def test_cifras_de_control(self):
        self.assertEqual(len(self.reg), 15)
        self.assertEqual(len(self.con), 12)
        self.assertEqual(len({r["ubigeo_inei"] for r in self.con}), 10)
        self.assertEqual(sum(r["n_ciudadanos"] or 0 for r in self.reg), 380)

    def test_item_sin_bandera_es_el_76(self):
        self.assertEqual([r["item"] for r in self.reg if not r["conflictividad"]], [76, 86, 87])

    def test_san_marcos_reune_80_y_84(self):
        sm = sorted(r["item"] for r in self.con if r["ubigeo_inei"] == "021014")
        self.assertEqual(sm, [80, 84, 88])

    def test_ubigeo_texto_y_ceros(self):
        for r in self.reg:
            self.assertRegex(r["ubigeo_inei"], r"^\d{6}$")
        self.assertEqual(next(r for r in self.reg if r["item"] == 82)["ubigeo_inei"], "010104")

    def test_fechas_asumidas(self):
        self.assertEqual(sorted(r["item"] for r in self.reg if r["fecha_asumida"]), [80, 82, 83, 86, 87, 88, 89])

    def test_lista_vacia_no_es_cero(self):
        self.assertIsNone(next(r for r in self.reg if r["item"] == 75)["n_ciudadanos"])

    def test_no_salen_campos_internos(self):
        for r in self.reg:
            for prohibido in ("documento", "supuestos", "observacion_registrada", "omitio"):
                self.assertFalse(any(prohibido in k for k in r))


if __name__ == "__main__":
    unittest.main()
