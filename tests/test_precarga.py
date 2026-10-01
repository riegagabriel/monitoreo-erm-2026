import json
import sys
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "scripts"))
import precarga  # noqa: E402


class PrecargaReal(unittest.TestCase):
    """Corre contra los productos reales de DENUNCIAS_REA (solo lectura)."""

    @classmethod
    def setUpClass(cls):
        cls.casos = json.loads(precarga.CASOS.read_text(encoding="utf-8"))
        cls.res = json.loads(precarga.RESTITUIDOS.read_text(encoding="utf-8"))
        cls.ver = json.loads(precarga.VERIFICACIONES.read_text(encoding="utf-8"))
        cls.por = precarga.construir(cls.casos, cls.res, cls.ver)

    def test_cifras_acordadas(self):
        self.assertEqual(sum(len(f["alertas_rea"]) for f in self.por.values()), 8)
        self.assertEqual(sum(1 for f in self.por.values() if f["verificacion"]), 67)
        self.assertEqual(sum(1 for f in self.por.values() if f["restituidos"]), 165)

    def test_solo_denuncias_con_bandera(self):
        items = {a["item"] for f in self.por.values() for a in f["alertas_rea"]}
        self.assertEqual(items, {68, 71, 72, 75, 77, 78, 79, 80})

    def test_validar_pasa(self):
        precarga.validar(self.por, self.res, self.ver, self.casos)

    def test_sin_pii_ni_documentos(self):
        txt = json.dumps(self.por, ensure_ascii=False).lower()
        for prohibido in ("proveido", "dni", "apellido"):
            self.assertNotIn(prohibido, txt)

    def test_ubigeo_texto_seis_digitos(self):
        for u in self.por:
            self.assertEqual(len(u), 6)
            self.assertTrue(u.isdigit())


if __name__ == "__main__":
    unittest.main()
