import sys
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "scripts"))
import actualizar_locales as al  # noqa: E402

ACTUAL = {"registros": [{"id": "L-15", "ubigeo_inei": "080910", "departamento": "CUSCO", "provincia": "LA CONVENCION", "distrito": "PICHARI",
                         "local": "IE LA VICTORIA", "oficina_regional": "J.R. 09 - CUSCO", "lon": -73.78, "lat": -12.41, "n_orientadores": 2}]}


class Adicionales(unittest.TestCase):
    def test_local_nuevo_hereda_ubigeo_del_distrito_y_cuenta_el_puesto(self):
        base = [("PICHARI", "IE LA VICTORIA"), ("PICHARI", "IE LA VICTORIA")]
        extra = [{"nombre": "X", "distrito": "PICHARI", "local": "IE PARQUE INDUSTRIAL"}]
        r = {x["id"]: x for x in al.construir(base, ACTUAL, extra)["registros"]}
        self.assertEqual((r["L-15"]["n_orientadores"], r["L-16"]["local"], r["L-16"]["n_orientadores"], r["L-16"]["ubigeo_inei"]),
                         (2, "IE PARQUE INDUSTRIAL", 1, "080910"))

    def test_local_adicional_en_distrito_desconocido_aborta(self):
        with self.assertRaises(SystemExit):
            al.construir([("PICHARI", "IE LA VICTORIA")], ACTUAL, [{"nombre": "X", "distrito": "OTRO", "local": "IE Z"}])

    def test_local_de_la_base_desconocido_sigue_abortando(self):
        with self.assertRaises(SystemExit):
            al.construir([("PICHARI", "IE NO EXISTE")], ACTUAL, [])


if __name__ == "__main__":
    unittest.main()
