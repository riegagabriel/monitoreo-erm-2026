import datetime as dt
import sys
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "scripts"))
import matriz_a_dashboard as m  # noqa: E402

CTX = {"r2i": {"010406": "010506", "020813": "021014"},
       "nombres": {("AMAZONAS", "LUYA", "INGUILPATA"): "010506", ("ANCASH", "HUARI", "SAN MARCOS"): "021014", ("ANCASH", "HUARI", "CHAVIN DE HUANTAR"): "020803"},
       "geo": {"010506": dict(dep="AMAZONAS", prov="LUYA", dist="INGUILPATA", lon=-78.0, lat=-6.1), "021014": dict(dep="ANCASH", prov="HUARI", dist="SAN MARCOS", lon=-77.1, lat=-9.6),
               "020803": dict(dep="ANCASH", prov="HUARI", dist="CHAVIN DE HUANTAR", lon=-77.2, lat=-9.5)},
       "banderas": {"021014": ["REA-88"]}, "prohibidos": {"CHUQUILLANQUI", "ANTONY"},
       "locales": [{"id": "L-05", "ubigeo_inei": "010506", "local": "IE 18105 SAN JUAN"}]}


def caso(**kw):
    base = dict(id="C99", dep="Amazonas", prov="Luya", dist="Inguilpata", ubigeo="010406", local=None, fecha="2026-10-04", hora="09:30", desc="Mesa sin instalar.",
                gravedad="Alta", bandera="Sí", validado="SI", motivo="Riesgo de control de accesos.", canal="pre")
    return {**base, **kw}


class Verde(unittest.TestCase):
    def test_registro_de_la_jornada_en_formato_de_jornada_json(self):
        r = m.evaluar(caso(local="IE 18105"), CTX)
        self.assertEqual((r["destino"], r["estado"]), ("verde", "LISTO PARA PUBLICAR"))
        self.assertEqual(r["registros"], [{"id": "MAT-C99", "h": 9.5, "g": 2, "c": "pre", "u": "010506", "l": "L-05", "t": "Mesa sin instalar."}])

    def test_sin_validar_queda_pendiente(self):
        self.assertEqual(m.evaluar(caso(validado=None), CTX)["estado"], "PENDIENTE DE VALIDAR")

    def test_exige_canal_gravedad_y_hora(self):
        e = " ".join(m.evaluar(caso(canal=None, gravedad="Por evaluar", hora="tarde"), CTX)["errores"])
        self.assertIn("Canal dashboard", e)
        self.assertIn("gravedad", e)
        self.assertIn("Hora", e)

    def test_texto_largo_con_dni_o_nombre_interno_es_error(self):
        self.assertEqual(m.evaluar(caso(desc="x" * 281), CTX)["estado"], "ERROR")
        self.assertEqual(m.evaluar(caso(desc="El DNI 12345678 consta"), CTX)["estado"], "ERROR")
        self.assertEqual(m.evaluar(caso(desc="Informa Antony Pérez"), CTX)["estado"], "ERROR")

    def test_ubigeo_que_no_corresponde_al_distrito_es_error(self):
        self.assertEqual(m.evaluar(caso(ubigeo="020813"), CTX)["estado"], "ERROR")


class UbicarPorCodigo(unittest.TestCase):
    CTX2 = {**CTX, "por_reniec": {"010406": ("LUYA", "INGUILPATA")}}

    def test_error_de_tipeo_en_el_departamento_no_bloquea_si_el_codigo_coincide(self):
        r = m.evaluar(caso(dep="Amzonas"), self.CTX2)
        self.assertEqual((r["estado"], r["registros"][0]["u"]), ("LISTO PARA PUBLICAR", "010506"))

    def test_si_la_provincia_no_coincide_con_el_codigo_sigue_siendo_error(self):
        self.assertEqual(m.evaluar(caso(dep="Amzonas", prov="Bagua"), self.CTX2)["estado"], "ERROR")


class Roja(unittest.TestCase):
    def test_alerta_previa_con_bandera_roja(self):
        r = m.evaluar(caso(fecha="2026-10-02"), CTX)
        reg = r["registros"][0]
        self.assertEqual((r["destino"], r["estado"], reg["id"], reg["ubigeo_inei"], reg["conflictividad"], reg["fecha_recepcion"]),
                         ("roja", "LISTO PARA PUBLICAR", "MAT-C99", "010506", True, "02/10/2026"))

    def test_exige_motivo(self):
        self.assertEqual(m.evaluar(caso(fecha="2026-10-02", motivo=None), CTX)["estado"], "ERROR")

    def test_dos_distritos_generan_dos_registros_y_avisa_duplicado(self):
        r = m.evaluar(caso(fecha=dt.datetime(2026, 10, 1), dep="Áncash", prov="Huari", dist="San Marcos; Chavín de Huántar", ubigeo="020813"), CTX)
        self.assertEqual([x["id"] for x in r["registros"]], ["MAT-C99-1", "MAT-C99-2"])
        self.assertTrue(any("REA-88" in a for a in r["avisos"]))

    def test_sin_bandera_o_por_definir_no_se_publica(self):
        self.assertEqual(m.evaluar(caso(fecha="2026-10-02", bandera="No"), CTX)["estado"], "NO APLICA")
        self.assertEqual(m.evaluar(caso(fecha="2026-10-02", bandera="Por definir"), CTX)["estado"], "BANDERA POR DEFINIR")

    def test_roja_sin_distrito_es_error(self):
        self.assertEqual(m.evaluar(caso(fecha="2026-10-02", dist=None, ubigeo=None), CTX)["estado"], "ERROR")


if __name__ == "__main__":
    unittest.main()
