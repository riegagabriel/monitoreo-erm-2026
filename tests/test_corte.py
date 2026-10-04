import datetime as dt
import sys
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "scripts"))
import corte  # noqa: E402

INICIO = dt.datetime(2026, 10, 4)
CAB = ["Timestamp", "Seleccione su nombre y local de votación.", "¿Qué corte va a registrar?", "Hora de llegada al local de votación",
       "Hora de término de la orientación electoral",
       "Seleccione su nombre y local de votación. 2", "¿Qué corte va a registrar? 2", "Hora de llegada al local de votación 2",
       "Hora de término de la orientación electoral 2"]


def fila(ts, nombre, corte_t, llegada=None, termino=None, bloque="B"):
    r = [None] * len(CAB)
    r[0] = ts
    o = 5 if bloque == "B" else 1
    r[o], r[o + 1], r[o + 2], r[o + 3] = f"X · IE 1 · {nombre}", corte_t, llegada, termino
    return r


class Lectura(unittest.TestCase):
    def test_une_bloques_repetidos_y_lee_el_corte(self):
        filas = [fila(dt.datetime(2026, 10, 4, 7, 5), "ANA PEREZ", "Corte 1 · Mañana (~08:00)", dt.time(6, 45)),
                 fila(dt.datetime(2026, 10, 4, 7, 6), "LUIS RAMOS", "Corte 1 · Mañana (~08:00)", dt.time(7, 0), bloque="A")]
        r = corte.normalizar(CAB, filas)
        self.assertEqual([(x["k"], x["corte"], x["llegada"]) for x in r], [("ANA PEREZ", 1, dt.time(6, 45)), ("LUIS RAMOS", 1, dt.time(7, 0))])

    def test_hora_en_texto_hhmm_y_hora_invalida_aborta(self):
        ok = corte.normalizar(CAB, [fila(dt.datetime(2026, 10, 4, 7), "ANA PEREZ", "Corte 1 · M", "06:50")])
        self.assertEqual(ok[0]["llegada"], dt.time(6, 50))
        with self.assertRaises(SystemExit):
            corte.normalizar(CAB, [fila(dt.datetime(2026, 10, 4, 7), "ANA PEREZ", "Corte 1 · M", "temprano")])

    def test_filas_vacias_se_ignoran(self):
        self.assertEqual(corte.normalizar(CAB, [[None] * len(CAB)]), [])


class Avance(unittest.TestCase):
    BASE = {"ANA PEREZ": ["L-01"], "LUIS RAMOS": ["L-01"], "ROSA DIAZ": ["L-02"]}
    N = {"L-01": 2, "L-02": 1, "L-03": 1}

    def reg(self, ts, k, c, ll=None, te=None):
        return dict(ts=ts, k=k, corte=c, llegada=ll, termino=te)

    def test_cuenta_orientadores_distintos_y_excluye_pruebas(self):
        t = dt.datetime(2026, 10, 4, 7)
        reg = [self.reg(t, "ANA PEREZ", 1, dt.time(6, 40)), self.reg(t, "ANA PEREZ", 1, dt.time(6, 40)),  # reenvio
               self.reg(dt.datetime(2026, 10, 3, 20), "LUIS RAMOS", 1, dt.time(6, 0)),  # prueba
               self.reg(t, "ROSA DIAZ", 1, dt.time(7, 0))]
        av, info = corte.avance_por_local(reg, self.BASE, self.N, INICIO)
        self.assertEqual(av["L-01"], {"orientadores_base": 2, "llegaron": 1, "cierre": 0})
        self.assertEqual(av["L-02"]["llegaron"], 1)
        self.assertEqual(av["L-03"]["llegaron"], 0)
        self.assertEqual((info["envios_reales"], info["pruebas_excluidas"], info["orientadores_con_llegada"]), (3, 1, 2))

    def test_cierre_solo_con_corte_3_y_hora_de_termino(self):
        t = dt.datetime(2026, 10, 4, 16, 30)
        reg = [self.reg(t, "ANA PEREZ", 3, te=dt.time(16, 20)), self.reg(t, "LUIS RAMOS", 3)]
        av, _ = corte.avance_por_local(reg, self.BASE, self.N, INICIO)
        self.assertEqual(av["L-01"]["cierre"], 1)

    def test_nombre_fuera_de_la_base_aborta(self):
        with self.assertRaises(SystemExit):
            corte.avance_por_local([self.reg(dt.datetime(2026, 10, 4, 7), "NADIE", 1, dt.time(7, 0))], self.BASE, self.N, INICIO)

    def test_orientador_en_dos_locales_cuenta_en_ambos_y_una_vez_en_el_total(self):
        base = {"ANA PEREZ": ["L-01", "L-02"], "LUIS RAMOS": ["L-01"]}
        n = {"L-01": 2, "L-02": 1, "L-03": 0}
        av, info = corte.avance_por_local([self.reg(dt.datetime(2026, 10, 4, 7), "ANA PEREZ", 1, dt.time(7))], base, n, INICIO)
        self.assertEqual((av["L-01"]["llegaron"], av["L-02"]["llegaron"], info["orientadores_con_llegada"]), (1, 1, 1))

    def test_llegada_no_puede_superar_el_total_del_local(self):
        n = {"L-01": 1, "L-02": 1, "L-03": 1}
        t = dt.datetime(2026, 10, 4, 7)
        reg = [self.reg(t, "ANA PEREZ", 1, dt.time(7)), self.reg(t, "LUIS RAMOS", 1, dt.time(7))]
        with self.assertRaises(SystemExit):
            corte.avance_por_local(reg, self.BASE, n, INICIO)


class VariosLocales(unittest.TestCase):
    def test_etiqueta_sin_espacio_antes_del_punto_medio(self):
        cab = ["Timestamp", "Seleccione su nombre y local de votación.", "¿Qué corte va a registrar?", "Hora de llegada al local de votación", "Hora de término de la orientación electoral"]
        r = corte.normalizar(cab, [[dt.datetime(2026, 10, 4, 12, 29), "PICHARI · IE PARQUE INDUSTRIAL· ERICK FLORES", "Corte 1 · M", dt.time(7), None]])
        self.assertEqual((r[0]["k"], r[0]["dist"], r[0]["local"]), ("ERICK FLORES", "PICHARI", "IE PARQUE INDUSTRIAL"))

    def test_orientador_de_dos_locales_cuenta_en_el_que_eligio(self):
        base, n = {"ANA": ["L-01", "L-02"]}, {"L-01": 1, "L-02": 1}
        reg = [dict(ts=dt.datetime(2026, 10, 4, 7), k="ANA", corte=1, llegada=dt.time(7), termino=None, dist="D", local="DOS")]
        av, info = corte.avance_por_local(reg, base, n, INICIO, {("D", "DOS"): "L-02"})
        self.assertEqual((av["L-01"]["llegaron"], av["L-02"]["llegaron"], info["orientadores_con_llegada"]), (0, 1, 1))


class Consultas(unittest.TestCase):
    CAB = ["Timestamp", "Seleccione su nombre y local de votación.", "¿Qué corte va a registrar?", "Hora de llegada al local de votación",
           "Hora de término de la orientación electoral", "Consultas hasta el corte 2 · DNI vencido", "Consultas hasta el corte 2 · Otros tipos de consultas",
           "Consultas hasta el corte 2 · Restitución de domicilio — dashboard 2", "Consultas hasta el corte 2 · Otras consultas",
           "DNI de los ciudadanos fallecidos (opcional) · corte 2", "Consultas de todo el día · DNI vencido 2"]

    def fila(self, ts, c, vals):
        r = [None] * len(self.CAB)
        r[0], r[1], r[2] = ts, "D · LOCAL · ANA", f"Corte {c} · X"
        for i, v in vals.items():
            r[i] = v
        return r

    def test_une_nombres_de_tipo_de_los_dos_bloques_y_no_lee_el_dni(self):
        r = corte.normalizar(self.CAB, [self.fila(dt.datetime(2026, 10, 4, 12), 2, {5: 3, 6: 2, 7: 4, 8: 1, 9: "12345678"})])
        self.assertEqual(r[0]["consultas"], {"DNI vencido": 3, "Restitución de domicilio": 4, "Otras consultas": 3})

    def test_suma_todos_los_registros_de_los_cortes_2_y_3(self):
        filas = [self.fila(dt.datetime(2026, 10, 4, 12), 2, {5: 10}), self.fila(dt.datetime(2026, 10, 4, 13), 2, {5: 4, 7: 6}),
                 self.fila(dt.datetime(2026, 10, 4, 16), 3, {10: 20})]
        reg = corte.normalizar(self.CAB, filas)
        reg.append(dict(reg[0], k="LUIS", ts=dt.datetime(2026, 10, 4, 12), consultas={"DNI vencido": 5}))
        self.assertEqual(corte.resumen_consultas(reg[:2] + [reg[3]], INICIO), {"total": 25, "orientadores": 2, "por_tipo": {"DNI vencido": 19, "Restitución de domicilio": 6}})
        self.assertEqual(corte.resumen_consultas(reg[:3], INICIO)["por_tipo"], {"DNI vencido": 34, "Restitución de domicilio": 6})


class IncidenciasForms(unittest.TestCase):
    LOC = {"L-01": {"ubigeo_inei": "010101"}}
    POR = {("D", "LOCAL"): "L-01"}
    T = dt.datetime(2026, 10, 4, 8, 1, 10)

    def reg(self, inc="Sí"):
        return [dict(fila=30, ts=self.T, dist="D", local="LOCAL", inc=inc, det="texto crudo")]

    def csv(self, filas):
        import tempfile
        f = Path(tempfile.mkdtemp()) / "i.csv"
        f.write_text("\n".join(["id_envio,hora,local_id,tipo,resumen_publicable,validado,nota", *filas]), encoding="utf-8")
        return f

    def test_validada_sale_con_ubigeo_del_local_y_gravedad(self):
        inc, pend = corte.incidencias_forms(self.reg(), self.POR, self.LOC, self.csv(["ORI-20261004080110-L-01,08:01,L-01,A,Resumen.,SI,"]), INICIO)
        self.assertEqual((inc, pend), ([{"id": "ORI-20261004080110-L-01", "h": 8.02, "g": 2, "c": "ori", "u": "010101", "l": "L-01", "t": "Resumen."}], []))

    def test_sin_fila_en_el_csv_queda_pendiente_y_no_se_publica(self):
        inc, pend = corte.incidencias_forms(self.reg(), self.POR, self.LOC, self.csv([]), INICIO)
        self.assertEqual((inc, [p["fila"] for p in pend]), ([], [30]))

    def test_validado_no_excluye_sin_error(self):
        inc, pend = corte.incidencias_forms(self.reg(), self.POR, self.LOC, self.csv(["ORI-20261004080110-L-01,08:01,L-01,B,,NO,"]), INICIO)
        self.assertEqual((inc, pend), ([], []))

    def test_resumen_con_dni_o_largo_aborta(self):
        for texto in ("DNI 12345678 aqui", "x" * 281):
            with self.assertRaises(SystemExit):
                corte.incidencias_forms(self.reg(), self.POR, self.LOC, self.csv([f"ORI-20261004080110-L-01,08:01,L-01,B,{texto},SI,"]), INICIO)

    def test_id_que_no_existe_en_la_hoja_aborta(self):
        with self.assertRaises(SystemExit):
            corte.incidencias_forms(self.reg(), self.POR, self.LOC, self.csv(["ORI-20261004999999-L-01,08:01,L-01,B,Texto.,SI,"]), INICIO)


if __name__ == "__main__":
    unittest.main()
