"""
Tests unitarios para 4_pagos/efectivo/main.py.

Cobertura:
  · normalizar_llave   — formato, mayúsculas, espacios
  · limpiar_monto      — strings, comas, invalidos
  · es_ok              — variantes de Sí / No / vacío
  · comparar_fresco    — todos los tipos de resultado y pre-llenado
  · cargar_registro    — filas válidas, inválidas, archivo faltante
  · cargar_discrepancias  — doble cabecera, OK leído, archivo faltante
  · cargar_pagos_existentes — doble cabecera, archivo faltante
"""
import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from openpyxl import Workbook

sys.path.insert(0, str(Path(__file__).parent.parent))
import main


# ── Helpers ──────────────────────────────────────────────────────────────────

def _registro_wb(filas: list) -> Workbook:
    """Workbook con cabecera simple: MZ | LT | MONTO | FECHA."""
    wb = Workbook()
    ws = wb.active
    ws.append(["MZ", "LT", "MONTO", "FECHA"])
    for f in filas:
        ws.append(f)
    return wb

def _discrepancias_wb(filas: list) -> Workbook:
    """Workbook de discrepancias con doble cabecera (igual que exportar_discrepancias_xlsx)."""
    wb = Workbook()
    ws = wb.active
    ws.title = "discrepancias"
    ws.append(["¿CUÁL ES EL LOTE?", "", "", "", "",
               "¿QUÉ DICE CADA CUADERNO?", "", "",
               "¿CUÁL ES EL CORRECTO?", "", "",
               "¿CUÁNDO?", ""])
    ws.append(["TIPO", "MZ", "LT", "LLAVE", "",
               "MONTO_MIA", "MONTO_AMIGA", "",
               "MONTO_FINAL", "OK", "",
               "FECHA", "CICLO"])
    for f in filas:
        ws.append(f)
    return wb

def _pagos_efectivo_wb(filas: list) -> Workbook:
    """Workbook de pagos_efectivo con doble cabecera."""
    wb = Workbook()
    ws = wb.active
    ws.title = "pagos_efectivo"
    ws.append(["¿DÓNDE VIVE?", "", "", "", "¿CUÁNTO PAGÓ?", "", "¿CUÁNDO?", ""])
    ws.append(["MZ", "LT", "LLAVE", "", "MONTO", "", "FECHA", "CICLO"])
    for f in filas:
        ws.append(f)
    return wb

def _fila_disc(tipo="diff_monto", mz="A", lt="8C",
               monto_mia=38.0, monto_amiga=40.0, monto_final=40.0,
               ok="", fecha="03/05/2026", ciclo=1) -> list:
    """Fila con separadores vacíos en posiciones correctas."""
    llave = main.normalizar_llave(mz, lt)
    return [tipo, mz, lt, llave, "", monto_mia, monto_amiga, "",
            monto_final, ok, "", fecha, ciclo]

def _fila_pago(mz="A", lt="8C", monto=40.0, fecha="03/05/2026", ciclo=1) -> list:
    llave = main.normalizar_llave(mz, lt)
    return [mz, lt, llave, "", monto, "", fecha, ciclo]

def _reg(mz, lt, monto, fecha="03/05/2026") -> dict:
    """Dict de registro para comparar_fresco (ya procesado)."""
    return {"llave": main.normalizar_llave(mz, lt),
            "mz": mz, "lt": lt, "monto": monto, "fecha": fecha}


class _TmpWbMixin:
    """Guarda un Workbook en archivo temporal y lo registra para limpieza."""
    def _tmp_wb(self, wb: Workbook) -> Path:
        f = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
        f.close()
        p = Path(f.name)
        wb.save(p)
        self.addCleanup(p.unlink, missing_ok=True)
        return p


# ── 1. normalizar_llave ──────────────────────────────────────────────────────

class TestNormalizarLlave(unittest.TestCase):

    def test_resultado_mayusculas(self):
        self.assertEqual(main.normalizar_llave("a", "8c"), "A-8C")

    def test_elimina_espacios(self):
        self.assertEqual(main.normalizar_llave("  A  ", " 8C "), "A-8C")

    def test_formato_mz_guion_lt(self):
        llave = main.normalizar_llave("Z", "3")
        partes = llave.split("-")
        self.assertEqual(partes[0], "Z")
        self.assertEqual(partes[1], "3")

    def test_mismo_mz_diferente_lt_distintos(self):
        self.assertNotEqual(main.normalizar_llave("A", "1"),
                            main.normalizar_llave("A", "2"))

    def test_lote_alfanumerico_preservado(self):
        self.assertEqual(main.normalizar_llave("P", "11A"), "P-11A")


# ── 2. limpiar_monto ─────────────────────────────────────────────────────────

class TestLimpiarMonto(unittest.TestCase):

    def test_float_directo(self):
        self.assertAlmostEqual(main.limpiar_monto(38.0), 38.0)

    def test_string_numerico(self):
        self.assertAlmostEqual(main.limpiar_monto("40.00"), 40.0)

    def test_coma_como_decimal(self):
        self.assertAlmostEqual(main.limpiar_monto("38,50"), 38.5)

    def test_invalido_retorna_cero(self):
        self.assertEqual(main.limpiar_monto("abc"), 0.0)

    def test_none_retorna_cero(self):
        self.assertEqual(main.limpiar_monto(None), 0.0)

    def test_entero_se_convierte(self):
        self.assertAlmostEqual(main.limpiar_monto(20), 20.0)

    def test_redondea_dos_decimales(self):
        self.assertAlmostEqual(main.limpiar_monto("38.256"), 38.26, places=2)


# ── 3. es_ok ─────────────────────────────────────────────────────────────────

class TestEsOk(unittest.TestCase):

    def test_si_con_tilde(self):
        self.assertTrue(main.es_ok("Sí"))

    def test_si_sin_tilde(self):
        self.assertTrue(main.es_ok("si"))

    def test_si_mayusculas(self):
        self.assertTrue(main.es_ok("SI"))

    def test_s_solo(self):
        self.assertTrue(main.es_ok("s"))

    def test_yes_ingles(self):
        self.assertTrue(main.es_ok("yes"))

    def test_uno_string(self):
        self.assertTrue(main.es_ok("1"))

    def test_vacio_falso(self):
        self.assertFalse(main.es_ok(""))

    def test_no_falso(self):
        self.assertFalse(main.es_ok("No"))

    def test_guion_falso(self):
        self.assertFalse(main.es_ok("—"))

    def test_espacios_no_cuentan(self):
        self.assertTrue(main.es_ok("  Sí  "))


# ── 4. comparar_fresco ───────────────────────────────────────────────────────

class TestCompararFresco(unittest.TestCase):

    def test_coincidencia_exacta_confirma(self):
        """Mismo lote, mismo monto → confirmado, sin pendientes."""
        conf, pend = main.comparar_fresco([_reg("B", "12", 16.0)],
                                          [_reg("B", "12", 16.0)])
        self.assertEqual(len(conf), 1)
        self.assertEqual(len(pend), 0)
        self.assertAlmostEqual(conf[0]["monto"], 16.0)
        self.assertEqual(conf[0]["ciclo"], 1)

    def test_coincidencia_tolerancia_centavo(self):
        """Diferencia < 0.01 se trata como coincidencia."""
        conf, pend = main.comparar_fresco([_reg("A", "1", 38.00)],
                                          [_reg("A", "1", 38.009)])
        self.assertEqual(len(conf), 1)
        self.assertEqual(len(pend), 0)

    def test_diff_monto_tipo_correcto(self):
        """Montos distintos → tipo diff_monto."""
        _, pend = main.comparar_fresco([_reg("A", "8C", 38.0)],
                                       [_reg("A", "8C", 40.0)])
        self.assertEqual(len(pend), 1)
        self.assertEqual(pend[0]["tipo"], "diff_monto")

    def test_diff_monto_prefill_es_amiga(self):
        """MONTO_FINAL de diff_monto se pre-llena con valor de amiga."""
        _, pend = main.comparar_fresco([_reg("A", "8C", 38.0)],
                                       [_reg("A", "8C", 40.0)])
        self.assertAlmostEqual(pend[0]["monto_final"], 40.0)

    def test_diff_monto_monto_mia_y_amiga_llenos(self):
        """En diff_monto ambos montos están presentes."""
        _, pend = main.comparar_fresco([_reg("A", "8C", 38.0)],
                                       [_reg("A", "8C", 40.0)])
        self.assertAlmostEqual(float(pend[0]["monto_mia"]),   38.0)
        self.assertAlmostEqual(float(pend[0]["monto_amiga"]), 40.0)

    def test_solo_mia_tipo_correcto(self):
        """Lote solo en mia → tipo solo_mia."""
        _, pend = main.comparar_fresco([_reg("H", "3", 18.0)], [])
        self.assertEqual(pend[0]["tipo"], "solo_mia")

    def test_solo_mia_monto_amiga_vacio(self):
        """En solo_mia, MONTO_AMIGA es vacío."""
        _, pend = main.comparar_fresco([_reg("H", "3", 18.0)], [])
        self.assertEqual(pend[0]["monto_amiga"], "")

    def test_solo_mia_prefill_es_mia(self):
        """MONTO_FINAL de solo_mia se pre-llena con el valor de mia."""
        _, pend = main.comparar_fresco([_reg("H", "3", 18.0)], [])
        self.assertAlmostEqual(pend[0]["monto_final"], 18.0)

    def test_solo_amiga_tipo_correcto(self):
        """Lote solo en amiga → tipo solo_amiga."""
        _, pend = main.comparar_fresco([], [_reg("Z", "7", 25.0)])
        self.assertEqual(pend[0]["tipo"], "solo_amiga")

    def test_solo_amiga_monto_mia_vacio(self):
        """En solo_amiga, MONTO_MIA es vacío."""
        _, pend = main.comparar_fresco([], [_reg("Z", "7", 25.0)])
        self.assertEqual(pend[0]["monto_mia"], "")

    def test_mixto_tres_casos_a_la_vez(self):
        """B-12 coincide, A-8C diff_monto, H-3 solo_mia, Z-7 solo_amiga."""
        mia = [_reg("B", "12", 16.0), _reg("A", "8C", 38.0), _reg("H", "3", 18.0)]
        amiga = [_reg("B", "12", 16.0), _reg("A", "8C", 40.0), _reg("Z", "7", 25.0)]
        conf, pend = main.comparar_fresco(mia, amiga)
        self.assertEqual(len(conf), 1)
        self.assertEqual(len(pend), 3)
        tipos = {p["tipo"] for p in pend}
        self.assertEqual(tipos, {"diff_monto", "solo_mia", "solo_amiga"})

    def test_registros_vacios_sin_resultado(self):
        """Ambos vacíos → listas vacías."""
        conf, pend = main.comparar_fresco([], [])
        self.assertEqual(conf, [])
        self.assertEqual(pend, [])

    def test_llave_en_confirmados_normalizada(self):
        """LLAVE en confirmados está en mayúsculas."""
        conf, _ = main.comparar_fresco([_reg("a", "8c", 38.0)],
                                       [_reg("A", "8C", 38.0)])
        self.assertEqual(conf[0]["llave"], "A-8C")


# ── 5. cargar_registro ───────────────────────────────────────────────────────

class TestCargarRegistro(_TmpWbMixin, unittest.TestCase):

    def test_fila_valida_cargada(self):
        p = self._tmp_wb(_registro_wb([["A", "8C", 38.0, "03/05/2026"]]))
        regs = main.cargar_registro(p, "test")
        self.assertEqual(len(regs), 1)
        self.assertEqual(regs[0]["mz"],   "A")
        self.assertEqual(regs[0]["lt"],   "8C")
        self.assertAlmostEqual(regs[0]["monto"], 38.0)

    def test_llave_calculada(self):
        p = self._tmp_wb(_registro_wb([["P", "11A", 20.0, "05/05/2026"]]))
        regs = main.cargar_registro(p, "test")
        self.assertEqual(regs[0]["llave"], "P-11A")

    def test_mz_convertido_a_mayusculas(self):
        p = self._tmp_wb(_registro_wb([["a", "8c", 38.0, "03/05/2026"]]))
        regs = main.cargar_registro(p, "test")
        self.assertEqual(regs[0]["mz"], "A")
        self.assertEqual(regs[0]["lt"], "8C")

    def test_monto_cero_ignorado(self):
        p = self._tmp_wb(_registro_wb([["A", "1", 0.0, "03/05/2026"]]))
        self.assertEqual(main.cargar_registro(p, "test"), [])

    def test_mz_vacio_ignorado(self):
        p = self._tmp_wb(_registro_wb([["", "1", 38.0, "03/05/2026"]]))
        self.assertEqual(main.cargar_registro(p, "test"), [])

    def test_lt_vacio_ignorado(self):
        p = self._tmp_wb(_registro_wb([["A", "", 38.0, "03/05/2026"]]))
        self.assertEqual(main.cargar_registro(p, "test"), [])

    def test_multiples_filas(self):
        p = self._tmp_wb(_registro_wb([
            ["A", "8C", 38.0, "03/05/2026"],
            ["B", "12", 16.0, "03/05/2026"],
            ["Z", "3",  22.0, "04/05/2026"],
        ]))
        self.assertEqual(len(main.cargar_registro(p, "test")), 3)

    def test_archivo_faltante_lanza_error(self):
        with self.assertRaises(FileNotFoundError):
            main.cargar_registro(Path("/tmp/no_existe_xyz_1234.xlsx"), "test")


# ── 6. cargar_discrepancias ──────────────────────────────────────────────────

class TestCargarDiscrepancias(_TmpWbMixin, unittest.TestCase):

    def test_sin_archivo_lista_vacia(self):
        with patch("main.DISC_FILE", Path("/tmp/no_existe_disc_1234.xlsx")):
            self.assertEqual(main.cargar_discrepancias(), [])

    def test_fila_valida_cargada(self):
        p = self._tmp_wb(_discrepancias_wb([_fila_disc()]))
        with patch("main.DISC_FILE", p):
            result = main.cargar_discrepancias()
        self.assertEqual(len(result), 1)
        self.assertEqual(result[0]["tipo"],  "diff_monto")
        self.assertEqual(result[0]["mz"],    "A")
        self.assertEqual(result[0]["llave"], "A-8C")
        self.assertAlmostEqual(result[0]["monto_final"], 40.0)

    def test_ok_vacio_no_es_ok(self):
        p = self._tmp_wb(_discrepancias_wb([_fila_disc(ok="")]))
        with patch("main.DISC_FILE", p):
            result = main.cargar_discrepancias()
        self.assertFalse(main.es_ok(result[0]["ok"]))

    def test_ok_si_es_ok(self):
        p = self._tmp_wb(_discrepancias_wb([_fila_disc(ok="Sí")]))
        with patch("main.DISC_FILE", p):
            result = main.cargar_discrepancias()
        self.assertTrue(main.es_ok(result[0]["ok"]))

    def test_ciclo_leido(self):
        p = self._tmp_wb(_discrepancias_wb([_fila_disc(ciclo=2)]))
        with patch("main.DISC_FILE", p):
            result = main.cargar_discrepancias()
        self.assertEqual(result[0]["ciclo"], 2)

    def test_multiples_filas(self):
        p = self._tmp_wb(_discrepancias_wb([
            _fila_disc("diff_monto", "A", "8C"),
            _fila_disc("solo_mia",   "H", "3",  monto_mia=18.0, monto_amiga="", monto_final=18.0),
            _fila_disc("solo_amiga", "Z", "7",  monto_mia="",   monto_amiga=25.0, monto_final=25.0),
        ]))
        with patch("main.DISC_FILE", p):
            result = main.cargar_discrepancias()
        self.assertEqual(len(result), 3)
        tipos = {r["tipo"] for r in result}
        self.assertEqual(tipos, {"diff_monto", "solo_mia", "solo_amiga"})


# ── 7. cargar_pagos_existentes ───────────────────────────────────────────────

class TestCargarPagosExistentes(_TmpWbMixin, unittest.TestCase):

    def test_sin_archivo_lista_vacia(self):
        with patch("main.OUTPUT_FILE", Path("/tmp/no_existe_pagos_1234.xlsx")):
            self.assertEqual(main.cargar_pagos_existentes(), [])

    def test_fila_valida_cargada(self):
        p = self._tmp_wb(_pagos_efectivo_wb([_fila_pago()]))
        with patch("main.OUTPUT_FILE", p):
            result = main.cargar_pagos_existentes()
        self.assertEqual(len(result), 1)
        self.assertEqual(result[0]["mz"],    "A")
        self.assertEqual(result[0]["llave"], "A-8C")
        self.assertAlmostEqual(result[0]["monto"], 40.0)

    def test_ciclo_preservado(self):
        p = self._tmp_wb(_pagos_efectivo_wb([_fila_pago(ciclo=2)]))
        with patch("main.OUTPUT_FILE", p):
            result = main.cargar_pagos_existentes()
        self.assertEqual(result[0]["ciclo"], 2)

    def test_multiples_filas_cargadas(self):
        p = self._tmp_wb(_pagos_efectivo_wb([
            _fila_pago("A", "8C", 40.0, ciclo=1),
            _fila_pago("B", "12", 16.0, ciclo=1),
            _fila_pago("H", "3",  18.0, ciclo=2),
        ]))
        with patch("main.OUTPUT_FILE", p):
            result = main.cargar_pagos_existentes()
        self.assertEqual(len(result), 3)


if __name__ == "__main__":
    unittest.main(verbosity=2)
