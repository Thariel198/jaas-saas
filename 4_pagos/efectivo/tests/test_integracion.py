"""
Tests de integración para 4_pagos/efectivo/main.py.
Corren el flujo completo usando directorios temporales aislados.

Cobertura:
  · Ciclo 1 sin discrepancias  — todo coincide → pagos_efectivo listo directo
  · Ciclo 1 con discrepancias  — tipos generados, MONTO_FINAL pre-llenado
  · Ciclo 1 todo discrepancias — pagos_efectivo no se crea si nada coincide
  · Ciclo 2 resolución parcial — resueltas se agregan, pendientes persisten
  · Ciclo 2 todo resuelto      — discrepancias eliminado, ciclos correctos
  · Ciclo 2 monto cero         — OK=Sí con monto 0 no va a pagos
  · Sin duplicados             — re-run no duplica llave ya confirmada
"""
import shutil
import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from openpyxl import Workbook, load_workbook

sys.path.insert(0, str(Path(__file__).parent.parent))
import main


# ── Helpers de escritura ──────────────────────────────────────────────────────

def _crear_registro(path: Path, filas: list):
    wb = Workbook()
    ws = wb.active
    ws.append(["MZ", "LT", "MONTO", "FECHA"])
    for f in filas:
        ws.append(f)
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)

def _crear_discrepancias(path: Path, filas: list):
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
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)

def _crear_pagos_efectivo(path: Path, filas: list):
    wb = Workbook()
    ws = wb.active
    ws.title = "pagos_efectivo"
    ws.append(["¿DÓNDE VIVE?", "", "", "", "¿CUÁNTO PAGÓ?", "", "¿CUÁNDO?", ""])
    ws.append(["MZ", "LT", "LLAVE", "", "MONTO", "", "FECHA", "CICLO"])
    for f in filas:
        ws.append(f)
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)

def _fila_disc(tipo, mz, lt, monto_mia, monto_amiga, monto_final,
               ok="", fecha="03/05/2026", ciclo=1) -> list:
    llave = main.normalizar_llave(mz, lt)
    return [tipo, mz, lt, llave, "", monto_mia, monto_amiga, "",
            monto_final, ok, "", fecha, ciclo]

def _fila_pago(mz, lt, monto, fecha="03/05/2026", ciclo=1) -> list:
    llave = main.normalizar_llave(mz, lt)
    return [mz, lt, llave, "", monto, "", fecha, ciclo]


# ── Helpers de lectura ────────────────────────────────────────────────────────

def _leer_pagos(path: Path) -> list:
    """Lee filas de datos de pagos_efectivo.xlsx (doble cabecera)."""
    wb = load_workbook(path, read_only=True, data_only=True)
    ws = wb.active
    datos = list(ws.values)
    wb.close()
    if len(datos) < 3:
        return []
    headers = [str(h).strip().upper() if h else "" for h in datos[1]]
    result = []
    for fila in datos[2:]:
        if not fila:
            continue
        row = {headers[i]: fila[i] for i in range(min(len(headers), len(fila)))}
        if row.get("MZ"):
            result.append(row)
    return result

def _leer_disc(path: Path) -> list:
    """Lee filas de datos de discrepancias.xlsx (doble cabecera)."""
    wb = load_workbook(path, read_only=True, data_only=True)
    ws = wb.active
    datos = list(ws.values)
    wb.close()
    if len(datos) < 3:
        return []
    headers = [str(h).strip().upper() if h else "" for h in datos[1]]
    result = []
    for fila in datos[2:]:
        if not fila:
            continue
        row = {headers[i]: fila[i] for i in range(min(len(headers), len(fila)))}
        if row.get("MZ"):
            result.append(row)
    return result


# ── Base ──────────────────────────────────────────────────────────────────────

class _Base(unittest.TestCase):
    """Crea directorios temporales y parchea los paths de main.py."""

    def setUp(self):
        self.tmpdir    = tempfile.mkdtemp()
        self.tmp       = Path(self.tmpdir)
        self.inp_dir   = self.tmp / "inputs"
        self.out_dir   = self.tmp / "outputs"
        self.inp_dir.mkdir()
        self.out_dir.mkdir()
        self.reg_mia   = self.inp_dir / "registro_mia.xlsx"
        self.reg_amiga = self.inp_dir / "registro_amiga.xlsx"
        self.disc_file = self.out_dir / "discrepancias.xlsx"
        self.out_file  = self.out_dir / "pagos_efectivo.xlsx"

    def tearDown(self):
        shutil.rmtree(self.tmpdir, ignore_errors=True)

    def _run(self):
        with patch("main.REGISTRO_MIA",   self.reg_mia),   \
             patch("main.REGISTRO_AMIGA", self.reg_amiga), \
             patch("main.DISC_FILE",      self.disc_file),  \
             patch("main.OUTPUT_FILE",    self.out_file),   \
             patch("main.OUTPUTS_DIR",    self.out_dir):
            main.main()

    def _pagos(self) -> list:
        return _leer_pagos(self.out_file)

    def _disc(self) -> list:
        return _leer_disc(self.disc_file)

    def _llaves_pagos(self) -> set:
        return {main.normalizar_llave(f["MZ"], f["LT"]) for f in self._pagos()}


# ── Ciclo 1: sin discrepancias ───────────────────────────────────────────────

class TestCiclo1SinDiscrepancias(_Base):
    """Los dos registros son idénticos → todo pasa directo."""

    def setUp(self):
        super().setUp()
        datos = [["A", "8C", 38.0, "03/05/2026"],
                 ["B", "12", 16.0, "03/05/2026"],
                 ["Z", "3",  22.0, "04/05/2026"]]
        _crear_registro(self.reg_mia,   datos)
        _crear_registro(self.reg_amiga, datos)

    def test_pagos_efectivo_creado(self):
        self._run()
        self.assertTrue(self.out_file.exists())

    def test_discrepancias_no_creado(self):
        self._run()
        self.assertFalse(self.disc_file.exists())

    def test_tres_cobros_confirmados(self):
        self._run()
        self.assertEqual(len(self._pagos()), 3)

    def test_ciclo_es_1(self):
        self._run()
        for f in self._pagos():
            self.assertEqual(int(f["CICLO"]), 1)

    def test_montos_correctos(self):
        self._run()
        montos = {main.normalizar_llave(f["MZ"], f["LT"]): float(f["MONTO"])
                  for f in self._pagos()}
        self.assertAlmostEqual(montos["A-8C"], 38.0)
        self.assertAlmostEqual(montos["B-12"], 16.0)
        self.assertAlmostEqual(montos["Z-3"],  22.0)

    def test_llave_escrita(self):
        self._run()
        for f in self._pagos():
            llave_esperada = main.normalizar_llave(f["MZ"], f["LT"])
            self.assertEqual(f["LLAVE"], llave_esperada)


# ── Ciclo 1: con discrepancias ───────────────────────────────────────────────

class TestCiclo1ConDiscrepancias(_Base):

    def setUp(self):
        super().setUp()
        _crear_registro(self.reg_mia, [
            ["B", "12", 16.0, "03/05/2026"],   # coincide
            ["A", "8C", 38.0, "03/05/2026"],   # diff_monto
            ["H", "3",  18.0, "05/05/2026"],   # solo_mia
        ])
        _crear_registro(self.reg_amiga, [
            ["B", "12", 16.0, "03/05/2026"],   # coincide
            ["A", "8C", 40.0, "03/05/2026"],   # diff_monto
            ["Z", "7",  25.0, "06/05/2026"],   # solo_amiga
        ])

    def test_discrepancias_creado(self):
        self._run()
        self.assertTrue(self.disc_file.exists())

    def test_solo_coincidente_en_pagos(self):
        """B-12 coincide → solo ese en pagos_efectivo."""
        self._run()
        self.assertEqual(self._llaves_pagos(), {"B-12"})

    def test_tres_discrepancias_generadas(self):
        self._run()
        self.assertEqual(len(self._disc()), 3)

    def test_tipos_correctos(self):
        self._run()
        tipos = {f["TIPO"] for f in self._disc()}
        self.assertEqual(tipos, {"diff_monto", "solo_mia", "solo_amiga"})

    def test_diff_monto_prefill_es_amiga(self):
        """MONTO_FINAL de diff_monto pre-llenado con valor de amiga (40.00)."""
        self._run()
        diff = next(f for f in self._disc() if f["TIPO"] == "diff_monto")
        self.assertAlmostEqual(float(diff["MONTO_FINAL"]), 40.0)

    def test_solo_mia_prefill_es_mia(self):
        """MONTO_FINAL de solo_mia pre-llenado con valor de mia (18.00)."""
        self._run()
        sm = next(f for f in self._disc() if f["TIPO"] == "solo_mia")
        self.assertAlmostEqual(float(sm["MONTO_FINAL"]), 18.0)

    def test_ok_vacio_en_todas_las_discrepancias(self):
        """OK queda vacío — el usuario lo completa."""
        self._run()
        for f in self._disc():
            self.assertFalse(main.es_ok(str(f.get("OK", ""))))

    def test_ciclo_1_en_discrepancias(self):
        self._run()
        for f in self._disc():
            self.assertEqual(int(f["CICLO"]), 1)


# ── Ciclo 1: todo discrepancias (ninguna coincide) ───────────────────────────

class TestCiclo1TodoDiscrepancias(_Base):

    def setUp(self):
        super().setUp()
        _crear_registro(self.reg_mia,   [["A", "1", 10.0, "01/05/2026"]])
        _crear_registro(self.reg_amiga, [["A", "1", 20.0, "01/05/2026"]])

    def test_pagos_efectivo_no_creado(self):
        """Si nada coincide en ciclo 1 → pagos_efectivo.xlsx no se crea."""
        self._run()
        self.assertFalse(self.out_file.exists())

    def test_discrepancias_creado(self):
        self._run()
        self.assertTrue(self.disc_file.exists())


# ── Ciclo 2: resolución parcial ──────────────────────────────────────────────

class TestCiclo2ResolucionParcial(_Base):
    """
    Estado inicial: pagos_efectivo tiene B-12 (ciclo 1).
    discrepancias tiene A-8C (OK=Sí) y H-3 (pendiente).
    """

    def setUp(self):
        super().setUp()
        _crear_registro(self.reg_mia,   [])
        _crear_registro(self.reg_amiga, [])
        _crear_pagos_efectivo(self.out_file, [
            _fila_pago("B", "12", 16.0, ciclo=1),
        ])
        _crear_discrepancias(self.disc_file, [
            _fila_disc("diff_monto", "A", "8C", 38.0, 40.0, 40.0, ok="Sí"),
            _fila_disc("solo_mia",   "H", "3",  18.0, "",   18.0, ok=""),
        ])

    def test_pagos_crece_con_resuelta(self):
        """A-8C resuelta (OK=Sí) se agrega a pagos_efectivo."""
        self._run()
        self.assertIn("A-8C", self._llaves_pagos())

    def test_pagos_conserva_ciclo1(self):
        """B-12 del ciclo 1 se conserva."""
        self._run()
        self.assertIn("B-12", self._llaves_pagos())

    def test_monto_final_respetado(self):
        """MONTO de A-8C es MONTO_FINAL=40 aprobado por el usuario."""
        self._run()
        a8c = next(f for f in self._pagos()
                   if main.normalizar_llave(f["MZ"], f["LT"]) == "A-8C")
        self.assertAlmostEqual(float(a8c["MONTO"]), 40.0)

    def test_discrepancias_persiste_con_pendiente(self):
        """H-3 sigue pendiente → discrepancias.xlsx no se borra."""
        self._run()
        self.assertTrue(self.disc_file.exists())

    def test_discrepancias_solo_tiene_pendiente(self):
        """Discrepancias actualizado tiene solo H-3."""
        self._run()
        filas = self._disc()
        self.assertEqual(len(filas), 1)
        self.assertEqual(filas[0]["MZ"], "H")


# ── Ciclo 2: todo resuelto ───────────────────────────────────────────────────

class TestCiclo2TodoResuelto(_Base):
    """
    Estado: pagos_efectivo tiene B-12.
    discrepancias tiene A-8C y H-3 ambas con OK=Sí.
    """

    def setUp(self):
        super().setUp()
        _crear_registro(self.reg_mia,   [])
        _crear_registro(self.reg_amiga, [])
        _crear_pagos_efectivo(self.out_file, [
            _fila_pago("B", "12", 16.0, ciclo=1),
        ])
        _crear_discrepancias(self.disc_file, [
            _fila_disc("diff_monto", "A", "8C", 38.0, 40.0, 40.0, ok="Sí"),
            _fila_disc("solo_mia",   "H", "3",  18.0, "",   18.0, ok="Sí"),
        ])

    def test_discrepancias_eliminado(self):
        """Todas OK=Sí → discrepancias.xlsx se borra."""
        self._run()
        self.assertFalse(self.disc_file.exists())

    def test_pagos_tiene_los_tres(self):
        """B-12 + A-8C + H-3 = 3 cobros confirmados."""
        self._run()
        self.assertEqual(len(self._pagos()), 3)

    def test_ciclo_1_preservado_en_b12(self):
        """B-12 conserva CICLO=1 (fue confirmado en ciclo 1)."""
        self._run()
        b12 = next(f for f in self._pagos()
                   if main.normalizar_llave(f["MZ"], f["LT"]) == "B-12")
        self.assertEqual(int(b12["CICLO"]), 1)

    def test_ciclo_2_en_nuevos_resueltos(self):
        """A-8C y H-3 resueltos en este run tienen CICLO=2."""
        self._run()
        por_llave = {main.normalizar_llave(f["MZ"], f["LT"]): int(f["CICLO"])
                     for f in self._pagos()}
        self.assertEqual(por_llave["A-8C"], 2)
        self.assertEqual(por_llave["H-3"],  2)

    def test_todos_los_montos_correctos(self):
        self._run()
        montos = {main.normalizar_llave(f["MZ"], f["LT"]): float(f["MONTO"])
                  for f in self._pagos()}
        self.assertAlmostEqual(montos["B-12"], 16.0)
        self.assertAlmostEqual(montos["A-8C"], 40.0)
        self.assertAlmostEqual(montos["H-3"],  18.0)


# ── Ciclo 2: MONTO_FINAL=0 no va a pagos ─────────────────────────────────────

class TestCiclo2MontoCero(_Base):
    """OK=Sí pero MONTO_FINAL=0 → el cobro no se incluye en pagos_efectivo."""

    def setUp(self):
        super().setUp()
        _crear_registro(self.reg_mia,   [])
        _crear_registro(self.reg_amiga, [])
        _crear_discrepancias(self.disc_file, [
            _fila_disc("solo_mia", "X", "99", 15.0, "", 0.0, ok="Sí"),
        ])

    def test_x99_no_va_a_pagos(self):
        """X-99 con monto 0 no aparece en pagos_efectivo."""
        self._run()
        if self.out_file.exists():
            self.assertNotIn("X-99", self._llaves_pagos())

    def test_discrepancias_eliminado_porque_ok(self):
        """Aunque monto=0, OK=Sí elimina discrepancias."""
        self._run()
        self.assertFalse(self.disc_file.exists())


# ── Invariante: sin duplicados ───────────────────────────────────────────────

class TestSinDuplicados(_Base):
    """
    Si una llave ya está en pagos_efectivo y vuelve a aparecer como resuelta
    en discrepancias, no debe duplicarse.
    """

    def setUp(self):
        super().setUp()
        _crear_registro(self.reg_mia,   [])
        _crear_registro(self.reg_amiga, [])
        _crear_pagos_efectivo(self.out_file, [
            _fila_pago("A", "8C", 40.0, ciclo=1),
        ])
        _crear_discrepancias(self.disc_file, [
            _fila_disc("diff_monto", "A", "8C", 38.0, 40.0, 40.0, ok="Sí"),
        ])

    def test_no_duplica_llave_existente(self):
        """A-8C ya confirmada → re-run no la duplica."""
        self._run()
        llaves = [main.normalizar_llave(f["MZ"], f["LT"]) for f in self._pagos()]
        self.assertEqual(llaves.count("A-8C"), 1,
                         f"A-8C aparece {llaves.count('A-8C')} veces (esperado 1)")


# ── Ciclo 2 con tres ciclos de corrección ────────────────────────────────────

class TestTresCiclosCorrecion(_Base):
    """
    Simula que el usuario resuelve de a uno: ciclo 1 → 2 → 3.
    Verifica que cada ciclo acumula sin pérdida.
    """

    def setUp(self):
        super().setUp()
        _crear_registro(self.reg_mia,   [])
        _crear_registro(self.reg_amiga, [])

    def test_acumulacion_por_ciclos(self):
        # Ciclo 2: solo A-8C resuelta, H-3 y Z-7 pendientes
        _crear_pagos_efectivo(self.out_file, [
            _fila_pago("B", "12", 16.0, ciclo=1),
        ])
        _crear_discrepancias(self.disc_file, [
            _fila_disc("diff_monto", "A", "8C", 38.0, 40.0, 40.0, ok="Sí", ciclo=1),
            _fila_disc("solo_mia",   "H", "3",  18.0, "",   18.0, ok="",  ciclo=1),
            _fila_disc("solo_amiga", "Z", "7",  "",   25.0, 25.0, ok="",  ciclo=1),
        ])
        self._run()
        self.assertEqual(len(self._pagos()), 2)   # B-12 + A-8C
        self.assertEqual(len(self._disc()),  2)   # H-3 + Z-7 pendientes

        # Ciclo 3: H-3 resuelta, Z-7 pendiente
        disc_actual = self._disc()
        for f in disc_actual:
            if f["MZ"] == "H":
                f["OK"] = "Sí"
        _crear_discrepancias(self.disc_file, [
            _fila_disc("solo_mia",   "H", "3",  18.0, "", 18.0, ok="Sí", ciclo=2),
            _fila_disc("solo_amiga", "Z", "7",  "",  25.0, 25.0, ok="",  ciclo=2),
        ])
        self._run()
        self.assertEqual(len(self._pagos()), 3)   # B-12 + A-8C + H-3
        self.assertEqual(len(self._disc()),  1)   # Z-7 pendiente

        # Ciclo 4: Z-7 resuelta → todo listo
        _crear_discrepancias(self.disc_file, [
            _fila_disc("solo_amiga", "Z", "7", "", 25.0, 25.0, ok="Sí", ciclo=3),
        ])
        self._run()
        self.assertEqual(len(self._pagos()), 4)
        self.assertFalse(self.disc_file.exists())


if __name__ == "__main__":
    unittest.main(verbosity=2)
