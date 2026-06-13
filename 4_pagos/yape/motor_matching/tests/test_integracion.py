"""
Tests de integración — usan los archivos reales del proyecto.
Se saltan automáticamente si el archivo requerido no existe.

Cobertura:
  · cargar_maestro      — índice y ambiguos cargados correctamente
  · cargar_planilla     — usuarios y manzanas válidas
  · buscar_uid          — lookup real en usuarios_id.xlsx
  · leer_correcciones   — Ciclo 1 sin trazabilidad previa → vacío
  · matching completo   — run real en carpeta temporal, verifica invariantes
  · trazabilidad Ciclo1 → Ciclo1 re-run — sin duplicados ni pérdida
"""
import shutil
import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

sys.path.insert(0, str(Path(__file__).parent.parent))
import main
from exportar_motor import exportar_trazabilidad

BASE     = Path(__file__).parent.parent
SHARED   = BASE.parent.parent.parent / "shared"
MAESTRO_PATH   = BASE.parent / "construir_maestro" / "crear_maestro" / "outputs" / "maestro_yape.xlsx"
PLANILLA_PATH  = SHARED / "planilla_mes"
REPORTE_PATH   = SHARED / "reporte_mes_crudo"
USUARIOS_PATH  = SHARED / "usuarios_id.xlsx"
OUTPUT_PATH    = BASE / "outputs" / "pagos_yape_tepago.xlsx"
TRAZ_DIR       = BASE / "trazabilidad"


def _planilla_existe():
    return any(PLANILLA_PATH.glob("*.xlsx")) if PLANILLA_PATH.exists() else False

def _reporte_existe():
    return any(REPORTE_PATH.glob("*.xlsx")) if REPORTE_PATH.exists() else False


# ── Carga de maestro ──────────────────────────────────────────────────────────

@unittest.skipUnless(MAESTRO_PATH.exists(), "maestro_yape.xlsx no disponible")
class TestCargaMaestroReal(unittest.TestCase):

    @classmethod
    def setUpClass(cls):
        cls.indice, cls.indice_ambiguo = main.cargar_maestro()

    def test_indice_no_vacio(self):
        self.assertGreater(len(self.indice), 0, "El índice de maestro no debe estar vacío")

    def test_indice_ambiguo_no_vacio(self):
        self.assertGreater(len(self.indice_ambiguo), 0, "Debe haber orígenes ambiguos en el maestro")

    def test_estructura_entrada_indice(self):
        """Cada entrada del índice tiene mz y lote."""
        for origen, v in list(self.indice.items())[:5]:
            self.assertIn("mz",   v, f"Falta mz en {origen}")
            self.assertIn("lote", v, f"Falta lote en {origen}")

    def test_estructura_entrada_ambiguo(self):
        """Cada entrada ambigua tiene lista de candidatos con mz y lote."""
        for origen, candidatos in list(self.indice_ambiguo.items())[:5]:
            self.assertIsInstance(candidatos, list)
            self.assertGreater(len(candidatos), 1, f"{origen} debería tener 2+ candidatos")
            for c in candidatos:
                self.assertIn("mz",   c)
                self.assertIn("lote", c)

    def test_claves_uppercase(self):
        """Todas las claves del índice están en mayúsculas."""
        for k in list(self.indice.keys())[:20]:
            self.assertEqual(k, k.upper(), f"Clave no está en upper: {k!r}")


# ── Carga de planilla ─────────────────────────────────────────────────────────

@unittest.skipUnless(_planilla_existe(), "Carpeta planilla_mes sin archivos .xlsx")
class TestCargaPlanillaReal(unittest.TestCase):

    @classmethod
    def setUpClass(cls):
        cls.planilla, cls.mzs = main.cargar_planilla()

    def test_planilla_no_vacia(self):
        self.assertGreater(len(self.planilla), 0)

    def test_mzs_validas_no_vacias(self):
        self.assertGreater(len(self.mzs), 0)

    def test_estructura_entrada_planilla(self):
        """Cada lote tiene deuda_total y nombre."""
        for (mz, lote), v in list(self.planilla.items())[:10]:
            self.assertIn("deuda_total",  v, f"{mz}-{lote} sin deuda_total")
            self.assertIn("nombre",       v, f"{mz}-{lote} sin nombre")

    def test_deuda_total_numerica(self):
        """deuda_total es float en todos los registros."""
        for (mz, lote), v in self.planilla.items():
            self.assertIsInstance(v["deuda_total"], (int, float),
                                  f"{mz}-{lote} tiene deuda no numérica: {v['deuda_total']!r}")

    def test_mzs_son_strings_no_vacios(self):
        for mz in self.mzs:
            self.assertIsInstance(mz, str)
            self.assertGreater(len(mz), 0)


# ── buscar_uid ────────────────────────────────────────────────────────────────

@unittest.skipUnless(USUARIOS_PATH.exists(), "usuarios_id.xlsx no disponible")
class TestBuscarUidReal(unittest.TestCase):

    def test_retorna_tupla_dos_elementos(self):
        """buscar_uid devuelve siempre (str, str) aunque no encuentre."""
        uid, nom = main.buscar_uid("A", "1")
        self.assertIsInstance(uid, str)
        self.assertIsInstance(nom, str)

    def test_mz_lote_invalido_retorna_vacios(self):
        """MZ-LOTE inexistente → ('', '')."""
        uid, nom = main.buscar_uid("ZZZ", "9999")
        self.assertEqual(uid, "")
        self.assertEqual(nom, "")

    @unittest.skipUnless(_planilla_existe(), "planilla no disponible para obtener MZ-LOTE válido")
    def test_mz_lote_valido_retorna_uid_no_vacio(self):
        """Un MZ-LOTE que existe en planilla debe devolver UID y nombre."""
        planilla, _ = main.cargar_planilla()
        if not planilla:
            self.skipTest("Planilla vacía")
        mz, lote = next(iter(planilla))
        uid, nom = main.buscar_uid(mz, lote)
        self.assertGreater(len(uid),  0, f"UID vacío para {mz}-{lote}")
        self.assertGreater(len(nom),  0, f"Nombre vacío para {mz}-{lote}")


# ── leer_correcciones con Ciclo 1 limpio ─────────────────────────────────────

class TestLeerCorreccionesLimpio(unittest.TestCase):

    def test_ciclo1_sin_archivos_todo_vacio(self):
        """Sin trazabilidad ni pendientes → cinco valores vacíos."""
        with tempfile.TemporaryDirectory() as tmpdir:
            with patch("main.BASE_DIR",        Path(tmpdir)), \
                 patch("main.CORRECCIONES_DIR", Path(tmpdir) / "correcciones"):
                cs, ca, cm, va, vm, cp = main.leer_correcciones({})
        self.assertEqual(cs, {})
        self.assertEqual(ca, {})
        self.assertEqual(cm, {})
        self.assertEqual(va, [])
        self.assertEqual(vm, [])

    @unittest.skipUnless(_planilla_existe(), "planilla no disponible")
    def test_ciclo1_con_planilla_real_todo_vacio(self):
        """Sin pendientes, aunque planilla exista → correcciones vacías."""
        planilla, _ = main.cargar_planilla()
        with tempfile.TemporaryDirectory() as tmpdir:
            with patch("main.BASE_DIR",        Path(tmpdir)), \
                 patch("main.CORRECCIONES_DIR", Path(tmpdir) / "correcciones"):
                cs, ca, cm, va, vm, cp = main.leer_correcciones(planilla)
        self.assertEqual(cs, {})
        self.assertEqual(va, [])
        self.assertEqual(vm, [])


# ── Matching completo en carpeta temporal ────────────────────────────────────

@unittest.skipUnless(
    MAESTRO_PATH.exists() and _planilla_existe() and _reporte_existe(),
    "Requiere maestro + planilla + reporte para matching completo"
)
class TestMatchingCompletoReal(unittest.TestCase):
    """
    Corre el matching real en una carpeta temporal aislada.
    Verifica invariantes del resultado sin pisar archivos de producción.
    """

    @classmethod
    def setUpClass(cls):
        cls.tmpdir = tempfile.mkdtemp()
        tmp = Path(cls.tmpdir)
        out_dir  = tmp / "outputs"
        corr_dir = tmp / "correcciones"
        traz_dir = tmp / "trazabilidad"
        out_dir.mkdir(); corr_dir.mkdir(); traz_dir.mkdir()

        with patch("main.OUTPUT_DIR",        out_dir), \
             patch("main.CORRECCIONES_DIR",  corr_dir), \
             patch("main.BASE_DIR",          tmp):
            planilla, mzs   = main.cargar_planilla()
            indice, ind_amb = main.cargar_maestro()
            ancla           = main.obtener_ancla()
            df, df_pagaste, mapa = main.cargar_reportes(ancla)
            cls.todos       = main.ejecutar_matching(
                df, mapa, indice, planilla, {}, {}, {}, mzs, 1, ind_amb
            )
            cls.planilla    = planilla

    @classmethod
    def tearDownClass(cls):
        shutil.rmtree(cls.tmpdir, ignore_errors=True)

    def test_resultado_no_vacio(self):
        self.assertGreater(len(self.todos), 0, "El matching no produjo ningún resultado")

    def test_todos_tienen_estado(self):
        """Todos los registros tienen campo estado."""
        for r in self.todos:
            self.assertIn("estado", r, f"Registro sin estado: {r.get('origen')}")

    def test_identificados_tienen_mz_y_lote(self):
        """Registros identificados tienen MZ y LOTE no vacíos."""
        for r in self.todos:
            if r.get("estado") == "identificado":
                self.assertTrue(r.get("mz"), f"MZ vacío: {r.get('origen')}")
                self.assertTrue(r.get("lote"), f"LOTE vacío: {r.get('origen')}")

    def test_porcentaje_identificacion_razonable(self):
        """Al menos 70% de pagos identificados."""
        identificados = sum(1 for r in self.todos if r.get("estado") == "identificado")
        pct = identificados / len(self.todos) * 100
        self.assertGreaterEqual(pct, 70.0, f"Sólo {pct:.1f}% identificados — revisar maestro")

    def test_montos_positivos(self):
        """Todos los montos son positivos."""
        for r in self.todos:
            monto = r.get("monto_pago", r.get("monto", 0))
            if monto:
                self.assertGreater(float(monto), 0, f"Monto no positivo: {r.get('origen')}")

    def test_mz_en_planilla(self):
        """Los MZ de identificados existen en planilla."""
        for r in self.todos:
            if r.get("estado") == "identificado" and r.get("mz") and r.get("mz") != "BLANCO":
                mz, lote = r["mz"], r.get("lote", "")
                self.assertIn((mz, lote), self.planilla,
                              f"{mz}-{lote} identificado pero no está en planilla")


# ── Trazabilidad Ciclo 1 → re-run sin pérdida ni duplicación ─────────────────

class TestTrazabilidadCiclo1ReRun(unittest.TestCase):
    """
    Simula: corre matching → escribe trazabilidad → vuelve a correr
    (Ciclo 1 re-run tras borrar output). Verifica que trazabilidad no
    duplica filas y que las correcciones se recuperan íntegras.
    """

    def _ciclo(self, tmpdir, corr_simples_in, corr_multiples_in):
        from datetime import datetime
        tmp      = Path(tmpdir)
        traz_dir = tmp / "trazabilidad"
        traz_dir.mkdir(exist_ok=True)
        corr_dir = tmp / "correcciones"
        corr_dir.mkdir(exist_ok=True)
        ruta = traz_dir / f"trazabilidad_{datetime.today().strftime('%Y_%m')}.xlsx"

        exportar_trazabilidad(str(ruta), corr_simples_in, [], corr_multiples_in, 1, "24/05/2026 10:00")

        with patch("main.BASE_DIR",        tmp), \
             patch("main.CORRECCIONES_DIR", corr_dir):
            cs, _, cm, _, _, _ = main.leer_correcciones({})
        return cs, cm

    def test_dos_corr_simples_sobreviven_rerun(self):
        """2 correcciones simples → escribir → leer → mismas 2."""
        corr = {
            "PEDRO P*|02/05/2026 23:02:11": {"mz": "A", "lote": "1"},
            "JUANA M*|04/05/2026 08:33:00": {"mz": "Q", "lote": "12"},
        }
        with tempfile.TemporaryDirectory() as tmpdir:
            cs, _ = self._ciclo(tmpdir, corr, {})
        self.assertEqual(len(cs), 2)
        self.assertIn("PEDRO P*|02/05/2026 23:02:11", cs)
        self.assertIn("JUANA M*|04/05/2026 08:33:00", cs)

    def test_rerun_no_duplica(self):
        """Re-run (leer → escribir → leer): sigue siendo 2, no 4."""
        corr = {
            "ANA R*|10/05/2026 09:15:00": {"mz": "C", "lote": "5"},
            "LUIS V*|12/05/2026 14:22:30": {"mz": "D", "lote": "8"},
        }
        with tempfile.TemporaryDirectory() as tmpdir:
            cs1, _ = self._ciclo(tmpdir, corr,  {})   # primera pasada
            cs2, _ = self._ciclo(tmpdir, cs1,   {})   # re-run con lo leído
        self.assertEqual(len(cs2), 2, "Re-run duplicó filas en trazabilidad")

    def test_mz_lote_preservados_tras_rerun(self):
        """Los valores MZ y LOTE se conservan exactos tras round-trip."""
        corr = {"ROSA M*|10/05/2026 11:00:00": {"mz": "E", "lote": "2"}}
        with tempfile.TemporaryDirectory() as tmpdir:
            cs, _ = self._ciclo(tmpdir, corr, {})
        v = cs.get("ROSA M*|10/05/2026 11:00:00", {})
        self.assertEqual(v.get("mz"),   "E")
        self.assertEqual(v.get("lote"), "2")


if __name__ == "__main__":
    unittest.main(verbosity=2)
