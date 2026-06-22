"""Tests para sin_servicio/actualizar_lista.py.

Cubre:
    _detectar_origen   — auto / default_sin_agua / supervisor por prefijo de NOTAS
    _validar_decisiones — detecta vacías e inválidas
    _planificar        — cada combinación (DECISIÓN, estado en lista)
    _aplicar_a_lista   — INSERT, UPDATE, idempotencia, no muta input
"""
import pytest

from sin_servicio import config
from sin_servicio.actualizar_lista import (
    _aplicar_a_lista,
    _detectar_origen,
    _planificar,
    _validar_decisiones,
)


def _fila(mz, lt, decision, notas="", nombre="X"):
    """Helper que construye una fila del reporte con defaults razonables."""
    return {
        "MZ": mz, "LT": lt, "NOMBRE": nombre,
        "ESCENARIO": "SIN_LECTURA",
        "ÚLTIMO_MES_CON_LECTURA": "", "MARCACION_ÚLTIMA": "",
        "MESES_SIN_LECTURA": "1", "ESTADO_LISTA": "NO_ESTÁ",
        "DECISIÓN": decision, "NOTAS": notas,
        "CICLO": "1", "FECHA_DETECCIÓN": "2026-07-01 00:00",
    }


# ── _detectar_origen ─────────────────────────────────────────────────────────

class TestDetectarOrigen:
    def test_prefijo_auto(self):
        assert _detectar_origen(f"{config.PREFIJO_AUTO} · 2026-06-21") == "auto"

    def test_prefijo_default_sin_agua(self):
        assert _detectar_origen(
            f"{config.PREFIJO_DEFAULT_SIN_AGUA} · 2026-06-21"
        ) == "default_sin_agua"

    def test_otras_notas_son_supervisor(self):
        assert _detectar_origen("supervisor confirmó en campo") == "supervisor"

    def test_notas_vacias_supervisor(self):
        assert _detectar_origen("") == "supervisor"


# ── _validar_decisiones ──────────────────────────────────────────────────────

class TestValidarDecisiones:
    def test_todas_llenas_devuelve_lista_vacia(self):
        filas = [_fila("A", "1", "sin_servicio")]
        assert _validar_decisiones(filas) == []

    def test_detecta_vacia(self):
        filas = [_fila("A", "1", "")]
        invalidas = _validar_decisiones(filas)
        assert len(invalidas) == 1
        assert invalidas[0]["_motivo"] == "vacía"

    def test_detecta_valor_invalido(self):
        filas = [_fila("A", "1", "no_existe_esta_decision")]
        invalidas = _validar_decisiones(filas)
        assert len(invalidas) == 1
        assert "valor inválido" in invalidas[0]["_motivo"]


# ── _planificar ──────────────────────────────────────────────────────────────

class TestPlanificar:
    def test_sin_servicio_sobre_no_esta_inserta_sin_medidor(self):
        plan = _planificar([_fila("A", "1", "sin_servicio")], {})
        assert plan[0]["accion"] == "INSERT"
        assert plan[0]["tipo_nuevo"] == "SIN_MEDIDOR"

    def test_sin_servicio_sobre_sin_medidor_skip(self):
        lista = {("A", "1"): {"tipo": "SIN_MEDIDOR", "nombre": "X"}}
        plan = _planificar([_fila("A", "1", "sin_servicio")], lista)
        assert plan[0]["accion"] == "SKIP_DUPLICATE"

    def test_sin_servicio_sobre_sin_agua_update(self):
        lista = {("A", "1"): {"tipo": "SIN_AGUA", "nombre": "X"}}
        plan = _planificar([_fila("A", "1", "sin_servicio")], lista)
        assert plan[0]["accion"] == "UPDATE"
        assert plan[0]["tipo_nuevo"] == "SIN_MEDIDOR"

    def test_sin_servicio_sobre_en_investigacion_update(self):
        lista = {("A", "1"): {"tipo": "EN_INVESTIGACIÓN", "nombre": "X"}}
        plan = _planificar([_fila("A", "1", "sin_servicio")], lista)
        assert plan[0]["accion"] == "UPDATE"
        assert plan[0]["tipo_nuevo"] == "SIN_MEDIDOR"

    def test_investigar_sobre_no_esta_inserta(self):
        plan = _planificar([_fila("A", "1", "investigar")], {})
        assert plan[0]["accion"] == "INSERT"
        assert plan[0]["tipo_nuevo"] == "EN_INVESTIGACIÓN"

    def test_investigar_sobre_presente_skip(self):
        lista = {("A", "1"): {"tipo": "SIN_AGUA", "nombre": "X"}}
        plan = _planificar([_fila("A", "1", "investigar")], lista)
        assert plan[0]["accion"] == "SKIP_DUPLICATE"

    def test_error_captura_no_op(self):
        plan = _planificar([_fila("A", "1", "error_captura")], {})
        assert plan[0]["accion"] == "NO_OP_ERROR_CAPTURA"
        assert plan[0]["tipo_nuevo"] == ""

    def test_ignorar_no_op(self):
        plan = _planificar([_fila("A", "1", "ignorar")], {})
        assert plan[0]["accion"] == "NO_OP_IGNORAR"

    def test_origen_auto_se_detecta_en_plan(self):
        notas = f"{config.PREFIJO_AUTO} · 2026-06-21"
        plan = _planificar([_fila("A", "1", "sin_servicio", notas)], {})
        assert plan[0]["origen"] == "auto"

    def test_origen_default_sin_agua_se_detecta(self):
        notas = f"{config.PREFIJO_DEFAULT_SIN_AGUA} · 2026-06-21"
        plan = _planificar([_fila("A", "1", "sin_servicio", notas)], {})
        assert plan[0]["origen"] == "default_sin_agua"


# ── _aplicar_a_lista ─────────────────────────────────────────────────────────

def _plan_insert(mz, lt, tipo, notas=""):
    return {
        "accion": "INSERT", "mz": mz, "lt": lt, "nombre": "X",
        "tipo_nuevo": tipo, "notas": notas,
        "decision": "sin_servicio", "origen": "supervisor", "ciclo": "1",
    }


def _plan_update(mz, lt, tipo, notas=""):
    return {
        "accion": "UPDATE", "mz": mz, "lt": lt, "nombre": "X",
        "tipo_nuevo": tipo, "notas": notas,
        "decision": "sin_servicio", "origen": "supervisor", "ciclo": "1",
    }


class TestAplicarALista:
    def test_insert_agrega_entrada(self):
        nueva = _aplicar_a_lista(
            [_plan_insert("A", "1", "SIN_MEDIDOR", "nota")],
            {}, "2026-07-01",
        )
        assert ("A", "1") in nueva
        assert nueva[("A", "1")]["tipo"] == "SIN_MEDIDOR"
        assert nueva[("A", "1")]["fecha_inicio"] == "2026-07-01"
        assert nueva[("A", "1")]["ultima_revision"] == "2026-07-01"
        assert nueva[("A", "1")]["meses"] == 99
        assert nueva[("A", "1")]["notas"] == "nota"

    def test_insert_en_investigacion_meses_1(self):
        nueva = _aplicar_a_lista(
            [_plan_insert("A", "1", "EN_INVESTIGACIÓN")],
            {}, "2026-07-01",
        )
        assert nueva[("A", "1")]["meses"] == 1

    def test_update_promociona_y_preserva_fecha_inicio(self):
        lista = {("A", "1"): {
            "nombre": "X", "tipo": "SIN_AGUA",
            "fecha_inicio": "2026-03-01", "ultima_revision": "2026-03-01",
            "meses": 2, "notas": "vieja",
        }}
        nueva = _aplicar_a_lista(
            [_plan_update("A", "1", "SIN_MEDIDOR", "nueva")],
            lista, "2026-07-01",
        )
        assert nueva[("A", "1")]["tipo"] == "SIN_MEDIDOR"
        assert nueva[("A", "1")]["fecha_inicio"] == "2026-03-01"  # no se toca
        assert nueva[("A", "1")]["ultima_revision"] == "2026-07-01"
        # NOTAS concatenadas con separador.
        assert "vieja" in nueva[("A", "1")]["notas"]
        assert "nueva" in nueva[("A", "1")]["notas"]

    def test_skip_y_no_op_no_alteran_lista(self):
        lista = {("A", "1"): {
            "nombre": "X", "tipo": "SIN_MEDIDOR",
            "fecha_inicio": "2026-01-01", "ultima_revision": "2026-01-01",
            "meses": 99, "notas": "",
        }}
        plan = [
            {"accion": "SKIP_DUPLICATE", "mz": "A", "lt": "1", "nombre": "X",
             "tipo_nuevo": "SIN_MEDIDOR", "notas": "",
             "decision": "sin_servicio", "origen": "supervisor", "ciclo": "1"},
            {"accion": "NO_OP_ERROR_CAPTURA", "mz": "B", "lt": "2", "nombre": "Y",
             "tipo_nuevo": "", "notas": "",
             "decision": "error_captura", "origen": "supervisor", "ciclo": "1"},
        ]
        nueva = _aplicar_a_lista(plan, lista, "2026-07-01")
        assert nueva[("A", "1")]["ultima_revision"] == "2026-01-01"
        assert ("B", "2") not in nueva

    def test_no_muta_input_lista(self):
        lista = {}
        _aplicar_a_lista(
            [_plan_insert("A", "1", "SIN_MEDIDOR")], lista, "2026-07-01"
        )
        assert lista == {}
