"""Constantes y paths del submódulo sin_servicio.

Modificar si la junta cambia el umbral de auto-promoción o si se renombra alguna ruta.
Los schemas COLS_LISTA y COLS_VALIDACION son el contrato visual de:
    docs/formato_lista_sin_servicio.html
    docs/formato_validacion_ausencias.html
Si cambian acá, deben cambiar allá.
"""
from pathlib import Path

# ── Umbrales ──────────────────────────────────────────────────────────────────
# Ciclos consecutivos sin lectura para auto-promover EN_INVESTIGACIÓN → SIN_MEDIDOR.
# validar_ausencias.py detecta el caso y pre-llena DECISIÓN=sin_servicio en el reporte;
# actualizar_lista.py la aplica como UPDATE (no INSERT) preservando writer único.
UMBRAL_AUTO_PROMOCION = 3

# Valor que se escribe en MESES_SIN_LECTURA cuando el usuario nunca tuvo lectura
# registrada en el acumulado (no aparece en ningún ciclo previo).
MESES_NUNCA_TUVO = 99

# ── Vocabularios cerrados ─────────────────────────────────────────────────────
# Si una celda contiene un valor fuera de estos sets, el lector debe fallar ruidoso.
TIPOS_VALIDOS      = ("SIN_MEDIDOR", "SIN_AGUA", "EN_INVESTIGACIÓN")
ESCENARIOS_VALIDOS = ("SIN_LECTURA", "USUARIO_FANTASMA")
# SIN_MEDIDOR nunca llega al reporte (queda justificado); los demás sí.
ESTADOS_LISTA      = ("NO_ESTÁ", "SIN_AGUA", "EN_INVESTIGACIÓN")
DECISIONES_VALIDAS = ("sin_servicio", "error_captura", "investigar", "ignorar")

# Prefijos que escribe validar_ausencias.py en NOTAS cuando pre-llena DECISIÓN.
# actualizar_lista.py los reconoce para registrar ORIGEN_PROPUESTA en el audit
# (auto vs default) y distinguir UPDATE (auto-promoción de EN_INVEST) de INSERT.
PREFIJO_AUTO            = "auto: 3 ciclos sin lectura"
PREFIJO_DEFAULT_SIN_AGUA = "default: SIN_AGUA → sin_servicio"

# ── Schemas (orden de columnas en los xlsx) ───────────────────────────────────
COLS_LISTA = (
    "MZ", "LT", "NOMBRE",
    "TIPO", "FECHA_INICIO", "ÚLTIMA_REVISIÓN",
    "MESES_SIN_LECTURA",
    "NOTAS",
)

COLS_VALIDACION = (
    "MZ", "LT", "NOMBRE",
    "ESCENARIO", "ÚLTIMO_MES_CON_LECTURA", "MARCACION_ÚLTIMA",
    "MESES_SIN_LECTURA", "ESTADO_LISTA",
    "DECISIÓN", "NOTAS",
    "CICLO", "FECHA_DETECCIÓN",
)

# Columnas que llena el supervisor — deben preservarse entre re-ejecuciones de
# validar_ausencias.py para no destruir trabajo del supervisor en pasadas anteriores.
COLS_SUPERVISOR = ("DECISIÓN", "NOTAS")

COLS_AUDIT = (
    "MZ", "LT", "NOMBRE",
    "ACCION", "TIPO_NUEVO",
    "DECISION_ORIGEN", "ORIGEN_PROPUESTA", "NOTAS",
    "FECHA", "CICLO_ORIGEN",
)

# ── Paths del submódulo ───────────────────────────────────────────────────────
ROOT = Path(__file__).parent

INPUTS_DIR  = ROOT / "inputs"
OUTPUTS_DIR = ROOT / "outputs"
DOCS_DIR    = ROOT / "docs"
TESTS_DIR   = ROOT / "tests"
BACKUPS_DIR = INPUTS_DIR / "backups"

LISTA_PATH       = INPUTS_DIR / "lista_sin_servicio.xlsx"
AUDIT_LISTA_PATH = OUTPUTS_DIR / "audit_lista.xlsx"

# Archivo plano original — solo lo lee migrar_lista.py una sola vez.
LISTA_LEGACY_PATH = ROOT.parent / "inputs" / "Lista de usuarios sin servicio.xlsx"

# ── Paths del módulo padre que este submódulo consume ─────────────────────────
# Se derivan localmente para no crear dependencia de import con 1_lecturas/config.py.
# Si el padre los renombra el fallo es ruidoso y fácil de localizar.
PARENT_ROOT          = ROOT.parent
PARENT_INPUTS        = PARENT_ROOT / "inputs"
REGISTRO_ACUMULADO_PATH = PARENT_INPUTS / "registro_operario_acumulado.xlsx"
REGISTRO_MES_PATH       = PARENT_INPUTS / "registro_operario_mes.xlsx"

def validacion_ausencias_path(mes_ano: str) -> Path:
    return OUTPUTS_DIR / f"validacion_ausencias_{mes_ano}.xlsx"

# ── Nombres de hojas en los xlsx ──────────────────────────────────────────────
SHEET_LISTA = "CatálogoSinServicio"
SHEET_AUDIT = "Auditoría"

def sheet_ausencias(mes_ano: str) -> str:
    return f"Ausencias_{mes_ano}"
