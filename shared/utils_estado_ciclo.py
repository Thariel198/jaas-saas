"""
shared/utils_estado_ciclo.py — Primitivos para estado_ciclo.json

Estado del ciclo mensual: ABIERTO/CERRADO + sello de arrastre (generado/validado).
Tres escritores/lectores antes de este módulo (cada uno con su propia copia del
parseo JSON, ninguno atómico): 5_cobranza (marca generado), 5b_validacion (sella
validado), 2_planilla (solo lee). 7_cierre es el cuarto punto de contacto (marca
CERRADO) — con 3+ llamadores reales, se centraliza acá (ver metodologia_desarrollo.md,
"cross-módulo nunca es import directo").

`ruta` es parámetro explícito (no una constante fija acá) porque cada módulo
resuelve su propio `config.ESTADO_CICLO_PATH` — y los tests de 2_planilla lo
monkeypatchean a un archivo temporal para aislar la corrida.

Escritura atómica: temp file + os.replace(), con retry ante PermissionError
transitorio (mismo patrón que shared/seguimiento_repo.py — un antivirus/indexador
de Windows puede bloquear el archivo un instante).
"""
import json
import os
import time
from datetime import datetime
from pathlib import Path

DEFAULT_PATH = Path(__file__).parent / "reporte_acumulado_procesado" / "estado_ciclo.json"


def _leer(ruta: Path) -> dict:
    if not ruta.exists():
        return {}
    try:
        return json.loads(ruta.read_text(encoding="utf-8"))
    except Exception:
        return {}


def _guardar_atomic(data: dict, ruta: Path) -> None:
    ruta.parent.mkdir(parents=True, exist_ok=True)
    tmp = ruta.with_suffix(".tmp")
    tmp.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    for intento in range(5):
        try:
            os.replace(tmp, ruta)
            return
        except PermissionError:
            if intento == 4:
                raise
            time.sleep(0.3)


def ciclo_validado(mes: str, ruta: Path = DEFAULT_PATH) -> bool:
    """¿estado_ciclo[mes].arrastre.validado == true?"""
    return bool(_leer(ruta).get(mes, {}).get("arrastre", {}).get("validado"))


def ciclo_cerrado(mes: str, ruta: Path = DEFAULT_PATH) -> bool:
    """¿estado_ciclo[mes].estado == 'CERRADO'?"""
    return _leer(ruta).get(mes, {}).get("estado") == "CERRADO"


def marcar_generado(mes: str, ruta: Path = DEFAULT_PATH) -> None:
    """Marca arrastre.generado=true (idempotente). 5b lo lee para sellar validado."""
    data = _leer(ruta)
    ciclo = data.setdefault(mes, {"estado": "ABIERTO"})
    arr = ciclo.setdefault("arrastre", {})
    if not arr.get("generado"):
        ahora = datetime.now().strftime("%Y-%m-%dT%H:%M:%S")
        arr["generado"] = True
        arr["generado_en"] = ahora
        ciclo["ultima_actualizacion"] = ahora
        _guardar_atomic(data, ruta)


def sellar_validado(ruta: Path = DEFAULT_PATH) -> list[str]:
    """Marca arrastre.validado=true para todo ciclo con generado=true y validado
    aún pendiente (5b_validacion). Devuelve la lista de meses sellados."""
    data = _leer(ruta)
    if not data:
        return []
    ahora = datetime.now().strftime("%Y-%m-%dT%H:%M:%S")
    sellado = []
    for mes, info in data.items():
        arr = info.get("arrastre", {})
        if arr.get("generado") and not arr.get("validado"):
            arr["validado"] = True
            arr["validado_en"] = ahora
            info["ultima_actualizacion"] = ahora
            sellado.append(mes)
    if sellado:
        _guardar_atomic(data, ruta)
    return sellado


def marcar_cerrado(mes: str, ruta: Path = DEFAULT_PATH) -> None:
    """Marca estado=CERRADO (idempotente). Usado por 7_cierre al transicionar el período."""
    data = _leer(ruta)
    ciclo = data.setdefault(mes, {"estado": "ABIERTO"})
    if ciclo.get("estado") != "CERRADO":
        ahora = datetime.now().strftime("%Y-%m-%dT%H:%M:%S")
        ciclo["estado"] = "CERRADO"
        ciclo["cerrado_en"] = ahora
        ciclo["ultima_actualizacion"] = ahora
        _guardar_atomic(data, ruta)
