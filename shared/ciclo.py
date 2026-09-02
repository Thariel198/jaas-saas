"""Fuente única del mes que el pipeline está procesando (el "ciclo activo").

Por qué existe: hasta el 05/08/2026 cada módulo re-deducía el mes por su cuenta,
siempre con la misma heurística frágil — "el último archivo por orden alfabético":

    2_planilla   mes = Path(matches[-1]).stem.replace("lecturas_planilla_", "")
    5_cobranza   plan = sorted(PLAN_DIR.glob("planilla_*.xlsx"))[-1]

Con dos o tres meses conviviendo en la misma carpeta (el pipeline se copia entero
por mes), "el último" no es "el que estoy cobrando". El 06/07/2026 eso costó 15
pagos fantasma: 5_cobranza cobró el ciclo 2026-07 leyendo el archivo de pagos yape
que todavía era el de junio, y el residuo de esos pagos cayó sobre la deuda de
julio (ver docs/RETOMAR_auditoria_junio_julio_y_reimputacion_2026-08-04.md).

El dato nunca faltó: la plantilla del operario ya trae la columna MES_ANO. Lo que
faltaba era DECLARARLO en un lugar y que todos lo lean en vez de adivinarlo.

    plantilla del operario (MES_ANO)
        -> 1_lecturas valida que sea uno solo y llama a escribir()
            -> shared/ciclo_activo.json
                -> 2_planilla · 4_pagos · 5_cobranza · … leen con activo()

Uso:
    import ciclo
    mes = ciclo.activo()              # "2026-08" — lanza si no está declarado
    mes = ciclo.activo(default=None)  # None en vez de lanzar
    ciclo.escribir("2026-08", origen="1_lecturas/registro_operario_mes.xlsx")
"""

import json
import re
from datetime import datetime
from pathlib import Path

from work_guard_runtime import require_authoritative_writes

CICLO_PATH = Path(__file__).parent / "ciclo_activo.json"
_FORMATO = re.compile(r"^\d{4}-\d{2}$")

_FALTA = (
    f"No hay ciclo activo declarado ({CICLO_PATH}).\n"
    f"  -> correr 1_lecturas (declara el mes desde la columna MES_ANO de la\n"
    f"     plantilla del operario), o escribirlo a mano con ciclo.escribir()"
)


_MESES_ES = ("enero", "febrero", "marzo", "abril", "mayo", "junio",
             "julio", "agosto", "setiembre", "octubre", "noviembre", "diciembre")

# Desde este ciclo, todo output de 4_pagos/5_cobranza nace con el periodo en el
# nombre. Los ciclos anteriores quedaron con el nombre pelado y hay que poder
# seguir leyéndolos (re-correr julio, reportes históricos), pero SOLO ellos: si
# el nombre pelado se aceptara también para el ciclo actual, un mes sin generar
# leería el archivo del mes pasado — el incidente del 06/07/2026 exacto.
PRIMER_CICLO_CON_PERIODO = "2026-08"


def acepta_legacy(mes_ano: str) -> bool:
    """True para los ciclos generados antes de la convención con periodo."""
    return validar(mes_ano) < PRIMER_CICLO_CON_PERIODO


def resolver(carpeta: Path, base: str, mes_ano: str, ext: str = ".xlsx",
             legacy_sin_periodo: bool = False) -> Path:
    """Ruta del archivo de UN ciclo, tolerando cómo se llamó antes.

    El nombre canónico lleva el periodo — `planilla_cobrado_2026-06.xlsx` — por la
    convención del CLAUDE.md ("Archivos Excel con periodo: nombre_YYYY-MM.xlsx").
    Es la defensa más simple contra el incidente del 06/07/2026: si el módulo de
    arriba no corrió para este ciclo, el lector no encuentra el archivo y falla
    ruidosamente, en vez de leer el del mes pasado creyendo que es fresco.

    Se acepta además el mes en palabras (`planilla_cobrado_junio.xlsx`), que es
    como quedaron algunos archivos renombrados a mano en los ciclos congelados.

    `legacy_sin_periodo` acepta el nombre pelado (`planilla_cobrado.xlsx`) y por
    defecto está APAGADO a propósito: un archivo sin periodo no dice de qué mes
    es, así que aceptarlo siempre reintroduce el bug que este módulo previene —
    pedir el ciclo 2026-09 y recibir el archivo de julio sin enterarse. Solo lo
    encienden los lectores de un módulo que todavía escribe sin periodo, hasta
    que ese módulo migre.

    Si no existe ninguno se devuelve el canónico, para que el error hable del
    nombre correcto.
    """
    mes_ano = validar(mes_ano)
    mes_palabra = _MESES_ES[int(mes_ano[5:7]) - 1]
    nombres = [f"{base}_{mes_ano}{ext}",
               f"{base}_{mes_palabra}{ext}",
               f"{base}_{mes_palabra.capitalize()}{ext}"]
    if legacy_sin_periodo:
        nombres.append(f"{base}{ext}")
    for nombre in nombres:
        if (carpeta / nombre).exists():
            return carpeta / nombre
    return carpeta / f"{base}_{mes_ano}{ext}"


def validar(mes_ano: str) -> str:
    mes_ano = str(mes_ano).strip()
    if not _FORMATO.match(mes_ano):
        raise ValueError(f"MES_ANO inválido: {mes_ano!r} — se espera YYYY-MM (ej. 2026-08)")
    return mes_ano


def activo(default: str | None = "__lanzar__", path: Path | None = None) -> str | None:
    """Mes que el pipeline está procesando, en formato YYYY-MM.

    `path` existe para que cada módulo pase la ruta derivada de SU config (los
    tests monkey-patchean esas constantes con un shared/ sintético). Sin eso,
    correr un test escribía o leía el shared/ real del repo."""
    path = path or CICLO_PATH
    if not path.exists():
        if default == "__lanzar__":
            raise FileNotFoundError(_FALTA)
        return default
    dato = json.loads(path.read_text(encoding="utf-8"))
    return validar(dato["mes_ano"])


def escribir(mes_ano: str, origen: str = "", path: Path | None = None) -> str:
    """Declara el ciclo activo. Idempotente: re-escribirlo con el mismo mes no
    cambia nada salvo la marca de tiempo."""
    mes_ano = validar(mes_ano)
    path = path or CICLO_PATH
    require_authoritative_writes(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps({
        "mes_ano": mes_ano,
        "origen": origen,
        "declarado_en": datetime.now().isoformat(timespec="seconds"),
    }, ensure_ascii=False, indent=2), encoding="utf-8")
    return mes_ano
