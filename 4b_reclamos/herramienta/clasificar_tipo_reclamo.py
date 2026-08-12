"""
4b_reclamos/herramienta/clasificar_tipo_reclamo.py — Clasifica TIPO_RECLAMO por palabra clave
en el texto de RECLAMO, y duplica la fila cuando el reclamo mezcla más de un
concepto (ej. "ya pague mes anterior, techado y campo" -> 1 fila mes_anterior +
1 fila cuota).

Reglas de mapeo (palabra clave en RECLAMO -> TIPO_RECLAMO):
    mes anterior / mes pasado          -> mes_anterior
    faena / reunion / "multa"          -> multa
    medidor / "convenio"               -> convenio
    techado / campo / "cuota"          -> cuota

No toca filas donde el texto no matchea ninguna palabra clave (deja lo que
haya, incluido vacío) — esas se clasifican a mano. Si el texto matchea 2+
categorías, la fila se reemplaza por N filas (una por categoría), copiando el
resto de columnas igual; si matchea 1 sola y no coincide con lo que ya tenía
el supervisor, corrige el valor existente.

TIPO_RECLAMO es columna manual preservada por main.py entre corridas — correr
este script ANTES de rearmar la vista con main.py no pierde nada; correrlo
DESPUÉS es el orden normal de cada mes.

Uso (desde 4b_reclamos/herramienta/):
    py clasificar_tipo_reclamo.py                # mes = ciclo activo
    py clasificar_tipo_reclamo.py --mes 2026-08
"""

import argparse
import re
import sys
import unicodedata
from pathlib import Path

import pandas as pd

sys.path.insert(0, str(Path(__file__).parent.parent))
import main as m  # noqa: E402

ORDEN_TIPOS = ["mes_anterior", "multa", "convenio", "cuota"]


def _norm(s) -> str:
    s = str(s).lower()
    return "".join(c for c in unicodedata.normalize("NFD", s) if unicodedata.category(c) != "Mn")


def categorias(texto: str) -> list[str]:
    t = _norm(texto)
    cats = []
    if re.search(r"mes\s*anter", t) or re.search(r"mes\s*pasad", t):
        cats.append("mes_anterior")
    if "faena" in t or "reunion" in t or re.search(r"\bmulta\b", t):
        cats.append("multa")
    if "medidor" in t or re.search(r"\bconvenio\b", t):
        cats.append("convenio")
    if "techado" in t or "campo" in t or re.search(r"\bcuota\b", t):
        cats.append("cuota")
    return [c for c in ORDEN_TIPOS if c in cats]


_IDENTIDAD_COLS = ["MZ", "LT", "MESA", "COBRADOR", "FECHA_COBRO", "MONTO",
                    "MES_ANO_DETECTADO", "MES_ANO_ORIGEN", "RECLAMO",
                    "RESOLUCION", "ESTADO", "FECHA_RESOLUCION"]


def _identidad(row) -> tuple:
    """Todo lo que identifica el mismo reclamo, salvo TIPO_RECLAMO — clave
    de agrupación para no re-duplicar filas ya divididas."""
    vals = []
    for c in _IDENTIDAD_COLS:
        v = row.get(c, "")
        if pd.isna(v):
            v = None
        elif hasattr(v, "isoformat"):
            v = v.isoformat()
        vals.append(v)
    return tuple(vals)


def clasificar(df: pd.DataFrame) -> tuple[pd.DataFrame, list[tuple]]:
    filas = []
    cambios = []
    # Agrupa por identidad (todo menos TIPO_RECLAMO) para que correr el script
    # 2 veces sobre filas ya divididas no las vuelva a triplicar — sin esto,
    # cada fila hija re-matchea el mismo texto y se re-explota (bug real
    # encontrado al probar la corrida repetida).
    grupos: dict = {}
    orden_keys = []
    for _, row in df.iterrows():
        k = _identidad(row)
        if k not in grupos:
            grupos[k] = []
            orden_keys.append(k)
        grupos[k].append(row.to_dict())

    for k in orden_keys:
        grupo = grupos[k]
        texto = grupo[0]["RECLAMO"]
        cats = categorias(texto)

        tipos_existentes = []
        for r in grupo:
            t = r.get("TIPO_RECLAMO", "")
            t = "" if pd.isna(t) else str(t)
            if t:
                tipos_existentes.append(t)

        if not cats:
            filas.extend(grupo)
            continue

        if set(tipos_existentes) == set(cats) and len(grupo) == len(cats):
            filas.extend(grupo)  # ya clasificado correctamente, no tocar
            continue

        cambios.append((grupo[0]["MZ"], grupo[0]["LT"], tipos_existentes, cats, texto))
        base = grupo[0]
        for c in cats:
            r2 = dict(base)
            r2["TIPO_RECLAMO"] = c
            filas.append(r2)

    return pd.DataFrame(filas), cambios


def main_(mes: str) -> None:
    path = m._ruta_vista(mes)
    if not path.exists():
        print(f"No existe {path}")
        return

    df = pd.read_excel(path, sheet_name="Reclamos", header=1)
    df_final, cambios = clasificar(df)

    print(f"Filas originales: {len(df)}  ->  filas finales: {len(df_final)}")
    print(f"Cambios (clasificaciones nuevas + correcciones): {len(cambios)}")
    for mz, lt, antes, despues, texto in cambios:
        print(f"  {mz}-{lt}: '{antes or '(vacio)'}' -> {despues}  |  {texto}")

    if not cambios:
        print("Nada que cambiar.")
        return

    backup = m._backup_con_timestamp(mes)
    if backup:
        print(f"Backup: {backup.name}")

    m._write_vista(df_final, mes)
    print(f"Regenerado: {path}")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Clasifica TIPO_RECLAMO por palabra clave")
    parser.add_argument("--mes", default=m._MES_CICLO, help="Mes a procesar (YYYY-MM)")
    args = parser.parse_args()
    if not args.mes:
        parser.error("--mes requerido (no hay ciclo activo en shared/ciclo_activo.json)")
    main_(args.mes)
