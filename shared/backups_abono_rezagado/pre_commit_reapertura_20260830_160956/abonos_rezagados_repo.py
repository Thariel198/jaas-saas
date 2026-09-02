"""Lectura validada de la fuente unica de abonos rezagados."""
from pathlib import Path

import pandas as pd


PATH = Path(__file__).resolve().parent / "abonos_rezagados.xlsx"
HOJA = "Abonos_Raw"
ESTADO_ACTIVO = "CONFIRMADO"
ESTADOS_VALIDOS = {ESTADO_ACTIVO, "DESCARTADO", "CONFIRMADO_SIN_APLICACION"}
MODOS_VALIDOS = {"CASCADA", "DIRIGIDO"}
CONCEPTOS_VALIDOS = {
    "AGUA", "MANTENIMIENTO", "CORTE_RECONEXION", "CONVENIO",
    "ACUERDOS", "MULTA", "OTROS",
}


def _normalizar_columnas(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(col).strip().upper() for col in df.columns]
    return df


def _texto(serie: pd.Series) -> pd.Series:
    return serie.fillna("").astype(str).str.strip()


def _validar_confirmados(df: pd.DataFrame) -> None:
    requeridas = {
        "ID_ABONO", "MZ", "LT", "MONTO", "MES_CICLO", "MES_ANO_APLICA",
        "ESTADO", "MODO_APLICACION", "CONCEPTO_DESTINO",
    }
    faltantes = sorted(requeridas - set(df.columns))
    if faltantes:
        raise RuntimeError(f"abonos_rezagados: faltan columnas {faltantes}")
    if df.empty:
        return

    vacias = {}
    for col in ("ID_ABONO", "MZ", "LT", "MES_CICLO", "MES_ANO_APLICA", "MODO_APLICACION"):
        filas = df.index[_texto(df[col]).eq("")].tolist()
        if filas:
            vacias[col] = [int(i) + 3 for i in filas]
    if vacias:
        raise RuntimeError(f"abonos_rezagados: campos obligatorios vacios {vacias}")

    ids = _texto(df["ID_ABONO"])
    repetidos = sorted(ids[ids.duplicated(keep=False)].unique())
    if repetidos:
        raise RuntimeError(f"abonos_rezagados: ID_ABONO duplicado {repetidos}")

    montos = pd.to_numeric(df["MONTO"], errors="coerce")
    invalidos = df.index[montos.isna() | montos.le(0)].tolist()
    if invalidos:
        raise RuntimeError(
            f"abonos_rezagados: MONTO debe ser positivo en filas {[int(i) + 3 for i in invalidos]}"
        )

    modos = _texto(df["MODO_APLICACION"]).str.upper()
    desconocidos = sorted(set(modos) - MODOS_VALIDOS)
    if desconocidos:
        raise RuntimeError(f"abonos_rezagados: MODO_APLICACION invalido {desconocidos}")

    dirigidos = df[modos.eq("DIRIGIDO")]
    conceptos = _texto(dirigidos["CONCEPTO_DESTINO"]).str.upper()
    conceptos_invalidos = sorted(set(conceptos) - CONCEPTOS_VALIDOS)
    if conceptos_invalidos:
        raise RuntimeError(f"abonos_rezagados: CONCEPTO_DESTINO invalido {conceptos_invalidos}")

    clave = pd.DataFrame({
        "MZ": _texto(df["MZ"]).str.upper(),
        "LT": _texto(df["LT"]).str.upper().str.replace(r"\.0$", "", regex=True),
        "MONTO": montos.round(2),
        "MES_CICLO": _texto(df["MES_CICLO"]).str[:7],
        "MES_ANO_APLICA": _texto(df["MES_ANO_APLICA"]).str[:7],
        "MODO_APLICACION": modos,
        "CONCEPTO_DESTINO": _texto(df["CONCEPTO_DESTINO"]).str.upper(),
    })
    duplicadas = clave.duplicated(keep=False)
    if duplicadas.any():
        raise RuntimeError(
            "abonos_rezagados: filas activas duplicadas "
            f"{[int(i) + 3 for i in clave.index[duplicadas]]}"
        )


def leer_abonos(path: Path | str = PATH, *, incluir_inactivos: bool = False) -> pd.DataFrame:
    """Lee la fuente; por defecto devuelve solo registros confirmados validos."""
    path = Path(path)
    if not path.exists():
        return pd.DataFrame()
    df = _normalizar_columnas(pd.read_excel(path, sheet_name=HOJA, header=1))
    if "ESTADO" not in df.columns:
        raise RuntimeError("abonos_rezagados: falta ESTADO; ejecutar el reinicio controlado")

    estados = _texto(df["ESTADO"]).str.upper()
    desconocidos = sorted(set(estados) - ESTADOS_VALIDOS)
    if desconocidos:
        raise RuntimeError(f"abonos_rezagados: ESTADO invalido {desconocidos}")

    confirmados = df[estados.eq(ESTADO_ACTIVO)].copy()
    _validar_confirmados(confirmados)
    return df if incluir_inactivos else confirmados
