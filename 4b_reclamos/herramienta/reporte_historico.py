"""Reporte historico mensual basado solo en el ledger oficial comprometido.

El PDF conserva el formato visual historico, pero no proyecta ciclos abiertos ni
mezcla precursores con saldos. Junio/julio tienen cobertura parcial; la cuenta
completa comienza en agosto de 2026.

Uso:
    py reporte_historico.py MZ LT
    py reporte_historico.py --con-deuda [MES_ANO]
    py reporte_historico.py --sin-deuda [MES_ANO]
    py reporte_historico.py --todos [MES_ANO]
"""

from __future__ import annotations

import sys
import unicodedata
from functools import lru_cache
from pathlib import Path

import fitz
import pandas as pd


BASE_DIR = Path(__file__).parent
REPO_ROOT = BASE_DIR.parent.parent
sys.path.insert(0, str(REPO_ROOT / "shared"))

import seguimiento_repo as repo  # noqa: E402
import utils_estado_ciclo as repo_estado  # noqa: E402
import comun  # noqa: E402


ESTADO_CICLO_PATH = REPO_ROOT / "shared" / "reporte_acumulado_procesado" / "estado_ciclo.json"
OUTPUTS = BASE_DIR.parent / "outputs"
APORTES_TANQUE_PATH = REPO_ROOT / "shared" / "aportes_tanque_manuales.xlsx"
MES_CUENTA_COMPLETA = "2026-08"
TOL = 0.005

CONCEPTOS = (
    "AGUA", "MANTENIMIENTO", "CORTE_RECONEXION",
    "CONVENIO", "MULTA", "ACUERDOS",
)
CAMPO = {
    "AGUA": "CONSUMO",
    "MANTENIMIENTO": "MANT",
    "CORTE_RECONEXION": "CORTE",
    "CONVENIO": "CONVENIO",
    "MULTA": "MULTA",
    "ACUERDOS": "ACUERDOS",
}
CAMPOS_TABLA = ("CONSUMO", "MANT", "MES_ANT", "CORTE", "CONVENIO", "MULTA", "ACUERDOS")

_AZUL, _AZUL_BG, _GRIS, _NEGRO, _VERDE, _ROJO, _ZEBRA = (
    comun._AZUL, comun._AZUL_BG, comun._GRIS, comun._NEGRO,
    comun._VERDE, comun._ROJO, comun._ZEBRA,
)
_PAGE_W, _PAGE_H, _M = comun._PAGE_W, comun._PAGE_H, comun._M


def _norm(v) -> str:
    s = str(v).strip().upper()
    return s[:-2] if s.endswith(".0") else s


def _mes_comprometido(solicitado: str | None = None) -> str:
    ultimo = repo_estado.ultimo_ledger_comprometido(ruta=ESTADO_CICLO_PATH)
    if ultimo is None:
        raise RuntimeError("No hay ningun ciclo comprometido en el ledger")
    if solicitado and solicitado > ultimo:
        raise ValueError(f"{solicitado} no esta comprometido; ultimo oficial: {ultimo}")
    return solicitado or ultimo


def _eventos_comprometidos(mes: str | None = None) -> tuple[pd.DataFrame, str]:
    corte = _mes_comprometido(mes)
    eventos = repo._leer_eventos()
    if eventos.empty:
        return eventos, corte
    eventos = eventos[eventos["MES"].astype(str) <= corte].copy()
    eventos["MZ"] = eventos["MZ"].map(_norm)
    eventos["LT"] = eventos["LT"].map(_norm)
    eventos["CONCEPTO"] = eventos["CONCEPTO"].astype(str).str.strip().str.upper()
    return eventos, corte


def _meses_ledger(eventos: pd.DataFrame, corte: str) -> list[str]:
    meses = sorted(m for m in eventos["MES"].astype(str).unique() if m <= corte)
    return meses or [corte]


def tabla_predio_ledger(mz: str, lt: str, eventos: pd.DataFrame, corte: str) -> pd.DataFrame:
    """Una fila por mes; importes derivados exclusivamente de eventos activos."""
    mz, lt = _norm(mz), _norm(lt)
    propio = eventos[(eventos["MZ"] == mz) & (eventos["LT"] == lt)].copy()
    propio = propio.sort_values(["MES", "TIMESTAMP"])
    saldos = {concepto: 0.0 for concepto in CONCEPTOS}
    filas = []

    for mes in _meses_ledger(eventos, corte):
        fila = {"MES": mes, "COBERTURA": "COMPLETA" if mes >= MES_CUENTA_COMPLETA else "PARCIAL"}
        for campo in CAMPOS_TABLA:
            for tipo in ("DEUDA", "PAGO", "AJUSTE", "SALDO"):
                fila[f"{tipo}_{campo}"] = 0.0

        for concepto in CONCEPTOS:
            campo = CAMPO[concepto]
            bloque = propio[
                (propio["MES"].astype(str) == mes)
                & (propio["CONCEPTO"] == concepto)
            ]
            saldo_inicial = saldos[concepto]
            cargo = float(pd.to_numeric(bloque["CARGO"], errors="coerce").fillna(0).sum())
            pago = float(pd.to_numeric(bloque["PAGO"], errors="coerce").fillna(0).sum())
            ajustes_visibles = bloque[
                ~(
                    (bloque["CLASE"].astype(str).str.strip() == "CORRECCION_SISTEMA")
                    & (bloque["SOURCE"].astype(str).str.strip() == "correccion_genesis_formula")
                )
            ]
            ajuste = float(pd.to_numeric(ajustes_visibles["AJUSTE"], errors="coerce").fillna(0).sum())
            saldo_final = float(bloque.iloc[-1]["SALDO"]) if not bloque.empty else saldo_inicial
            saldos[concepto] = saldo_final

            fila[f"DEUDA_{campo}"] = round(saldo_inicial + cargo, 2)
            fila[f"PAGO_{campo}"] = round(pago, 2)
            fila[f"AJUSTE_{campo}"] = round(ajuste, 2)
            fila[f"SALDO_{campo}"] = round(saldo_final, 2)

        for tipo in ("DEUDA", "PAGO", "AJUSTE", "SALDO"):
            fila[f"{tipo}_TOTAL"] = round(sum(fila[f"{tipo}_{c}"] for c in CAMPOS_TABLA), 2)
        filas.append(fila)
    return pd.DataFrame(filas)


def _fila_historica_reporte(mes: str, fuente: dict | pd.Series | None) -> dict:
    fila = {"MES": mes, "COBERTURA": "HISTORICO"}
    for campo in CAMPOS_TABLA:
        deuda = comun._numf(fuente.get(f"DEUDA_{campo}")) if fuente is not None else 0.0
        pago = comun._numf(fuente.get(campo)) if fuente is not None else 0.0
        fila[f"DEUDA_{campo}"] = deuda
        fila[f"PAGO_{campo}"] = pago
        fila[f"AJUSTE_{campo}"] = 0.0
        fila[f"SALDO_{campo}"] = round(deuda - pago, 2)
    for tipo in ("DEUDA", "PAGO", "AJUSTE", "SALDO"):
        fila[f"{tipo}_TOTAL"] = round(sum(fila[f"{tipo}_{campo}"] for campo in CAMPOS_TABLA), 2)
    return fila


@lru_cache(maxsize=1)
def _cargar_historicos() -> dict[str, pd.DataFrame]:
    return comun._cargar_historicos()


@lru_cache(maxsize=1)
def _mapa_raw() -> dict[tuple[str, str], tuple[str, str]]:
    return comun._cargar_mapa_raw()


def _coordenada_historica(mz: str, lt: str) -> tuple[str, str]:
    actual = (_norm(mz), _norm(lt))
    return _mapa_raw().get(actual, actual)


def _filas_historicas(mz: str, lt: str) -> pd.DataFrame:
    mz, lt = _norm(mz), _norm(lt)
    mz_hist, lt_hist = _coordenada_historica(mz, lt)
    nombre = repo._lookup_nombres().get((mz, lt), "")
    historicos = _cargar_historicos()
    filas = []
    for _, mes in comun._ARCHIVOS_HISTORICOS:
        fuente = comun._fila_historica(mz_hist, lt_hist, historicos.get(mes), mes, nombre)
        filas.append(_fila_historica_reporte(mes, fuente))
    return pd.DataFrame(filas)


@lru_cache(maxsize=None)
def _cargar_data_boletas(mes: str) -> pd.DataFrame:
    ruta = comun._repo_de_ciclo(mes) / "3_boletas" / "inputs" / "DATA_boletas.xlsx"
    if not ruta.exists():
        return pd.DataFrame()
    df = pd.read_excel(ruta, sheet_name="Data")
    df.columns = [str(c).strip().upper() for c in df.columns]
    return df


@lru_cache(maxsize=None)
def _cargar_planilla_cobrado(mes: str) -> pd.DataFrame:
    ruta = dict(comun._ciclos_recientes()).get(mes)
    if ruta is None or not ruta.exists():
        return pd.DataFrame()
    return pd.read_excel(ruta, sheet_name="planilla_cobrado", header=1)


def _fila_predio(df: pd.DataFrame, mz: str, lt: str) -> pd.Series | None:
    if df.empty or "MZ" not in df.columns or "LT" not in df.columns:
        return None
    sub = df[(df["MZ"].map(_norm) == _norm(mz)) & (df["LT"].map(_norm) == _norm(lt))]
    return None if sub.empty else sub.iloc[0]


@lru_cache(maxsize=1)
def _nombres_actuales() -> dict[tuple[str, str], str]:
    return repo._lookup_nombres()


def _mismo_nombre(a: str, b: str) -> bool:
    def tokens(nombre: str) -> set[str]:
        plano = unicodedata.normalize("NFKD", str(nombre)).encode("ascii", "ignore").decode().upper()
        return set(plano.split())

    ta, tb = tokens(a), tokens(b)
    return bool(ta and tb and (ta == tb or len(ta & tb) >= 2))


def _coordenada_ciclo(mz: str, lt: str, mes: str) -> tuple[str, str] | None:
    actual = (_norm(mz), _norm(lt))
    historica = _coordenada_historica(mz, lt)
    nombre_actual = _nombres_actuales().get(actual, "")
    boletas = _cargar_data_boletas(mes)
    for coordenada in dict.fromkeys((actual, historica)):
        fila = _fila_predio(boletas, *coordenada)
        if fila is not None and _mismo_nombre(nombre_actual, str(fila.get("NOMBRES", ""))):
            return coordenada
    return None


@lru_cache(maxsize=None)
def _pago_confirmado_ciclo(mz: str, lt: str, mes: str) -> float:
    refs = comun.referencias_pago(mz, lt, tabla=None, incluir_overlays=True)
    total = sum(
        float(ref["MONTO"])
        for ref in refs
        if str(ref.get("MES_APLICA", ref.get("MES", "")))[:7] == mes
    )
    aporte_tanque = sum(a["MONTO"] for a in _aportes_tanque(mz, lt, mes) if a["MES"] == mes)
    return round(max(total - aporte_tanque, 0.0), 2)


def _foto_boleta(mz: str, lt: str, mes: str) -> tuple[dict[str, float], dict[str, float]] | None:
    coordenada = _coordenada_ciclo(mz, lt, mes)
    if coordenada is None:
        return None
    boleta = _fila_predio(_cargar_data_boletas(mes), *coordenada)
    if boleta is None:
        return None

    deuda = {
        "CONSUMO": comun._numf(boleta.get("TOTAL MES ACTUAL")),
        "MANT": comun._numf(boleta.get("MANTENIMIENTO")),
        "MES_ANT": comun._numf(boleta.get("MES ANTERIOR")),
        "CORTE": comun._numf(boleta.get("CORTE Y RECONEXION")),
        "CONVENIO": comun._numf(boleta.get("CONVENIO")),
        "MULTA": comun._numf(boleta.get("MULTA (FAENA + REUNION)", boleta.get("MULTA (FAENA + REUNIÓN)"))),
        "ACUERDOS": comun._numf(boleta.get("CUOTA DIRECTA")),
    }
    restante = min(sum(deuda.values()), _pago_confirmado_ciclo(*coordenada, mes))
    pago = {campo: 0.0 for campo in CAMPOS_TABLA}
    for campo in ("CONSUMO", "MANT", "MES_ANT", "CORTE", "MULTA", "ACUERDOS", "CONVENIO"):
        pago[campo] = round(min(max(deuda[campo], 0.0), restante), 2)
        restante = round(max(restante - pago[campo], 0.0), 2)
    return deuda, pago


def tabla_predio_reporte(mz: str, lt: str, eventos: pd.DataFrame, corte: str) -> pd.DataFrame:
    tabla = tabla_predio_ledger(mz, lt, eventos, corte)
    for indice, fila in tabla.iterrows():
        mes = str(fila["MES"])
        if mes >= MES_CUENTA_COMPLETA:
            continue
        foto = _foto_boleta(mz, lt, mes)
        if foto is None:
            continue
        deuda, pago = foto
        tabla.at[indice, "COBERTURA"] = "DATA_BOLETAS"
        for campo in CAMPOS_TABLA:
            tabla.at[indice, f"DEUDA_{campo}"] = deuda[campo]
            tabla.at[indice, f"PAGO_{campo}"] = pago[campo]
            tabla.at[indice, f"SALDO_{campo}"] = round(deuda[campo] - pago[campo], 2)
        for tipo in ("DEUDA", "PAGO", "SALDO"):
            tabla.at[indice, f"{tipo}_TOTAL"] = round(
                sum(float(tabla.at[indice, f"{tipo}_{campo}"]) for campo in CAMPOS_TABLA), 2
            )
    return pd.concat([_filas_historicas(mz, lt), tabla], ignore_index=True)


def _ajustes_predio(mz: str, lt: str, eventos: pd.DataFrame) -> pd.DataFrame:
    sub = eventos[
        (eventos["MZ"] == _norm(mz))
        & (eventos["LT"] == _norm(lt))
        & (eventos["TIPO_EVENTO"].astype(str).str.strip() == "AJUSTE")
        & ~(
            (eventos["CLASE"].astype(str).str.strip() == "CORRECCION_SISTEMA")
            & (eventos["SOURCE"].astype(str).str.strip() == "correccion_genesis_formula")
        )
    ].copy()
    columnas = ["MES", "CONCEPTO", "AJUSTE", "CLASE", "SOURCE", "AUDIT_REF", "MOTIVO"]
    return sub[columnas].sort_values(["MES", "CONCEPTO", "AUDIT_REF"]).reset_index(drop=True)


@lru_cache(maxsize=1)
def _cargar_aportes_tanque() -> pd.DataFrame:
    if not APORTES_TANQUE_PATH.exists():
        return pd.DataFrame()
    return pd.read_excel(APORTES_TANQUE_PATH, header=1)


def _aportes_tanque(mz: str, lt: str, corte: str) -> list[dict]:
    df = _cargar_aportes_tanque()
    if df.empty:
        return []
    sub = df[
        (df["MZ"].map(_norm) == _norm(mz))
        & (df["LT"].map(_norm) == _norm(lt))
        & (df["BALDE"].astype(str).str.strip().str.lower() == "tanque")
        & (df["CANAL"].astype(str).str.strip().str.lower() == "yape")
    ]
    aportes = []
    for _, r in sub.iterrows():
        mes = str(r.get("MES_ANO_APLICA") or r.get("MES_CICLO") or "")[:7]
        if mes <= corte:
            aportes.append({"MES": mes, "MONTO": float(r["MONTO"]), "FECHA_REAL": r.get("FECHA_REAL", "")})
    return aportes


def _yapes_crudos(mz: str, lt: str, mes: str) -> list[dict]:
    df = comun._cargar_pagos_yape_crudo(mes)
    if df is None or df.empty:
        return []
    sub = df[(df["MZ"].map(_norm) == _norm(mz)) & (df["LOTE"].map(_norm) == _norm(lt))]
    yapes = []
    for _, r in sub.iterrows():
        asignado = pd.to_numeric(r.get("MONTO_ASIGNADO"), errors="coerce")
        monto = asignado if pd.notna(asignado) else pd.to_numeric(r.get("MONTO_PAGO"), errors="coerce")
        if pd.notna(monto):
            yapes.append({"MONTO": float(monto), "FECHA_HORA": str(r.get("FECHA", ""))})
    return yapes


def _referencias(mz: str, lt: str, corte: str, tabla: pd.DataFrame | None = None) -> list[dict]:
    tabla_refs = None
    if tabla is not None:
        tabla_refs = tabla[["MES", "PAGO_TOTAL"]].rename(columns={"PAGO_TOTAL": "TOTAL"})
    actual = (_norm(mz), _norm(lt))
    meses_reasignables = {"2026-06", "2026-07"}
    coordenadas = {mes: _coordenada_ciclo(mz, lt, mes) for mes in meses_reasignables}
    fuentes = {actual, *(c for c in coordenadas.values() if c is not None)}
    refs_por_fuente = {
        c: comun.referencias_pago(*c, tabla=tabla_refs, incluir_overlays=True) for c in fuentes
    }
    aportes_por_fuente = {c: _aportes_tanque(*c, corte) for c in fuentes}
    refs = [
        r for r in refs_por_fuente[actual]
        if str(r.get("MES_APLICA", r.get("MES", "")))[:7] not in meses_reasignables
    ]
    aportes = [a for a in aportes_por_fuente[actual] if a["MES"] not in meses_reasignables]
    for mes, coordenada in coordenadas.items():
        if coordenada is None:
            continue
        refs.extend(
            r for r in refs_por_fuente[coordenada]
            if str(r.get("MES_APLICA", r.get("MES", "")))[:7] == mes
        )
        aportes.extend(a for a in aportes_por_fuente[coordenada] if a["MES"] == mes)

    visibles = []
    for ref in refs:
        mes = str(ref.get("MES_APLICA", ref.get("MES", "")))[:7]
        if not comun._ARCHIVOS_HISTORICOS[0][1] <= mes <= corte:
            continue
        copia = dict(ref)
        copia["MES"] = mes
        copia["ESTADO_LEDGER"] = (
            "NO ASENTADO EN LEDGER"
            if ref.get("MEDIO") in {"ABONO REZ.", "BLANCO EF."}
            else "PAGO REGISTRADO"
        )
        del_mes = [a for a in aportes if a["MES"] == mes] if ref.get("MEDIO") == "YAPE" else []
        if not del_mes:
            visibles.append(copia)
            continue

        mz_ref, lt_ref = coordenadas.get(mes) or actual
        yapes = _yapes_crudos(mz_ref, lt_ref, mes)
        usados = set()
        for aporte in del_mes:
            fecha_aporte = pd.to_datetime(aporte["FECHA_REAL"], dayfirst=True, errors="coerce")
            for indice, yape in enumerate(yapes):
                fecha_yape = pd.to_datetime(yape["FECHA_HORA"], dayfirst=True, errors="coerce")
                misma_fecha = pd.notna(fecha_aporte) and pd.notna(fecha_yape) and fecha_aporte.date() == fecha_yape.date()
                if indice not in usados and misma_fecha and abs(yape["MONTO"] - aporte["MONTO"]) <= TOL:
                    usados.add(indice)
                    aporte["FECHA_HORA"] = yape["FECHA_HORA"]
                    break

        monto_pago = float(ref["MONTO"]) - sum(a["MONTO"] for a in del_mes)
        if monto_pago > TOL:
            fechas_pago = [y["FECHA_HORA"] for i, y in enumerate(yapes) if i not in usados]
            copia["MONTO"] = round(monto_pago, 2)
            if fechas_pago:
                copia["FECHA_HORA"] = " · ".join(fechas_pago)
            visibles.append(copia)
        for aporte in del_mes:
            visibles.append({
                "MES": mes,
                "MEDIO": "APORTE TANQUE",
                "FECHA_HORA": aporte.get("FECHA_HORA") or str(aporte["FECHA_REAL"]),
                "ESTADO_LEDGER": "NO REDUCE DEUDA",
                "MONTO": aporte["MONTO"],
            })
    return sorted(visibles, key=lambda r: str(r["MES"]))


def _saldo_actual(tabla: pd.DataFrame) -> dict[str, float]:
    if tabla.empty:
        return {c: 0.0 for c in CAMPOS_TABLA} | {"TOTAL": 0.0}
    ultima = tabla.iloc[-1]
    saldos = {c: float(ultima[f"SALDO_{c}"]) for c in CAMPOS_TABLA}
    saldos["TOTAL"] = round(sum(saldos.values()), 2)
    return saldos


def _texto_monto(v: float, *, firmado: bool = False) -> str:
    if abs(v) <= TOL:
        return "-"
    return f"{v:+,.2f}" if firmado else f"{v:,.2f}"


def _dibujar_tabla_ledger(page, x: float, y: float, w: float, tabla: pd.DataFrame) -> float:
    headers = ["Mes", "Tipo", "Agua", "Mant.", "Mes ant.", "Corte", "Convenio", "Multa", "Acuerdos", "Total"]
    resto = w - 58 - 52 - 78
    ancho_concepto = resto / 7
    anchos = [58, 52, *([ancho_concepto] * 7), 78]
    alto = 13

    page.draw_rect(fitz.Rect(x, y, x + w, y + alto), fill=_AZUL_BG, color=None)
    cx = x
    for header, ancho in zip(headers, anchos):
        page.insert_text((cx + 4, y + alto - 4), header, fontsize=7, fontname="hebo", color=_AZUL)
        cx += ancho
    y += alto

    tipos = (
        ("DEUDA", _NEGRO, False),
        ("PAGO", _VERDE, False),
        ("AJUSTE", _AZUL, True),
        ("SALDO", _ROJO, False),
    )
    for numero, (_, mes) in enumerate(tabla.iterrows()):
        tipos_mes = (tipos[0],)
        if abs(float(mes["PAGO_TOTAL"])) > TOL:
            tipos_mes += (tipos[1],)
        if any(abs(float(mes[f"AJUSTE_{campo}"])) > TOL for campo in CAMPOS_TABLA):
            tipos_mes += (tipos[2],)
        if numero == len(tabla) - 1:
            tipos_mes += (tipos[3],)
        if numero % 2:
            page.draw_rect(fitz.Rect(x, y, x + w, y + alto * len(tipos_mes)), fill=_ZEBRA, color=None)
        parcial = mes["COBERTURA"] != "COMPLETA"
        for tipo, color, firmado in tipos_mes:
            cx = x
            etiqueta_mes = f"{mes['MES']}*" if parcial else str(mes["MES"])
            page.insert_text(
                (cx + 4, y + alto - 4), etiqueta_mes if tipo == "DEUDA" else "",
                fontsize=6.5, fontname="hebo", color=_ROJO if parcial else _NEGRO,
            )
            cx += anchos[0]
            etiqueta_tipo = "PAGO SIM." if tipo == "PAGO" and mes["COBERTURA"] == "DATA_BOLETAS" else tipo
            page.insert_text((cx + 4, y + alto - 4), etiqueta_tipo, fontsize=6.5, fontname="hebo", color=color)
            cx += anchos[1]
            for campo, ancho in zip(CAMPOS_TABLA, anchos[2:-1]):
                valor = float(mes[f"{tipo}_{campo}"])
                texto = _texto_monto(valor, firmado=firmado)
                tw = fitz.get_text_length(texto, fontname="helv", fontsize=6.5)
                page.insert_text((cx + ancho - 5 - tw, y + alto - 4), texto, fontsize=6.5, fontname="helv", color=color)
                cx += ancho
            total = float(mes[f"{tipo}_TOTAL"])
            texto = _texto_monto(total, firmado=firmado)
            tw = fitz.get_text_length(texto, fontname="hebo", fontsize=6.5)
            page.insert_text((cx + anchos[-1] - 5 - tw, y + alto - 4), texto, fontsize=6.5, fontname="hebo", color=color)
            y += alto
    return y


def _dibujar_ajustes(page, y: float, ajustes: pd.DataFrame) -> float:
    if ajustes.empty:
        return y
    page.insert_text((_M, y), "AJUSTES AUDITABLES", fontsize=8, fontname="hebo", color=_AZUL)
    y += 11
    for _, ajuste in ajustes.iterrows():
        motivo = str(ajuste.get("MOTIVO", "") or "sin motivo registrado")
        texto = (
            f"{ajuste['MES']} | {ajuste['CONCEPTO']} {_texto_monto(float(ajuste['AJUSTE']), firmado=True)} | "
            f"{ajuste['CLASE']} | {motivo}"
        )
        page.insert_text((_M, y), texto[:165], fontsize=6.5, fontname="helv", color=_NEGRO)
        y += 9
    return y


def _dibujar_pagina_ledger(doc, mz: str, lt: str, nombre: str, tabla: pd.DataFrame,
                           ajustes: pd.DataFrame) -> None:
    page = doc.new_page(width=_PAGE_W, height=_PAGE_H)
    w = _PAGE_W - 2 * _M
    y = _M
    page.draw_rect(fitz.Rect(_M, y, _M + w, y + 30), fill=_AZUL, color=None)
    page.insert_text(
        (_M + 10, y + 20), f"Predio {mz}-{lt} - {nombre or '(sin nombre)'} - historial mensual de deuda y pagos",
        fontsize=12, fontname="hebo", color=(1, 1, 1),
    )
    y += 42
    page.insert_text(
        (_M, y), "Octubre-mayo: planillas historicas. Junio/julio: DATA_boletas. Agosto: ledger.",
        fontsize=9, fontname="helv", color=_GRIS,
    )
    y += 14
    y = _dibujar_tabla_ledger(page, _M, y, w, tabla)
    y += 10

    saldo = _saldo_actual(tabla)
    page.draw_rect(fitz.Rect(_M, y, _M + w, y + 22), fill=(0.99, 0.94, 0.90) if saldo["TOTAL"] > TOL else _AZUL_BG, color=None)
    page.insert_text((_M + 6, y + 15), "SALDO VIGENTE", fontsize=9, fontname="hebo", color=_NEGRO)
    detalle = " | ".join(
        f"{nombre}: S/{saldo[campo]:,.2f}"
        for nombre, campo in (("Agua", "CONSUMO"), ("Mant.", "MANT"), ("Corte", "CORTE"),
                              ("Convenio", "CONVENIO"), ("Multa", "MULTA"), ("Acuerdos", "ACUERDOS"))
        if abs(saldo[campo]) > TOL
    ) or "Al dia"
    page.insert_text((_M + 100, y + 15), detalle, fontsize=7.5, fontname="helv", color=_GRIS)
    total = f"TOTAL S/{saldo['TOTAL']:,.2f}"
    tw = fitz.get_text_length(total, fontname="hebo", fontsize=10)
    page.insert_text((_M + w - tw - 8, y + 15), total, fontsize=10, fontname="hebo", color=_ROJO if saldo["TOTAL"] > TOL else _VERDE)
    y += 34

    if y + 12 + 9 * len(ajustes) < _PAGE_H - _M:
        y = _dibujar_ajustes(page, y, ajustes)
    elif not ajustes.empty:
        pagina_ajustes = doc.new_page(width=_PAGE_W, height=_PAGE_H)
        pagina_ajustes.insert_text((_M, _M + 10), f"Predio {mz}-{lt} - ajustes (continuacion)", fontsize=10, fontname="hebo", color=_AZUL)
        _dibujar_ajustes(pagina_ajustes, _M + 26, ajustes)

    page.insert_text(
        (_M, _PAGE_H - 18),
        "* PAGO SIM. (junio/julio) es una reparticion simulada, no un asiento del ledger.",
        fontsize=7, fontname="helv", color=_GRIS,
    )


def generar_pdf(mz: str, lt: str, salida: Path | None = None, mes: str | None = None) -> Path:
    eventos, corte = _eventos_comprometidos(mes)
    mz, lt = _norm(mz), _norm(lt)
    nombre = repo._lookup_nombres().get((mz, lt), "")
    tabla = tabla_predio_reporte(mz, lt, eventos, corte)
    ajustes = _ajustes_predio(mz, lt, eventos)
    refs = _referencias(mz, lt, corte, tabla)

    doc = fitz.open()
    _dibujar_pagina_ledger(doc, mz, lt, nombre, tabla, ajustes)
    comun._dibujar_pagina_referencias(doc, mz, lt, nombre, refs)
    salida = salida or OUTPUTS / f"reporte_historico_ledger_{mz}-{lt}_{corte}.pdf"
    salida.parent.mkdir(parents=True, exist_ok=True)
    salida.unlink(missing_ok=True)
    doc.save(str(salida))
    doc.close()
    return salida


def _universo(eventos: pd.DataFrame) -> set[tuple[str, str]]:
    predios = set(repo._lookup_nombres())
    predios.update(zip(eventos["MZ"], eventos["LT"]))
    return {(_norm(mz), _norm(lt)) for mz, lt in predios if _norm(mz) not in {"", "NAN"} and _norm(lt) not in {"", "NAN"}}


def _saldos_vigentes(eventos: pd.DataFrame) -> pd.DataFrame:
    if eventos.empty:
        return pd.DataFrame(columns=["MZ", "LT", "TOTAL"])
    ultimos = (
        eventos[eventos["CONCEPTO"].isin(CONCEPTOS)]
        .sort_values(["MZ", "LT", "CONCEPTO", "MES", "TIMESTAMP"])
        .groupby(["MZ", "LT", "CONCEPTO"], as_index=False)
        .last()
    )
    tabla = ultimos.pivot(index=["MZ", "LT"], columns="CONCEPTO", values="SALDO").fillna(0)
    tabla = tabla.reindex(columns=CONCEPTOS, fill_value=0)
    tabla["TOTAL"] = tabla.sum(axis=1)
    return tabla.reset_index()


def _predios_con_deuda(eventos: pd.DataFrame) -> set[tuple[str, str]]:
    saldos = _saldos_vigentes(eventos)
    return {(r["MZ"], r["LT"]) for _, r in saldos[saldos["TOTAL"] > TOL].iterrows()}


def generar_lote(predios: list[tuple[str, str]], salida: Path | None = None,
                 mes: str | None = None) -> Path:
    eventos, corte = _eventos_comprometidos(mes)
    nombres = repo._lookup_nombres()
    doc = fitz.open()
    for numero, (mz, lt) in enumerate(predios, 1):
        mz, lt = _norm(mz), _norm(lt)
        tabla = tabla_predio_reporte(mz, lt, eventos, corte)
        _dibujar_pagina_ledger(doc, mz, lt, nombres.get((mz, lt), ""), tabla, _ajustes_predio(mz, lt, eventos))
        comun._dibujar_pagina_referencias(doc, mz, lt, nombres.get((mz, lt), ""), _referencias(mz, lt, corte, tabla))
        if numero % 25 == 0:
            print(f"  {numero}/{len(predios)} predios")
    salida = salida or OUTPUTS / f"reporte_historico_ledger_{corte}.pdf"
    salida.parent.mkdir(parents=True, exist_ok=True)
    salida.unlink(missing_ok=True)
    doc.save(str(salida))
    doc.close()
    return salida


def generar_excel(predios: list[tuple[str, str]], salida: Path | None = None,
                  mes: str | None = None) -> Path:
    eventos, corte = _eventos_comprometidos(mes)
    nombres = repo._lookup_nombres()
    resumen, mensual, ajustes, referencias = [], [], [], []
    for mz, lt in predios:
        mz, lt = _norm(mz), _norm(lt)
        tabla = tabla_predio_reporte(mz, lt, eventos, corte)
        saldo = _saldo_actual(tabla)
        resumen.append({"MZ": mz, "LT": lt, "NOMBRE": nombres.get((mz, lt), ""), **saldo})
        mensual.append(tabla.assign(MZ=mz, LT=lt, NOMBRE=nombres.get((mz, lt), "")))
        aj = _ajustes_predio(mz, lt, eventos)
        if not aj.empty:
            ajustes.append(aj.assign(MZ=mz, LT=lt, NOMBRE=nombres.get((mz, lt), "")))
        refs = _referencias(mz, lt, corte, tabla)
        if refs:
            referencias.append(pd.DataFrame(refs).assign(MZ=mz, LT=lt, NOMBRE=nombres.get((mz, lt), "")))

    salida = salida or OUTPUTS / f"reporte_historico_ledger_{corte}.xlsx"
    salida.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(salida) as writer:
        pd.DataFrame(resumen).to_excel(writer, sheet_name="Resumen", index=False)
        pd.concat(mensual, ignore_index=True).to_excel(writer, sheet_name="Mensual", index=False)
        pd.concat(ajustes, ignore_index=True).to_excel(writer, sheet_name="Ajustes", index=False) if ajustes else pd.DataFrame().to_excel(writer, sheet_name="Ajustes", index=False)
        pd.concat(referencias, ignore_index=True).to_excel(writer, sheet_name="Referencias", index=False) if referencias else pd.DataFrame().to_excel(writer, sheet_name="Referencias", index=False)
    return salida


def main() -> None:
    args = sys.argv[1:]
    eventos, corte = _eventos_comprometidos()
    if args and args[0] in {"--con-deuda", "--sin-deuda", "--todos"}:
        modo = args[0]
        solicitado = args[1] if len(args) > 1 else corte
        eventos, corte = _eventos_comprometidos(solicitado)
        universo = _universo(eventos)
        con_deuda = _predios_con_deuda(eventos)
        predios = con_deuda if modo == "--con-deuda" else universo - con_deuda if modo == "--sin-deuda" else universo
        predios = sorted(predios)
        print(generar_lote(predios, mes=corte))
        print(generar_excel(predios, mes=corte))
    else:
        mz, lt = (args[0], args[1]) if len(args) >= 2 else ("Q", "5")
        print(generar_pdf(mz, lt, mes=corte))
        print(generar_excel([(_norm(mz), _norm(lt))], salida=OUTPUTS / f"reporte_historico_ledger_{_norm(mz)}-{_norm(lt)}_{corte}.xlsx", mes=corte))


if __name__ == "__main__":
    main()
