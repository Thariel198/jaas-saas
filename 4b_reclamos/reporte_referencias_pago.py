"""
4b_reclamos/reporte_referencias_pago.py — Extiende reporte_historico.py con una
seccion de "referencia de pago" debajo del historial mensual: de donde vino
cada pago (yape con fecha/hora exacta, o efectivo) y su monto.

Fuentes por rango de fecha:
  - Oct 2025 -> Ene 2026: hoja "Cobranza" (Estado=c + Medio) decide yape/efectivo;
    si es yape, la hoja "Reporte" del mismo archivo trae fecha/hora exacta.
  - Feb 2026 -> May 2026: mismo criterio, pero el monto de efectivo sale de la
    hoja "Efectivo" dedicada (mas preciso que el total de Cobranza).
  - Jun 2026 -> Jul 2026: no hay archivo que preserve fecha/hora por transaccion
    para la mayoria de pagos yape (motor_matching no lo persiste) -- se muestra
    el total del ciclo desde planilla_cobrado.xlsx, sin desglose de hora.

Uso: py 4b_reclamos/reporte_referencias_pago.py
"""

import sys
import unicodedata
from pathlib import Path

import fitz
import pandas as pd

BASE_DIR = Path(__file__).parent
sys.path.insert(0, str(BASE_DIR.parent / "shared"))
sys.path.insert(0, str(BASE_DIR))
import seguimiento_repo as repo  # noqa: E402
import ciclo  # noqa: E402
import reporte_historico as rh  # noqa: E402
import reporte_convenio_multa as rcm  # noqa: E402

HIST_DIR = BASE_DIR.parent / "obligaciones" / "inputs" / "planillas anteriores"

_ARCHIVOS_HISTORICOS = rh._ARCHIVOS_HISTORICOS

_AZUL = (26/255, 82/255, 118/255)
_AZUL_BG = (235/255, 245/255, 251/255)
_GRIS = (0.42, 0.45, 0.5)
_NEGRO = (0.12, 0.16, 0.22)
_VERDE = (0.02, 0.37, 0.27)
_ZEBRA = (243/255, 244/255, 246/255)


def _norm(s: str) -> str:
    s = unicodedata.normalize("NFKD", str(s)).encode("ascii", "ignore").decode()
    return s.lower().replace("\n", " ").strip()


def _col(df: pd.DataFrame, *candidatos: str):
    normed = {}
    for c in df.columns:
        k = _norm(c)
        if k not in normed:
            normed[k] = c
    for cand in candidatos:
        cn = _norm(cand)
        if cn in normed:
            return normed[cn]
    for cand in candidatos:
        cn = _norm(cand)
        for k, real in normed.items():
            if cn in k:
                return real
    return None


_cache_hojas: dict[str, dict] = {}
_avisados: set[str] = set()


def _avisar_falta(p: Path, que: str) -> None:
    """Un archivo de ciclo cerrado que no esta donde dice el path se avisa UNA
    vez por corrida (esto se llama por predio). Fallaba en silencio hasta el
    04/08/2026 y las filas sin origen se leian como pagos fantasma."""
    if str(p) in _avisados:
        return
    _avisados.add(str(p))
    print(f"  AVISO: falta {que} -> {p}")


def _cargar_hojas_historicas(archivo: str) -> dict:
    """Cachea Cobranza/Reporte/Efectivo de UN archivo historico a nivel modulo
    -- sin esto, una corrida por lote reabre los mismos 8 excels una vez POR
    PREDIO (~500 predios x 8 archivos = ~4000 lecturas en vez de 8)."""
    if archivo in _cache_hojas:
        return _cache_hojas[archivo]
    p = HIST_DIR / archivo
    hojas = {}
    if p.exists():
        hojas["Cobranza"] = pd.read_excel(p, sheet_name="Cobranza", header=0)
        try:
            hojas["Reporte"] = pd.read_excel(p, sheet_name="Reporte", header=0)
        except Exception:
            hojas["Reporte"] = None
    _cache_hojas[archivo] = hojas
    return hojas


def referencias_pago(mz: str, lt: str, tabla: pd.DataFrame | None = None) -> list[dict]:
    """Lista de {MES, MEDIO, FECHA_HORA, MONTO} -- una entrada por pago detectado.

    El MONTO de cada mes SIEMPRE es el TOTAL que ya calculo _fila_historica()
    (mismo numero que aparece en la tabla de arriba) -- no una relectura aparte
    de una columna suelta. Hasta mayo 2026 ese total ya incluye mantenimiento
    arrastrado y exceso devuelto via "mes anterior" (asi se manejaba antes);
    la hoja Reporte/Efectivo solo aporta DONDE vino la plata (yape con hora,
    o efectivo), no el monto -- eso evita que un monto parcial de una sola
    transaccion descuadre contra el total real del mes."""
    out = []

    mapa_raw = rh._cargar_mapa_raw()
    mz_hist, lt_hist = mapa_raw.get((mz, lt), (mz, lt))

    for archivo, mes in _ARCHIVOS_HISTORICOS:
        hojas = _cargar_hojas_historicas(archivo)
        dfc = hojas.get("Cobranza")
        if dfc is None:
            continue
        cmz, clt, cest, cmed = _col(dfc, "mz"), _col(dfc, "lt", "lote"), _col(dfc, "estado"), _col(dfc, "medio")
        sub = dfc[(dfc[cmz].astype(str).str.strip() == mz_hist) & (dfc[clt].astype(str).str.strip() == lt_hist)]
        if sub.empty or str(sub.iloc[0][cest]).strip().lower() != "c":
            continue

        fila_mes = None
        if tabla is not None:
            m = tabla[tabla["MES"] == mes]
            if not m.empty:
                fila_mes = m.iloc[0]
        monto_mes = float(fila_mes["TOTAL"]) if fila_mes is not None else None
        if monto_mes is None or monto_mes <= 0.005:
            continue

        medio = str(sub.iloc[0][cmed]).strip().lower()
        if medio == "y":
            dfr = hojas.get("Reporte")
            if dfr is None:
                out.append({"MES": mes, "MEDIO": "YAPE", "FECHA_HORA": "sin hoja Reporte", "MONTO": monto_mes})
                continue
            cmzr, cltr = _col(dfr, "mz"), _col(dfr, "lote", "lt")
            ctr, cfr = _col(dfr, "tipo de transaccion"), _col(dfr, "fecha de operacion")
            subr = dfr[(dfr[cmzr].astype(str).str.strip() == mz_hist) & (dfr[cltr].astype(str).str.strip() == lt_hist)]
            subr = subr[subr[ctr].astype(str).str.upper().str.startswith("TE PAG", na=False)]
            fecha = str(subr.iloc[0][cfr]) if not subr.empty else "sin match en Reporte"
            out.append({"MES": mes, "MEDIO": "YAPE", "FECHA_HORA": fecha, "MONTO": monto_mes})
        else:
            out.append({"MES": mes, "MEDIO": "EFECTIVO", "FECHA_HORA": "--", "MONTO": monto_mes})

    # Junio y julio: total real pagado (yape+efectivo) desde el planilla_cobrado
    # de CADA ciclo (junio vive en el repo "jass_system - junio", julio en el
    # activo) vs. lo que la tabla de arriba muestra aplicado a conceptos --
    # desde junio 2026 el exceso se RETIENE (no se aplica ni se muestra solo,
    # ver nota del reporte), asi que puede haber una diferencia real y
    # legitima. Se muestra como linea aparte, nunca mezclada.
    _PLANILLAS_RECIENTES = [
        ("2026-06", rh._planilla_junio()),
        ("2026-07", BASE_DIR.parent / "5_cobranza" / "outputs" / "planilla_cobrado.xlsx"),
    ]
    for mes_ciclo, f in _PLANILLAS_RECIENTES:
        # Los overlays entran a la cascada del ciclo pero NO viven en
        # planilla_cobrado -- van primero para que aparezcan aunque el ciclo
        # no tenga fila de yape/efectivo.
        out.extend(_overlays_de_plata(mz, lt, mes_ciclo))
        if not f.exists():
            _avisar_falta(f, f"referencias de pago de {mes_ciclo}")
            continue
        dfp = pd.read_excel(f, sheet_name="planilla_cobrado", header=1)
        subp = dfp[(dfp["MZ"].astype(str).str.strip() == mz) & (dfp["LT"].astype(str).str.strip() == lt)]
        if subp.empty:
            continue
        r = subp.iloc[0]
        yape = rh._numf(r.get("MONTO_YAPE"))
        efectivo = rh._numf(r.get("MONTO_EFECTIVO"))
        total_real = yape + efectivo
        total_aplicado = 0.0
        if tabla is not None:
            m = tabla[tabla["MES"] == mes_ciclo]
            if not m.empty:
                total_aplicado = float(m.iloc[0]["TOTAL"])
        if yape > 0.005:
            # pagos_yape_tepago.xlsx directo -- ya viene resuelto 1 a 1 por
            # MZ/LOTE, no hace falta adivinar nada. _yape_fecha_hora (cruce
            # de nombre por prefijo contra el banco crudo) se sacó del
            # camino: cuando el origen registrado en maestro_yape.xlsx es
            # generico (ej. J-1: origen = "Janet Villanueva", la tesorera --
            # su nombre aparece en decenas de transacciones ajenas) pegaba
            # fechas de pagos de otros predios (bug encontrado 01/08/2026).
            fecha_yape = _fecha_yape_cruda(mz, lt, mes_ciclo)
            out.append({"MES": mes_ciclo, "MEDIO": "YAPE",
                        "FECHA_HORA": fecha_yape or "sin match en pagos_yape_tepago.xlsx (total del ciclo)",
                        "MONTO": yape})
        if efectivo > 0.005:
            dia, cobrador = _cobrador_efectivo(mz, lt, mes_ciclo)
            detalle = f"{dia} · {cobrador}" if cobrador else "sin detalle de transaccion (total del ciclo)"
            out.append({"MES": mes_ciclo, "MEDIO": "EFECTIVO", "FECHA_HORA": detalle, "MONTO": efectivo})
        # Si total_real y total_aplicado no coinciden (en cualquier sentido),
        # no se afirma de donde viene ni a donde fue la diferencia -- esto
        # solia mostrar una nota especulando el origen (blancos/abonos/
        # condonacion/convenio instalacion) y generaba confusion, porque esa
        # explicacion no esta confirmada (a veces es solo lo que la secretaria
        # declaro despues). Se deja en blanco a proposito: se muestra
        # unicamente el pago real encontrado, sin inventar el resto.

    return out


_ABONOS_REZAGADOS = BASE_DIR.parent / "shared" / "abonos_rezagados.xlsx"
_BLANCOS_EFECTIVO = BASE_DIR.parent / "shared" / "blancos_efectivo.xlsx"
_cache_overlays: dict[str, pd.DataFrame] = {}


def _cargar_overlay(p: Path) -> pd.DataFrame:
    if str(p) not in _cache_overlays:
        if not p.exists():
            _avisar_falta(p, "overlay de pagos")
            _cache_overlays[str(p)] = pd.DataFrame()
        else:
            df = pd.read_excel(p, header=1)
            df.columns = [str(c).strip().upper() for c in df.columns]
            _cache_overlays[str(p)] = df
    return _cache_overlays[str(p)]


def _overlays_de_plata(mz: str, lt: str, mes_ciclo: str) -> list[dict]:
    """Plata real que entra a la cascada del ciclo pero NO esta en
    MONTO_YAPE/MONTO_EFECTIVO de planilla_cobrado: abonos que el cobrador
    retuvo y se regularizaron despues, y efectivo que habia entrado como
    blanco. Sin estas lineas el pago que salda la deuda queda sin origen
    visible y se lee como pago fantasma (caso encontrado 05/08/2026: los 4
    abonos retenidos por Wagner en julio -- S-5, D-16, D1-6, L-4)."""
    out = []
    for p, medio, quien_cols in ((_ABONOS_REZAGADOS, "ABONO REZ.", ("RETENIDO_POR", "CANAL_ORIGEN")),
                                  (_BLANCOS_EFECTIVO, "BLANCO EF.", ("ORIGEN", "CANAL"))):
        df = _cargar_overlay(p)
        if df.empty or "MES_ANO_APLICA" not in df.columns:
            continue
        sub = df[(df["MZ"].astype(str).str.strip() == mz) &
                 (df["LT"].apply(_norm_lote) == _norm_lote(lt)) &
                 (df["MES_ANO_APLICA"].astype(str).str.strip() == mes_ciclo)]
        for _, r in sub.iterrows():
            quien = " · ".join(str(r[c]).strip() for c in quien_cols
                               if c in df.columns and pd.notna(r.get(c)))
            fecha = str(r.get("FECHA_REAL", "")).strip()
            out.append({"MES": mes_ciclo, "MEDIO": medio,
                        "FECHA_HORA": f"{fecha} · del ciclo {r.get('MES_CICLO')} · {quien}",
                        "MONTO": float(r["MONTO"])})
    return out


_PAGOS_EFECTIVO_CRUDO = {
    "2026-06": ciclo.resolver(rh.REPO_JUNIO / "4_pagos" / "efectivo" / "outputs",
                              "pagos_efectivo", "2026-06"),
    # legacy_sin_periodo: 4_pagos del repo activo todavía escribe sin periodo
    # (migra en la tanda B); junio ya está renombrado y no lo necesita.
    "2026-07": ciclo.resolver(BASE_DIR.parent / "4_pagos" / "efectivo" / "outputs",
                              "pagos_efectivo", "2026-07", legacy_sin_periodo=True),
}
_cache_pagos_efectivo: dict[str, pd.DataFrame] = {}


def _cargar_pagos_efectivo_crudo(mes_ciclo: str) -> pd.DataFrame | None:
    """FECHA real por transaccion -- trazabilidad_{mes}.xlsx (hoja
    solo_un_cobrador) tiene FECHA_COBRO vacio para TODO julio (bug de
    construccion, 221/221 filas verificado 01/08/2026) aunque el archivo
    crudo que la alimenta si trae fecha. Se usa como fuente de la fecha en
    vez de trazabilidad; trazabilidad sigue siendo la fuente del COBRADOR
    resuelto (para pagos divididos entre mesas)."""
    if mes_ciclo not in _cache_pagos_efectivo:
        p = _PAGOS_EFECTIVO_CRUDO.get(mes_ciclo)
        if p is None or not p.exists():
            if p is not None:
                _avisar_falta(p, f"pagos_efectivo de {mes_ciclo}")
            _cache_pagos_efectivo[mes_ciclo] = pd.DataFrame()
        else:
            df = pd.read_excel(p, header=1)
            df.columns = [str(c).strip() for c in df.columns]
            _cache_pagos_efectivo[mes_ciclo] = df
    return _cache_pagos_efectivo[mes_ciclo]


def _fecha_efectivo_cruda(mz: str, lt: str, mes_ciclo: str, cobrador: str) -> str | None:
    df = _cargar_pagos_efectivo_crudo(mes_ciclo)
    if df is None or df.empty or "FECHA" not in df.columns:
        return None
    sub = df[(df["MZ"].astype(str).str.strip() == mz) & (df["LT"].astype(str).str.strip() == str(lt))]
    if sub.empty:
        return None
    if cobrador and "COBRADOR" in sub.columns:
        con_cobrador = sub[sub["COBRADOR"].astype(str).str.strip() == cobrador]
        if not con_cobrador.empty:
            sub = con_cobrador
    fechas = sub["FECHA"].dropna()
    if fechas.empty:
        return None
    return " · ".join(sorted({str(f) for f in fechas}))


def _cobrador_fecha_efectivo_crudo(mz: str, lt: str, mes_ciclo: str) -> tuple[str, str] | None:
    """Respaldo cuando el predio NO aparece en trazabilidad_{mes}.xlsx en
    absoluto (135 de 357 casos de julio, verificado 01/08/2026 -- gap
    distinto al de FECHA_COBRO vacio: acá falta la fila entera). Se toma
    directo de pagos_efectivo.xlsx filtrando ESTADO=solo_un_cobrador (mismo
    criterio que usaría trazabilidad si el predio estuviera ahí)."""
    df = _cargar_pagos_efectivo_crudo(mes_ciclo)
    if df is None or df.empty or "ESTADO" not in df.columns:
        return None
    sub = df[(df["MZ"].astype(str).str.strip() == mz) & (df["LT"].astype(str).str.strip() == str(lt)) &
              (df["ESTADO"] == "solo_un_cobrador")]
    if sub.empty:
        return None
    r = sub.iloc[0]
    fecha = r.get("FECHA")
    dia = str(fecha) if pd.notna(fecha) else "(fecha no registrada)"
    cobrador = str(r.get("COBRADOR", "")).strip()
    return dia, cobrador


_PAGOS_YAPE_CRUDO = {
    "2026-06": ciclo.resolver(rh.REPO_JUNIO / "4_pagos" / "yape" / "motor_matching" / "outputs",
                              "pagos_yape_tepago", "2026-06"),
    "2026-07": ciclo.resolver(BASE_DIR.parent / "4_pagos" / "yape" / "motor_matching" / "outputs",
                              "pagos_yape_tepago", "2026-07", legacy_sin_periodo=True),
}
_cache_pagos_yape_crudo: dict[str, pd.DataFrame] = {}


def _cargar_pagos_yape_crudo(mes_ciclo: str) -> pd.DataFrame | None:
    """FECHA real por transaccion, directo de pagos_yape_tepago.xlsx (lo que
    alimenta maestro_yape.xlsx) -- respaldo cuando _yape_fecha_hora() no
    encuentra match cruzando maestro_yape contra el banco crudo (cruce por
    nombre truncado + fecha, se cae seguido). Un paso mas atras que ese
    cruce: esta fecha ya viene resuelta contra MZ/LOTE, sin necesidad de
    adivinar el origen."""
    if mes_ciclo not in _cache_pagos_yape_crudo:
        p = _PAGOS_YAPE_CRUDO.get(mes_ciclo)
        if p is None or not p.exists():
            if p is not None:
                _avisar_falta(p, f"pagos_yape_tepago de {mes_ciclo}")
            _cache_pagos_yape_crudo[mes_ciclo] = pd.DataFrame()
        else:
            df = pd.read_excel(p, header=1)
            df.columns = [str(c).strip() for c in df.columns]
            _cache_pagos_yape_crudo[mes_ciclo] = df
    return _cache_pagos_yape_crudo[mes_ciclo]


def _norm_lote(v) -> str:
    """pagos_yape_tepago.xlsx guarda LOTE como numero (4.0) para lotes sin
    letra -- comparar como texto plano ("4") sin esto nunca matchea, aunque
    la fila SI exista (bug encontrado 01/08/2026, caso I-4: el dato estaba,
    la comparacion de texto lo escondia)."""
    s = str(v).strip()
    if s.endswith(".0"):
        s = s[:-2]
    return s


def _fecha_yape_cruda(mz: str, lt: str, mes_ciclo: str) -> str | None:
    df = _cargar_pagos_yape_crudo(mes_ciclo)
    if df is None or df.empty or "FECHA" not in df.columns or "LOTE" not in df.columns:
        return None
    sub = df[(df["MZ"].astype(str).str.strip() == mz) & (df["LOTE"].apply(_norm_lote) == _norm_lote(lt))]
    fechas = sub["FECHA"].dropna()
    if fechas.empty:
        return None
    return " · ".join(sorted({str(f) for f in fechas}))


def _cobrador_efectivo(mz: str, lt: str, mes_ciclo: str) -> tuple[str, str]:
    """Dia de cobro y cobrador para un pago en efectivo de junio/julio 2026.
    Fuente primaria: 4_pagos/efectivo/trazabilidad/trazabilidad_{mes}.xlsx
    (hoja solo_un_cobrador). Dos gaps verificados en julio (01/08/2026), los
    dos con respaldo en pagos_efectivo.xlsx (el crudo que alimenta la
    trazabilidad):
      1. El predio SI esta en trazabilidad pero FECHA_COBRO vacio (221/221
         filas de julio) -> _fecha_efectivo_cruda().
      2. El predio NO esta en trazabilidad en absoluto (135/357 de julio,
         gap distinto) -> _cobrador_fecha_efectivo_crudo().
    Un solo fix acá cubre cualquier reporte que llame referencias_pago()."""
    f = BASE_DIR.parent / "4_pagos" / "efectivo" / "trazabilidad" / f"trazabilidad_{mes_ciclo}.xlsx"
    df = None
    if f.exists():
        try:
            df = pd.read_excel(f, sheet_name="solo_un_cobrador", header=1)
        except Exception:
            df = None

    sub = None
    if df is not None:
        sub = df[(df["MZ"].astype(str).str.strip() == mz) & (df["LT"].astype(str).str.strip() == str(lt))]

    if sub is None or sub.empty:
        crudo = _cobrador_fecha_efectivo_crudo(mz, lt, mes_ciclo)
        return crudo if crudo else ("", "")

    r = sub.iloc[0]
    cobrador = str(r.get("COBRADOR", "")).strip()
    fecha = r.get("FECHA_COBRO")
    if pd.notna(fecha):
        dia = str(fecha)
    else:
        dia = _fecha_efectivo_cruda(mz, lt, mes_ciclo, cobrador) or "(fecha no registrada)"
    return dia, cobrador


def _dibujar_tabla_referencias(page, x: float, y: float, w: float, refs: list[dict]) -> float:
    headers = ["Mes", "Medio", "Fecha / hora", "Monto"]
    anchos = [70, 80, w - 70 - 80 - 110, 110]
    rh_row = 16

    page.insert_text((x, y), "Referencia de pago (de donde vino cada pago)", fontsize=9, fontname="hebo", color=_AZUL)
    y += 16

    page.draw_rect(fitz.Rect(x, y, x + w, y + rh_row), fill=_AZUL_BG, color=None)
    cx = x
    for h, cw in zip(headers, anchos):
        page.insert_text((cx + 4, y + rh_row - 5), h, fontsize=8, fontname="hebo", color=_AZUL)
        cx += cw
    y += rh_row

    for n, r in enumerate(refs):
        if n % 2 == 1:
            page.draw_rect(fitz.Rect(x, y, x + w, y + rh_row), fill=_ZEBRA, color=None)
        cx = x
        page.insert_text((cx + 4, y + rh_row - 5), str(r["MES"]), fontsize=8, fontname="hebo", color=_NEGRO)
        cx += anchos[0]
        es_nota = r["MEDIO"] == "(nota)"
        color_medio = _GRIS if es_nota else (_VERDE if r["MEDIO"] == "YAPE" else _GRIS)
        page.insert_text((cx + 4, y + rh_row - 5), r["MEDIO"], fontsize=8, fontname="hebo", color=color_medio)
        cx += anchos[1]
        page.insert_text((cx + 4, y + rh_row - 5), str(r["FECHA_HORA"]), fontsize=7.5 if es_nota else 8,
                          fontname="helv", color=_GRIS if es_nota else _NEGRO)
        cx += anchos[2]
        if not es_nota:
            texto_monto = f"S/ {r['MONTO']:,.2f}"
            tw = fitz.get_text_length(texto_monto, fontname="hebo", fontsize=8)
            page.insert_text((cx + anchos[-1] - 6 - tw, y + rh_row - 5), texto_monto, fontsize=8, fontname="hebo", color=_NEGRO)
        y += rh_row

    return y


def verificar_predio(mz: str, lt: str, tabla: pd.DataFrame | None = None, refs: list[dict] | None = None) -> list[dict]:
    """Compara, mes a mes, el TOTAL de la tabla de historial contra la suma de
    referencias de pago de ese mes. Devuelve solo los meses que NO cuadran
    (lista vacia = todo cuadra)."""
    if tabla is None:
        tabla = rh.tabla_predio(mz, lt)
        tabla = rcm.corregir_tabla_por_redirects(tabla, mz, lt, rcm._cargar_redirects())
    if refs is None:
        refs = referencias_pago(mz, lt)

    ref_por_mes: dict[str, float] = {}
    for r in refs:
        ref_por_mes[r["MES"]] = ref_por_mes.get(r["MES"], 0.0) + r["MONTO"]

    _MESES_RECIENTES = {"2026-06", "2026-07"}

    problemas = []
    for _, row in tabla.iterrows():
        mes = row["MES"]
        total_tabla = float(row["TOTAL"])
        total_ref = ref_por_mes.get(mes, 0.0)
        # jun-jul: no se afirma el origen de una diferencia (ver
        # referencias_pago) -- un descuadre en cualquier sentido ahi es
        # esperado, no se marca como "problema" del reporte.
        if mes in _MESES_RECIENTES:
            continue
        if abs(total_tabla - total_ref) > 0.5:
            problemas.append({"MES": mes, "TOTAL_TABLA": total_tabla, "TOTAL_REFERENCIA": total_ref,
                               "DIFERENCIA": round(total_tabla - total_ref, 2)})
    return problemas


def generar_pdf(mz: str, lt: str, salida: Path | None = None) -> Path:
    tabla = rh.tabla_predio(mz, lt)
    tabla = rcm.corregir_tabla_por_redirects(tabla, mz, lt, rcm._cargar_redirects())
    refs = referencias_pago(mz, lt, tabla=tabla)
    nombres = repo._lookup_nombres()

    doc = fitz.open()
    rh._dibujar_pagina(doc, mz, lt, nombres.get((mz, lt), ""), tabla)
    page = doc[-1]
    w = rh._PAGE_W - 2 * rh._M
    y_tabla_fin = rh._M + 42 + 14 + 18 + (18 * len(tabla))
    y = y_tabla_fin + 45
    _dibujar_tabla_referencias(page, rh._M, y, w, refs)

    salida = salida or (BASE_DIR / "outputs" / f"reporte_referencias_{mz}-{lt}.pdf")
    salida.parent.mkdir(exist_ok=True)
    doc.save(str(salida))
    doc.close()
    print(f"PDF -> {salida} ({len(refs)} referencias de pago)")
    return salida


def generar_lote(mes_ano: str = "2026-07", salida: Path | None = None) -> Path:
    """1 PDF: portada resumen (reusa reporte_convenio_multa) + 1 pagina por
    predio con historial + referencia de pago, para los predios con CONVENIO
    pendiente."""
    df = rcm.calcular_tabla(mes_ano)
    redirects = rcm._cargar_redirects()

    historicos = rh._cargar_historicos()
    eventos = repo._leer_eventos()
    mapa_raw = rh._cargar_mapa_raw()
    f = BASE_DIR.parent / "5_cobranza" / "outputs" / "planilla_cobrado.xlsx"
    dfp = pd.read_excel(f, sheet_name="planilla_cobrado", header=1)
    nombres = repo._lookup_nombres()

    doc = fitz.open()
    rcm._dibujar_portada(doc, df, mes_ano)
    for _, r in df.iterrows():
        mz, lt = r["MZ"], r["LT"]
        tabla = rh.tabla_predio(mz, lt, historicos, eventos, dfp, mapa_raw, nombres.get((mz, lt), ""))
        tabla = rcm.corregir_tabla_por_redirects(tabla, mz, lt, redirects)
        refs = referencias_pago(mz, lt, tabla=tabla)
        rh._dibujar_pagina(doc, mz, lt, nombres.get((mz, lt), ""), tabla)
        page = doc[-1]
        w = rh._PAGE_W - 2 * rh._M
        y_tabla_fin = rh._M + 42 + 14 + 18 + (18 * len(tabla))
        _dibujar_tabla_referencias(page, rh._M, y_tabla_fin + 45, w, refs)

    salida = salida or (BASE_DIR / "outputs" / f"reporte_convenio_multa_referencias_{mes_ano}.pdf")
    salida.parent.mkdir(exist_ok=True)
    doc.save(str(salida))
    doc.close()
    print(f"{len(df)} predios -> PDF {salida}")
    return salida


if __name__ == "__main__":
    generar_pdf("D", "1")
