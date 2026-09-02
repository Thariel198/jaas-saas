"""4b_reclamos/herramienta/auditar_pago_vs_ledger.py — detecta deuda escondida
por el bug de signo (ya corregido en el código, NO en los datos viejos) de
`5_cobranza/main.py::_reconciliar_pagos_pueblo()`, y por qué el PAGO_TOTAL del
historial a veces no cuadra contra la plata real trazada (referencias_pago).

HALLAZGO PRINCIPAL (validado contra el código y su historial git — sin
ambigüedad): la reconciliación es INCREMENTAL — cada corrida de 5_cobranza
recalcula la cascada completa de un predio y compara contra lo YA registrado
en el ledger; si el cálculo fresco da MENOS que antes para un concepto, tiene
que ESCRIBIR DE VUELTA la deuda que ese pago ya no cubre (un AJUSTE positivo,
el saldo debe SUBIR). Hasta el commit `bda176d` (12/08/2026 15:57:27), el
código llamaba `registrar_ajuste(..., delta, ...)` con `delta` crudo (negativo)
en vez de `-delta` — el saldo, en lugar de subir, volvía a bajar: la deuda se
perdonaba el DOBLE de lo que correspondía. El commit lo corrigió (nombra a
D-16 y D1-6 como los casos que lo destaparon) pero el ledger es append-only:
los eventos ya escritos con el signo malo se quedaron así. Barrido completo
del ledger (ver `deuda_escondida_ledger()`): 48 eventos, jul-2026, S/4,106 de
deuda real que hoy NO aparece en ningún reporte porque el propio ledger la
tiene registrada como perdonada dos veces en vez de restaurada una vez.

Caso que lo destapó: D-16, julio 2026 — ACUERDOS PAGO 25 (20/07) seguido de
ACUERDOS AJUSTE -25 (31/07, CLASE=CORRECCION_SISTEMA, source=5_cobranza,
ANTES del fix) → SALDO quedó en 0. Con el signo correcto habría quedado en
50 (25 que ya tenía + 25 que el ajuste debía devolver). Ese predio debe S/25
más de ACUERDOS de lo que el ledger admite hoy — y nada posterior lo corrigió.

NO TODOS los 48 eventos con el signo malo siguen rotos hoy — dos matices
confirmados contra los datos crudos (no son hipótesis):

  (a) 31/48 YA SE AUTOCORRIGIERON. Una corrida posterior (5_cobranza de nuevo,
      ya con el código arreglado, o una corrección manual) volvió a tocar el
      mismo (predio, concepto) y restauró el saldo. Ejemplo D-1/MULTA:
      CARGO 30 → PAGO 30 (06/07) → AJUSTE -30 (31/07 18:11, sign-bug) →
      AJUSTE +30 (31/07 18:29, source=manual, audit_ref="notas_2026-07|
      D-1-MULTA-estabilizador-31072026") → SALDO final 0, correcto.

  (b) La causa que el usuario sospechaba SÍ existe: reidentificación
      documentada. `shared/reasignaciones_aplicacion.xlsx` redirige a
      propósito un pago de un concepto a otro (ej. D-1: MULTA→CONVENIO,
      MOTIVO dice literal "el motor va a generar un AJUSTE automático
      negativo en MULTA — ESPERADO, no es un bug — se estabiliza después con
      un AJUSTE manual"). Para estos, el AJUSTE negativo es INTENCIONAL; lo
      único a verificar es que el estabilizador se haya escrito — que es
      exactamente el caso (a).

  Quedan 17/48 SIN ningún ajuste posterior que los toque — esos son la deuda
  escondida ACTIVA hoy (`deuda_escondida_ledger()` los separa de los 31 ya
  resueltos). No se pudo aislar con certeza un tercer patrón distinto
  ("el pago existía y se eliminó") — no hay snapshots históricos de
  pagos_efectivo/pagos_yape_tepago para probarlo; queda como limitación
  conocida, no se inventa una detección sin poder validarla.

Efecto secundario en el REPORTE (más chico, aparte): `_filas_recientes()` en
comun.py suma TODOS los PAGO de un concepto/mes para armar PAGO_TOTAL sin
netear contra un AJUSTE "recalculado a la baja" del mismo concepto/mes — el
PAGO original (25 en el ejemplo) sigue contando aunque el propio ledger ya lo
haya tocado. `explicar_predio()` sigue detectando esto para contexto, pero
`deuda_escondida_ledger()` es el hallazgo que importa.

Uso:
  py auditar_pago_vs_ledger.py --deuda-escondida        (el barrido completo)
  py auditar_pago_vs_ledger.py MZ LT
  py auditar_pago_vs_ledger.py --todos [MES_ANO]
  py auditar_pago_vs_ledger.py --archivo shared/abonos_rezagados.xlsx
"""

import sys
from pathlib import Path

import pandas as pd

HERRAMIENTA_DIR = Path(__file__).parent
sys.path.insert(0, str(HERRAMIENTA_DIR))
import comun  # noqa: E402

TOL_DIFERENCIA = 0.5  # mismo umbral que usa comun.verificar_predio()

# CLASE="CORRECCION_SISTEMA" la escribe UN SOLO lugar en todo el repo:
# 5_cobranza/main.py::_reconciliar_pagos_pueblo() linea ~2512, siempre con
# TIPO_EVENTO=AJUSTE y SOURCE="5_cobranza" -- firma determinista del patron
# "recalculado a la baja", sin ambiguedad. No se puede filtrar por MOTIVO:
# los AJUSTE escritos antes del fix de registrar_ajuste (03/08/2026, commit
# 134a4e7) tienen MOTIVO vacio aunque el codigo llamante SI lo mandaba -- el
# caso D-16/julio es uno de esos (verificado: MOTIVO=NaN, CLASE=CORRECCION_SISTEMA).
_CLASE_RECALCULO = "CORRECCION_SISTEMA"
_SOURCE_RECALCULO = "5_cobranza"

# Segunda causa confirmada (caso D-16, el resto de la diferencia que el
# patron de arriba no explicaba): shared/planilla_mes/planilla_<mes>.xlsx
# (el CARGO real, ver comun._cargar_planilla_correcta) puede corregirse
# DESPUES de que 5_cobranza ya corrio y reconcilio ese mes contra la foto
# vieja de planilla_cobrado.xlsx -- la reconciliacion cascadea con el CARGO
# de ESE momento, no con el corregido despues. Mismo patron que D1-6/S-5
# (ver LEER_ANTES.md). No se puede sumar 1 a 1 contra la diferencia (depende
# de en que escalon de la cascada caiga el faltante) -- se reporta como
# evidencia de contexto, no como monto que neteal la diferencia.
_CAMPOS_AGUA = ("MES_ACTUAL", "MANTENIMIENTO", "MES_ANTERIOR", "CORTE_RECONEXION")

# Commit bda176d (12/08/2026 15:57:27) corrigió `registrar_ajuste(..., delta, ...)`
# -> `registrar_ajuste(..., -delta, ...)`. Todo AJUSTE CORRECCION_SISTEMA de
# 5_cobranza escrito ANTES de este instante tiene el signo invertido.
_FIX_SIGNO_AJUSTE = pd.Timestamp("2026-08-12 15:57:27")

_reasignaciones_cache: pd.DataFrame | None = None


def _reasignaciones_aplicacion() -> pd.DataFrame:
    """shared/reasignaciones_aplicacion.xlsx — redirecciones documentadas de un
    concepto a otro (ej. MULTA->CONVENIO). Cuando el predio/concepto de un
    AJUSTE con signo malo aparece acá como CONCEPTO_ORIGEN, el AJUSTE negativo
    era intencional (ver docstring del módulo, caso D-1) -- no un accidente."""
    global _reasignaciones_cache
    if _reasignaciones_cache is None:
        ruta = comun.SHARED_DIR / "reasignaciones_aplicacion.xlsx"
        if ruta.exists():
            df = pd.read_excel(ruta, header=1)
            df.columns = [str(c).strip().upper() for c in df.columns]
            _reasignaciones_cache = df
        else:
            _reasignaciones_cache = pd.DataFrame()
    return _reasignaciones_cache


def deuda_escondida_ledger(eventos: pd.DataFrame | None = None) -> pd.DataFrame:
    """Barre TODO el ledger (no un predio a la vez) buscando AJUSTE
    CORRECCION_SISTEMA de 5_cobranza escritos antes del fix de signo
    (ver docstring del módulo) — cada uno de estos debía SUMAR al saldo
    (devolver deuda) y en cambio restó (perdonó el doble). La columna
    DEUDA_ESCONDIDA = 2 * |AJUSTE| es la deuda que ese evento, aislado, dejó
    sin admitir.

    Clasifica cada uno en vez de tratarlos a todos como igual de urgentes:
      TIENE_REDIRECT_ASOCIADO  hay una reasignación documentada para este
                                predio/concepto (el ajuste negativo pudo ser
                                intencional, no un accidente)
      SUMA_AJUSTE_POSTERIOR    Σ AJUSTE (cualquier source/clase) para el mismo
                                predio+concepto DESPUÉS de este evento
      YA_ESTABILIZADO          esa suma posterior ya cubre el monto que este
                                evento dejó sin admitir — alguien (5_cobranza
                                en una corrida posterior, o a mano) ya lo
                                corrigió; no es deuda activa hoy

    Devuelve TODOS los eventos con el signo malo (activos y ya estabilizados
    mezclados) — usar el filtro `YA_ESTABILIZADO` para separar lo urgente de
    lo que ya es solo historia. `auditar_deuda_escondida()` hace ese split."""
    eventos = eventos if eventos is not None else comun.repo._leer_eventos()
    ev = eventos.copy()
    ev["TIMESTAMP"] = pd.to_datetime(ev["TIMESTAMP"], errors="coerce")
    afectados = ev[
        (ev["TIPO_EVENTO"] == "AJUSTE") &
        (ev["CLASE"].astype(str).str.strip() == _CLASE_RECALCULO) &
        (ev["SOURCE"].astype(str).str.strip() == _SOURCE_RECALCULO) &
        (ev["TIMESTAMP"] < _FIX_SIGNO_AJUSTE) &
        (pd.to_numeric(ev["AJUSTE"], errors="coerce") < 0)
    ].copy()
    if afectados.empty:
        return afectados
    afectados["AJUSTE_MAL_APLICADO"] = pd.to_numeric(afectados["AJUSTE"], errors="coerce")
    afectados["DEUDA_ESCONDIDA"] = afectados["AJUSTE_MAL_APLICADO"].abs() * 2
    afectados["SALDO_ACTUAL_LEDGER"] = pd.to_numeric(afectados["SALDO"], errors="coerce")
    afectados["SALDO_CORREGIDO"] = afectados["SALDO_ACTUAL_LEDGER"] + afectados["DEUDA_ESCONDIDA"]

    redirects = _reasignaciones_aplicacion()
    ev_ajuste = ev[ev["TIPO_EVENTO"] == "AJUSTE"]

    tiene_redirect, suma_posterior, ya_estabilizado = [], [], []
    for _, r in afectados.iterrows():
        mz, lt, concepto, ts = r["MZ"], r["LT"], r["CONCEPTO"], r["TIMESTAMP"]
        if redirects.empty:
            tiene_redirect.append(False)
        else:
            match = redirects[
                (redirects["MZ"].astype(str).str.strip() == str(mz).strip()) &
                (redirects["LT"].astype(str).str.strip() == str(lt).strip()) &
                (redirects["CONCEPTO_ORIGEN"].astype(str).str.strip().str.upper() == str(concepto).strip().upper())
            ]
            tiene_redirect.append(not match.empty)
        posteriores = ev_ajuste[
            (ev_ajuste["MZ"] == mz) & (ev_ajuste["LT"] == lt) &
            (ev_ajuste["CONCEPTO"] == concepto) & (ev_ajuste["TIMESTAMP"] > ts)
        ]
        suma = pd.to_numeric(posteriores["AJUSTE"], errors="coerce").sum()
        suma_posterior.append(suma)
        ya_estabilizado.append(suma >= abs(float(r["AJUSTE_MAL_APLICADO"])) - TOL_DIFERENCIA)

    afectados["TIENE_REDIRECT_ASOCIADO"] = tiene_redirect
    afectados["SUMA_AJUSTE_POSTERIOR"] = suma_posterior
    afectados["YA_ESTABILIZADO"] = ya_estabilizado

    cols = ["MZ", "LT", "CONCEPTO", "MES", "AJUSTE_MAL_APLICADO", "DEUDA_ESCONDIDA",
            "SALDO_ACTUAL_LEDGER", "SALDO_CORREGIDO", "TIENE_REDIRECT_ASOCIADO",
            "SUMA_AJUSTE_POSTERIOR", "YA_ESTABILIZADO", "TIMESTAMP", "AUDIT_REF"]
    return afectados[cols].sort_values(["YA_ESTABILIZADO", "DEUDA_ESCONDIDA"],
                                       ascending=[True, False]).reset_index(drop=True)


def auditar_deuda_escondida(eventos: pd.DataFrame | None = None) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Separa deuda_escondida_ledger() en (activa, ya_resuelta) — solo la
    primera es un problema hoy."""
    todos = deuda_escondida_ledger(eventos)
    if todos.empty:
        return todos, todos
    activa = todos[~todos["YA_ESTABILIZADO"]].reset_index(drop=True)
    resuelta = todos[todos["YA_ESTABILIZADO"]].reset_index(drop=True)
    return activa, resuelta


def _eventos_predio(mz: str, lt: str, eventos: pd.DataFrame) -> pd.DataFrame:
    ev = eventos[(eventos["MZ"].astype(str).str.strip() == mz) &
                 (eventos["LT"].astype(str).str.strip() == lt)].copy()
    ev["TIMESTAMP"] = pd.to_datetime(ev["TIMESTAMP"], errors="coerce")
    ev["MES"] = ev["MES"].astype(str).str.strip()
    ev["CONCEPTO"] = ev["CONCEPTO"].astype(str).str.strip().str.upper()
    return ev.sort_values("TIMESTAMP")


def _dfp_para_mes(mes_ano: str) -> pd.DataFrame | None:
    """planilla_cobrado.xlsx del ciclo que corresponde a mes_ano -- cerrado
    (repo congelado) o el activo, mismo criterio que usa comun._ciclos_recientes()."""
    if mes_ano in comun.REPOS_CICLO_CERRADO:
        df = comun._cargar_dfp_ciclo_cerrado(mes_ano)
        return df if df is not None and not df.empty else None
    for m, ruta in comun._ciclos_recientes():
        if m == mes_ano and ruta.exists():
            return pd.read_excel(ruta, sheet_name="planilla_cobrado", header=1)
    return None


def _diferencias_planilla(mz: str, lt: str, mes_ano: str) -> list[dict]:
    """Campos de agua (MES_ACTUAL/MANTENIMIENTO/MES_ANTERIOR/CORTE_RECONEXION)
    donde shared/planilla_mes (el CARGO corregido, lo que usa el historial)
    no coincide con lo que planilla_cobrado.xlsx tenía anotado para ese
    predio -- señal de que 2_planilla se corrigió DESPUÉS de que 5_cobranza
    reconcilió ese mes contra el ledger (ver constante _CAMPOS_AGUA)."""
    correcta = comun._cargar_planilla_correcta(mes_ano)
    dfp = _dfp_para_mes(mes_ano)
    if correcta is None or dfp is None:
        return []
    fila_correcta = correcta[(correcta["MZ"] == mz) & (correcta["LT"] == lt)]
    fila_cobrado = dfp[(dfp["MZ"].astype(str).str.strip() == mz) & (dfp["LT"].astype(str).str.strip() == lt)]
    if fila_correcta.empty or fila_cobrado.empty:
        return []
    rc, rb = fila_correcta.iloc[0], fila_cobrado.iloc[0]
    out = []
    for campo in _CAMPOS_AGUA:
        v_correcto = comun._numf(rc.get(campo))
        v_cobrado = comun._numf(rb.get(campo))
        if abs(v_correcto - v_cobrado) > TOL_DIFERENCIA:
            out.append({"CAMPO": campo, "PLANILLA_MES_CORREGIDA": v_correcto,
                        "PLANILLA_COBRADO_AL_CORRER_5_COBRANZA": v_cobrado,
                        "DIFERENCIA": round(v_correcto - v_cobrado, 2)})
    return out


def explicar_predio(mz: str, lt: str, eventos: pd.DataFrame | None = None) -> list[dict]:
    """Por mes, compara PAGO_TOTAL del historial (ledger) contra la plata real
    trazada (referencias_pago). Donde no cuadra, busca en el ledger crudo la
    causa más común: un PAGO de un concepto pueblo (CONVENIO/MULTA/ACUERDOS)
    que una corrida posterior de 5_cobranza "recalculó a la baja" vía AJUSTE
    (ver docstring del módulo) sin que el historial neteara esa corrección
    contra el PAGO original — el reporte sigue sumando el PAGO bruto.

    Devuelve 1 fila por (mes, predio) con diferencia real, con el desglose de
    cada corrección encontrada en "DETALLE" (lista vacía = el patrón conocido
    no explica la diferencia; hace falta mirar el caso a mano)."""
    tabla = comun.tabla_predio(mz, lt, deuda_conceptos_desde_ledger=True, incluir_abonos_rezagados=False)
    refs = comun.referencias_pago(mz, lt, tabla=tabla, incluir_overlays=False)
    real_por_mes: dict[str, float] = {}
    for r in refs:
        real_por_mes[r["MES"]] = real_por_mes.get(r["MES"], 0.0) + r["MONTO"]

    eventos = eventos if eventos is not None else comun.repo._leer_eventos()
    ev_predio = _eventos_predio(mz, lt, eventos)
    correcciones = ev_predio[
        (ev_predio["TIPO_EVENTO"] == "AJUSTE") &
        (ev_predio["CLASE"].astype(str).str.strip() == _CLASE_RECALCULO) &
        (ev_predio["SOURCE"].astype(str).str.strip() == _SOURCE_RECALCULO) &
        (pd.to_numeric(ev_predio["AJUSTE"], errors="coerce") < 0)
    ]

    salida = []
    for _, row in tabla.iterrows():
        mes = str(row["MES"]).strip()
        pago_reporte = float(row["PAGO_TOTAL"])
        pago_real = real_por_mes.get(mes, 0.0)
        diferencia = round(pago_reporte - pago_real, 2)
        if abs(diferencia) <= TOL_DIFERENCIA:
            continue

        detalle = []
        explicado = 0.0
        for _, c in correcciones[correcciones["MES"] == mes].iterrows():
            concepto = c["CONCEPTO"]
            monto_ajuste = abs(float(c["AJUSTE"]))
            # El PAGO que este ajuste corrigió: el PAGO más reciente de ese
            # mismo concepto/mes, registrado ANTES de este ajuste en el tiempo.
            previos = ev_predio[
                (ev_predio["CONCEPTO"] == concepto) & (ev_predio["MES"] == mes) &
                (ev_predio["TIPO_EVENTO"] == "PAGO") & (ev_predio["TIMESTAMP"] < c["TIMESTAMP"])
            ]
            pago_original = float(previos["PAGO"].sum()) if not previos.empty else 0.0
            detalle.append({
                "CONCEPTO": concepto,
                "PAGO_ORIGINAL": pago_original,
                "FECHA_PAGO_ORIGINAL": str(previos.iloc[-1]["TIMESTAMP"]) if not previos.empty else "(no encontrado)",
                "AJUSTE_RECALCULO": monto_ajuste,
                "FECHA_AJUSTE": str(c["TIMESTAMP"]),
                "MOTIVO": str(c.get("MOTIVO", "")) if pd.notna(c.get("MOTIVO"))
                          else "(vacío en el ledger — corrección: pago recalculado a la baja en 5_cobranza)",
                "NETO_REAL_DEL_CONCEPTO": round(pago_original - monto_ajuste, 2),
            })
            explicado += monto_ajuste

        salida.append({
            "MZ": mz, "LT": lt, "MES": mes,
            "PAGO_REPORTE": pago_reporte, "PAGO_REAL_TRAZADO": pago_real,
            "DIFERENCIA": diferencia, "EXPLICADO_POR_RECALCULO": round(explicado, 2),
            "SIN_EXPLICAR": round(diferencia - explicado, 2),
            "DETALLE": detalle,
            "DIFERENCIAS_PLANILLA": _diferencias_planilla(mz, lt, mes),
        })
    return salida


def _predios_de_archivo(ruta: Path) -> list[tuple[str, str]]:
    df = pd.read_excel(ruta, header=1)
    df.columns = [str(c).strip().upper() for c in df.columns]
    sub = df.dropna(subset=["MZ", "LT"])
    return sorted({(str(r["MZ"]).strip(), str(r["LT"]).strip().replace(".0", "")) for _, r in sub.iterrows()})


def auditar(predios: list[tuple[str, str]], salida: Path | None = None) -> Path:
    """Corre explicar_predio() sobre una lista de predios y escribe un Excel
    con 3 hojas: Resumen (1 fila por mes con diferencia), Detalle (1 fila por
    corrección 'recalculado a la baja' encontrada) y Diferencias_Planilla
    (1 fila por campo de agua donde planilla_mes ya no coincide con la foto
    que usó 5_cobranza al reconciliar ese mes)."""
    eventos = comun.repo._leer_eventos()
    filas_resumen, filas_detalle, filas_planilla = [], [], []
    for mz, lt in predios:
        for caso in explicar_predio(mz, lt, eventos):
            detalle = caso.pop("DETALLE")
            dif_planilla = caso.pop("DIFERENCIAS_PLANILLA")
            filas_resumen.append(caso)
            for d in detalle:
                filas_detalle.append({"MZ": mz, "LT": lt, "MES": caso["MES"], **d})
            for d in dif_planilla:
                filas_planilla.append({"MZ": mz, "LT": lt, "MES": caso["MES"], **d})

    resumen = pd.DataFrame(filas_resumen)
    detalle_df = pd.DataFrame(filas_detalle)
    planilla_df = pd.DataFrame(filas_planilla)
    print(f"{len(predios)} predio(s) revisados · {len(resumen)} mes(es) con diferencia > S/{TOL_DIFERENCIA:.2f}")
    if not resumen.empty:
        explicados = int((resumen["SIN_EXPLICAR"].abs() <= TOL_DIFERENCIA).sum())
        print(f"  {explicados}/{len(resumen)} totalmente explicados por 'recalculado a la baja'")
        con_dif_planilla = resumen.set_index(["MZ", "LT", "MES"]).index.isin(
            planilla_df.set_index(["MZ", "LT", "MES"]).index) if not planilla_df.empty else []
        print(f"  {sum(con_dif_planilla)}/{len(resumen)} tienen además planilla_mes desactualizada vs. "
              f"lo que usó 5_cobranza (evidencia de contexto, no sumada al explicado)")

    salida = salida or (HERRAMIENTA_DIR / "outputs" / "auditoria_pago_vs_ledger.xlsx")
    salida.parent.mkdir(exist_ok=True)
    with pd.ExcelWriter(salida, engine="openpyxl") as w:
        resumen.to_excel(w, sheet_name="Resumen", startrow=1, index=False)
        detalle_df.to_excel(w, sheet_name="Detalle", startrow=1, index=False)
        planilla_df.to_excel(w, sheet_name="Diferencias_Planilla", startrow=1, index=False)
    print(f"Excel -> {salida}")
    return salida


if __name__ == "__main__":
    argv = sys.argv[1:]
    if argv and argv[0] == "--deuda-escondida":
        activa, resuelta = auditar_deuda_escondida()
        if activa.empty and resuelta.empty:
            print("Sin eventos con el bug de signo — nada que reportar.")
        else:
            pd.set_option("display.max_columns", None)
            pd.set_option("display.width", 300)
            print("=== ACTIVA — sin ningún ajuste posterior, deuda real hoy ===")
            print(activa.to_string(index=False) if not activa.empty else "(ninguna)")
            print(f"\n{len(activa)} eventos · {activa['MZ'].str.cat(activa['LT'], sep='-').nunique() if not activa.empty else 0} "
                  f"predios · S/{activa['DEUDA_ESCONDIDA'].sum() if not activa.empty else 0:,.2f} de deuda escondida ACTIVA")
            print(f"\n({len(resuelta)} eventos más ya se autocorrigieron o tenían un redirect "
                  f"documentado con estabilizador aplicado — ver hoja Ya_Resuelto)")
            salida = HERRAMIENTA_DIR / "outputs" / "deuda_escondida_signo_ajuste.xlsx"
            salida.parent.mkdir(exist_ok=True)
            with pd.ExcelWriter(salida, engine="openpyxl") as w:
                activa.to_excel(w, sheet_name="Deuda_Escondida_Activa", startrow=1, index=False)
                resuelta.to_excel(w, sheet_name="Ya_Resuelto", startrow=1, index=False)
            print(f"Excel -> {salida}")
    elif argv and argv[0] == "--archivo":
        predios = _predios_de_archivo(Path(argv[1]))
        auditar(predios)
    elif argv and argv[0] == "--todos":
        mes_ano = argv[1] if len(argv) > 1 else "2026-07"
        sys.path.insert(0, str(HERRAMIENTA_DIR.parent))
        import reporte_historico as rh  # noqa: E402
        auditar(sorted(rh._universo_predios(mes_ano)))
    elif len(argv) >= 2:
        mz, lt = argv[0], argv[1]
        casos = explicar_predio(mz, lt)
        if not casos:
            print(f"{mz}-{lt}: sin diferencia > S/{TOL_DIFERENCIA:.2f} en ningún mes")
        for caso in casos:
            detalle = caso.pop("DETALLE")
            dif_planilla = caso.pop("DIFERENCIAS_PLANILLA")
            print(f"\n{mz}-{lt} {caso['MES']}: reporte={caso['PAGO_REPORTE']:.2f} "
                  f"real={caso['PAGO_REAL_TRAZADO']:.2f} diferencia={caso['DIFERENCIA']:.2f}")
            for d in detalle:
                print(f"  {d['CONCEPTO']}: PAGO {d['PAGO_ORIGINAL']:.2f} ({d['FECHA_PAGO_ORIGINAL']}) "
                      f"-> AJUSTE -{d['AJUSTE_RECALCULO']:.2f} ({d['FECHA_AJUSTE']}) "
                      f"= neto {d['NETO_REAL_DEL_CONCEPTO']:.2f} · {d['MOTIVO']}")
            if abs(caso["SIN_EXPLICAR"]) > TOL_DIFERENCIA:
                print(f"  SIN EXPLICAR: S/{caso['SIN_EXPLICAR']:.2f}")
                if dif_planilla:
                    print(f"  posible causa — planilla_mes ya no coincide con lo que usó 5_cobranza:")
                    for d in dif_planilla:
                        print(f"    {d['CAMPO']}: planilla_mes={d['PLANILLA_MES_CORREGIDA']:.2f} vs. "
                              f"planilla_cobrado(al correr)={d['PLANILLA_COBRADO_AL_CORRER_5_COBRANZA']:.2f} "
                              f"(diff {d['DIFERENCIA']:+.2f})")
                else:
                    print(f"  revisar a mano — ningún patrón conocido explica esto")
    else:
        print(__doc__)
