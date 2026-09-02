"""Build and run an isolated delayed-payment mini-corrida."""
import json
import shutil
import sys
from datetime import datetime
from pathlib import Path

import pandas as pd

THIS = Path(__file__).resolve()
ROOT = THIS.parents[2]
sys.path.insert(0, str(THIS.parent.parent))
import main as mod  # noqa: E402

OUT = Path(r"C:\Users\wilde\AppData\Local\Temp\opencode\mini_corrida_lista_corte_7_20260816")
INPUTS = OUT / "inputs"
OUTPUTS = OUT / "outputs"
DOC = ROOT / "docs" / "aprendizaje" / "solucion de proble" / "02_visualizacion_por_conjuntos.md"
REFERENCIAS = ROOT / "docs" / "aprendizaje" / "solucion de proble" / "referencias"
LISTA_CORTE = ROOT / "6_corte" / "outputs" / "lista_corte.xlsx"
BACKUPS = OUT.parent / "backups_mini_corrida"
ACTUALIZAR_DOC = "--actualizar-doc" in sys.argv

_CAMPOS_WATERFALL_CICLO_MINI = (
    "mes_actual", "mantenimiento", "mes_anterior", "corte_reconexion",
    "convenio", "acuerdos_asamblea", "multa",
)


def _read(path):
    df = pd.read_excel(path, header=1)
    df.columns = mod._norm_cols(df)
    return df


def _corregir_i9_en_mini(source):
    """Prepare the approved I-9 amount without touching the real source."""
    mask = (
        source["MZ"].map(mod._norm_mz).eq("I") &
        source["LT"].map(mod._norm_lt).eq("9")
    )
    i9 = source[mask]
    if sorted(round(float(v), 2) for v in i9["MONTO"]) != [50.0, 86.0]:
        raise RuntimeError("I-9 mini: se esperaban exactamente los abonos S/86 y S/50")
    source.loc[mask & source["MONTO"].eq(50), "MONTO"] = 58
    source.loc[mask, "MES_ANO_APLICA"] = "2026-08"


def _manifest_mini_i9(keys):
    manifest = json.loads(mod.ABONOS_MANIFEST_PATH.read_text(encoding="utf-8"))
    manifest = [
        row for row in manifest
        if not (mod._norm_mz(row.get("MZ")) == "I" and mod._norm_lt(row.get("LT")) == "9")
    ]
    manifest.extend([
        {"MZ": "I", "LT": "9", "MONTO": 86, "MES_CICLO": "2026-06",
         "MES_ANO_APLICA": "2026-08", "ESTADO": "CONFIRMADO"},
        {"MZ": "I", "LT": "9", "MONTO": 58, "MES_CICLO": "2026-07",
         "MES_ANO_APLICA": "2026-08", "ESTADO": "CONFIRMADO"},
    ])
    return [
        row for row in manifest
        if (mod._norm_mz(row.get("MZ")), mod._norm_lt(row.get("LT"))) in keys
    ]


def _arrastre_cerrado_mini(keys):
    sys.path.insert(0, str(ROOT / "4b_reclamos"))
    import reporte_historico as rh

    root = rh.REPOS_CICLO_CERRADO["2026-07"]
    planilla = root / "2_planilla" / "outputs" / "planilla_2026-07.xlsx"
    if not planilla.exists():
        raise FileNotFoundError(f"Falta planilla cerrada de julio: {planilla}")
    usuarios, _ = mod._cargar_planilla(planilla)
    usuarios = [u for u in usuarios if (u["mz"], u["lt"]) in keys]
    resultado, _ = mod._calcular(
        usuarios, [], [], {}, {}, {}, {}, 1, set(), abonos_rezagados={}
    )
    arrastre = {}
    for row in resultado:
        _, sin_cubrir, total = mod._descomponer_saldo(row)
        arrastre[(row["mz"], row["lt"])] = (sin_cubrir, total)
    pd.DataFrame([
        {"MZ": mz, "LT": lt, "DEUDA_AGUA": values[0][0],
         "CORTE_RECONEXION": values[0][1], "CONVENIO": values[0][2],
         "ACUERDOS_ASAMBLEA": values[0][3], "MULTA": values[0][4],
         "TOTAL_ARRASTRE": values[1]}
        for (mz, lt), values in sorted(arrastre.items())
    ]).to_excel(INPUTS / "mini_arrastre_consolidado_2026-07.xlsx", index=False)
    return arrastre


def _actualizar_documentacion(report, ledger_rows):
    start = "<!-- MINI_PIPELINE_GENERATED_START -->"
    end = "<!-- MINI_PIPELINE_GENERATED_END -->"
    text = DOC.read_text(encoding="utf-8")
    before, marker, rest = text.partition(start)
    if not marker:
        raise RuntimeError(f"Falta marcador de inicio en {DOC}")
    generated, marker_end, after = rest.partition(end)
    if not marker_end:
        raise RuntimeError(f"Falta marcador de fin en {DOC}")

    lines = [
        "## Resultado mini-pipeline y cambios previstos",
        "",
        "Fuente: `mini_resultado_cascada.xlsx`; corrida aislada sobre los 7 lotes.",
        "",
        "| Lote | Abono | Total pagado | Saldo | Actual | Mant. | Anterior | Corte | Convenio | Acuerdos | Multa | Estado |",
        "|---|---:|---:|---:|---:|---:|---:|---:|---:|---:|---:|---|",
    ]
    for row in report.itertuples(index=False):
        lines.append(
            f"| {row.MZ}-{row.LT} | S/{row.ABONO_REZAGADO:.0f} | "
            f"S/{row.TOTAL_PAGADO:.0f} | S/{row.SALDO:.0f} | "
            f"S/{row.SALDO_MES_ACTUAL:.0f} | S/{row.SALDO_MANTENIMIENTO:.0f} | "
            f"S/{row.SALDO_MES_ANTERIOR:.0f} | S/{row.SALDO_CORTE_RECONEXION:.0f} | "
            f"S/{row.SALDO_CONVENIO:.0f} | S/{row.SALDO_ACUERDOS_ASAMBLEA:.0f} | "
            f"S/{row.SALDO_MULTA:.0f} | PENDIENTE_APLICAR |"
        )
    lines.extend([
        "",
        "### Cambios previstos en el ledger real",
        "",
        "`PENDIENTE_APLICAR`: estas filas son la proyeccion del mini-pipeline; el script no escribe el ledger real.",
        "",
        "| Lote | Mes | Concepto | Pago previsto | Source | Estado |",
        "|---|---|---|---:|---|---|",
    ])
    if ledger_rows:
        for row in ledger_rows:
            lines.append(
                f"| {row['MZ']}-{row['LT']} | {row['MES']} | {row['CONCEPTO']} | "
                f"S/{row['PAGO']:.0f} | `{row['SOURCE']}` | PENDIENTE_APLICAR |"
            )
    else:
        lines.append("| — | — | — | S/0 | — | SIN_CAMBIOS |")
    lines.extend([
        "",
        "No ejecutar el ledger real ni `5_cobranza --force` desde este script.",
    ])
    DOC.write_text(before + start + "\n" + "\n".join(lines) + "\n\n" + end + after, encoding="utf-8")


def _generar_pdf(report):
    import fitz

    pdf = fitz.open()
    azul = (26 / 255, 82 / 255, 118 / 255)
    gris = (0.42, 0.45, 0.5)
    negro = (0.12, 0.16, 0.22)
    conceptos = (
        ("Actual", "SALDO_MES_ACTUAL"),
        ("Mant.", "SALDO_MANTENIMIENTO"),
        ("Anterior", "SALDO_MES_ANTERIOR"),
        ("Corte", "SALDO_CORTE_RECONEXION"),
        ("Convenio", "SALDO_CONVENIO"),
        ("Acuerdos", "SALDO_ACUERDOS_ASAMBLEA"),
        ("Multa", "SALDO_MULTA"),
    )

    cover = pdf.new_page()
    cover.insert_text((40, 55), "Mini-pipeline de abonos rezagados", fontsize=16,
                      fontname="hebo", color=azul)
    cover.insert_text((40, 82), "7 lotes · resultado aislado con ledger vivo", fontsize=10,
                      fontname="helv", color=gris)
    y = 125
    for _, row in report.iterrows():
        cover.insert_text((45, y), f"{row['MZ']}-{row['LT']}", fontsize=9,
                          fontname="hebo", color=negro)
        cover.insert_text((115, y), f"Abono S/{row['ABONO_REZAGADO']:.0f}", fontsize=9,
                          fontname="helv", color=negro)
        cover.insert_text((230, y), f"Saldo S/{row['SALDO']:.0f}", fontsize=9,
                          fontname="helv", color=negro)
        y += 16
    cover.insert_text((40, y + 20), "PENDIENTE_APLICAR: no escribe el ledger real.", fontsize=8,
                      fontname="helv", color=gris)

    for _, row in report.iterrows():
        page = pdf.new_page()
        mzlt = f"{row['MZ']}-{row['LT']}"
        page.insert_text((40, 55), f"Predio {mzlt} · mini-pipeline de abono", fontsize=14,
                         fontname="hebo", color=azul)
        page.insert_text((40, 80), "Resultado con abono rezagado incluido", fontsize=9,
                         fontname="helv", color=gris)
        y = 120
        for label, value in (
            ("Deuda antes", row["TOTAL_DEUDA"]),
            ("Pago normal", row["TOTAL_PAGADO"] - row["ABONO_REZAGADO"]),
            ("Abono rezagado", row["ABONO_REZAGADO"]),
            ("Saldo final", row["SALDO"]),
        ):
            page.insert_text((45, y), label, fontsize=10, fontname="hebo", color=negro)
            page.insert_text((220, y), f"S/ {value:,.2f}", fontsize=10,
                             fontname="helv", color=negro)
            y += 18
        y += 15
        page.insert_text((45, y), "Saldo final por concepto", fontsize=10,
                         fontname="hebo", color=azul)
        y += 20
        for label, column in conceptos:
            page.insert_text((55, y), label, fontsize=9, fontname="helv", color=negro)
            page.insert_text((220, y), f"S/ {row[column]:,.2f}", fontsize=9,
                             fontname="helv", color=negro)
            y += 16
        y += 12
        page.insert_text((45, y), "Estado: PENDIENTE_APLICAR", fontsize=9,
                         fontname="hebo", color=gris)

    salida = REFERENCIAS / "mini_reporte_abonos_7_predios.pdf"
    pdf.save(str(salida))
    pdf.close()


def _generar_pdf_formato_historico(keys, report, source):
    sys.path.insert(0, str(ROOT / "4b_reclamos" / "herramienta"))
    import reporte_historico

    reporte_historico.generar_deuda_ledger(
        "2026-07",
        REFERENCIAS / "mini_reporte_abonos_7_predios.pdf",
        predios=set(keys),
        incluir_abonos_rezagados=True,
        mini_report=report,
        mini_abonos=source,
    )


def main():
    if OUT.exists():
        BACKUPS.mkdir(parents=True, exist_ok=True)
        backup = BACKUPS / f"{OUT.name}_{datetime.now():%Y%m%d_%H%M%S}"
        shutil.copytree(OUT, backup)
        print(f"BACKUP_PREVIO={backup}")
        shutil.rmtree(OUT)
    INPUTS.mkdir(parents=True)
    OUTPUTS.mkdir()

    source = _read(mod.ABONOS_REZAGADOS_PATH)
    _corregir_i9_en_mini(source)
    lista_corte = _read(LISTA_CORTE)
    source["_KEY"] = list(zip(source["MZ"].map(mod._norm_mz), source["LT"].map(mod._norm_lt)))
    lista_corte["_KEY"] = list(zip(lista_corte["MZ"].map(mod._norm_mz), lista_corte["LT"].map(mod._norm_lt)))
    target_keys = set(source["_KEY"]) & set(lista_corte["_KEY"])
    source = source[source["_KEY"].isin(target_keys)].drop(columns=["_KEY"])
    source_path = INPUTS / "abonos_rezagados.xlsx"
    source.to_excel(source_path, index=False, startrow=1)
    keys = {
        (mod._norm_mz(row.MZ), mod._norm_lt(row.LT))
        for row in source.itertuples()
    }
    manifest_path = INPUTS / "abonos_rezagados_manifest_2026-08.json"
    manifest_path.write_text(json.dumps(_manifest_mini_i9(keys), indent=2), encoding="utf-8")
    mod.ABONOS_REZAGADOS_PATH = source_path
    mod.ABONOS_MANIFEST_PATH = manifest_path

    ledger = _read(ROOT / "shared" / "seguimiento_pueblo.xlsx")
    ledger_keys = list(zip(ledger["MZ"].map(mod._norm_mz), ledger["LT"].map(mod._norm_lt)))
    ledger["_KEY"] = ledger_keys
    ledger[ledger["_KEY"].isin(keys)].drop(columns=["_KEY"]).to_excel(
        INPUTS / "ledger_subset.xlsx", index=False
    )

    pagos_yape = [p for p in mod._cargar_pagos_yape() if (p["mz"], p["lt"]) in keys]
    pagos_efectivo = [p for p in mod._cargar_pagos_efectivo() if (p["mz"], p["lt"]) in keys]
    pagos_yape = mod._aplicar_correcciones_lote(pagos_yape, mod._leer_correcciones())
    pagos_efectivo = mod._aplicar_correcciones_lote(pagos_efectivo, mod._leer_correcciones())
    pd.DataFrame(pagos_yape).to_excel(
        INPUTS / "pagos_yape_filtrados.xlsx", index=False
    )
    pd.DataFrame(pagos_efectivo).to_excel(
        INPUTS / "pagos_efectivo_filtrados.xlsx", index=False
    )

    planilla_path = mod._validar_inputs()
    usuarios, mes_ano = mod._cargar_planilla(planilla_path)
    shutil.copy2(planilla_path, INPUTS / planilla_path.name)
    saldos_ledger = {
        concepto: mod.repo.get_saldos_bulk(concepto, mes_ano)
        for concepto in ("CONVENIO", "MULTA", "ACUERDOS")
    }
    for usuario in usuarios:
        key = (usuario["mz"], usuario["lt"])
        usuario["convenio"] = max(0.0, float(saldos_ledger["CONVENIO"].get(key, 0.0)))
        usuario["multa"] = max(0.0, float(saldos_ledger["MULTA"].get(key, 0.0)))
        usuario["acuerdos_asamblea"] = max(
            0.0, float(saldos_ledger["ACUERDOS"].get(key, 0.0))
        )
    usuarios = [u for u in usuarios if (u["mz"], u["lt"]) in keys]
    arrastre_cerrado = _arrastre_cerrado_mini(keys)
    for usuario in usuarios:
        agua = arrastre_cerrado[(usuario["mz"], usuario["lt"])][0]
        usuario["mes_anterior"] = agua[0]
        usuario["corte_reconexion"] = agua[1]
    py_keys = {(p["mz"], p["lt"]) for p in pagos_yape}
    ef_keys = {(p["mz"], p["lt"]) for p in pagos_efectivo}
    blancos = {k: v for k, v in mod._cargar_blancos(mes_ano).items() if k in keys}
    aportes = {k: v for k, v in mod._cargar_aportes_tanque_manuales(mes_ano).items() if k in {f"{m}-{l}" for m, l in keys}}
    dev_yape = {k: v for k, v in mod._cargar_retornos_yape().items() if k in keys}
    dev_efec = {k: v for k, v in mod._cargar_retornos_efectivo().items() if k in keys}
    dev_devuelto = {k: v for k, v in mod._cargar_devueltos_yape().items() if k in keys}
    abonos = mod._cargar_abonos_rezagados(mes_ano)
    resultado, _ = mod._calcular(
        usuarios, pagos_yape, pagos_efectivo, blancos,
        dev_yape, dev_efec, dev_devuelto, 1, set(),
        aportes_tanque=aportes, abonos_rezagados=abonos,
    )
    reports = []
    ledger_rows = []
    saldo_campos = (
        "mes_actual", "mantenimiento", "mes_anterior", "corte_reconexion",
        "convenio", "acuerdos_asamblea", "multa",
    )
    for row in resultado:
        saldo_final = {c: round(row[c], 2) for c in saldo_campos}
        mod._aplicar_waterfall(
            saldo_final, row["total_pagado_normal"], _CAMPOS_WATERFALL_CICLO_MINI
        )
        abono_cerrado, abono_vigente = abonos.get((row["mz"], row["lt"]), (0.0, 0.0))
        antes_abono = saldo_final.copy()
        mod._aplicar_waterfall(
            saldo_final, abono_cerrado, mod._CAMPOS_WATERFALL_REIDENTIFICACION
        )
        mod._aplicar_waterfall(
            saldo_final, abono_vigente, _CAMPOS_WATERFALL_CICLO_MINI
        )
        abono = {
            "CONVENIO": round(antes_abono["convenio"] - saldo_final["convenio"], 2),
            "ACUERDOS": round(
                antes_abono["acuerdos_asamblea"] - saldo_final["acuerdos_asamblea"], 2
            ),
            "MULTA": round(antes_abono["multa"] - saldo_final["multa"], 2),
        }
        reports.append({
            "MZ": row["mz"], "LT": row["lt"], "MES_ANO": mes_ano,
            "PAGO_YAPE": row["monto_yape"], "PAGO_EFECTIVO": row["monto_efectivo"],
            "ABONO_REZAGADO": row["abono_rezagado"],
            "TOTAL_DEUDA": row["total_a_pagar"], "TOTAL_PAGADO": row["total_pagado"],
            "SALDO": row["saldo"], "CONVENIO": abono["CONVENIO"],
            "ACUERDOS": abono["ACUERDOS"], "MULTA": abono["MULTA"],
            **{f"DEUDA_{c.upper()}": row[c] for c in saldo_campos},
            **{f"SALDO_{c.upper()}": saldo_final[c] for c in saldo_campos},
        })
        for concepto, monto in abono.items():
            if monto > mod.TOL:
                ledger_rows.append({"MZ": row["mz"], "LT": row["lt"], "MES": mes_ano, "CONCEPTO": concepto, "PAGO": monto, "SOURCE": "abonos_rezagados"})

    report = pd.DataFrame(reports).sort_values(["MES_ANO", "MZ", "LT"])
    report.to_excel(OUTPUTS / "mini_resultado_cascada.xlsx", index=False)
    report[
        [
            "MZ", "LT", "MES_ANO", "ABONO_REZAGADO", "CONVENIO", "ACUERDOS", "MULTA",
            "SALDO", "SALDO_MES_ACTUAL", "SALDO_MANTENIMIENTO", "SALDO_MES_ANTERIOR",
            "SALDO_CORTE_RECONEXION", "SALDO_CONVENIO", "SALDO_ACUERDOS_ASAMBLEA",
            "SALDO_MULTA",
        ]
    ].rename(columns={
        "CONVENIO": "ABONO_CONVENIO",
        "ACUERDOS": "ABONO_ACUERDOS",
        "MULTA": "ABONO_MULTA",
        "SALDO": "SALDO_FINAL",
    }).to_excel(OUTPUTS / "mini_reporte_abonos.xlsx", index=False)
    report[
        [
            "MZ", "LT", "MES_ANO", "ABONO_REZAGADO", "CONVENIO", "ACUERDOS", "MULTA",
            "SALDO", "SALDO_MES_ACTUAL", "SALDO_MANTENIMIENTO", "SALDO_MES_ANTERIOR",
            "SALDO_CORTE_RECONEXION", "SALDO_CONVENIO", "SALDO_ACUERDOS_ASAMBLEA",
            "SALDO_MULTA",
        ]
    ].rename(columns={
        "CONVENIO": "ABONO_CONVENIO",
        "ACUERDOS": "ABONO_ACUERDOS",
        "MULTA": "ABONO_MULTA",
        "SALDO": "SALDO_FINAL",
    }).to_excel(REFERENCIAS / "mini_reporte_abonos.xlsx", index=False)
    _generar_pdf_formato_historico(keys, report, source)
    if ACTUALIZAR_DOC:
        _actualizar_documentacion(report, ledger_rows)
    (OUT / "README.txt").write_text(
        "Mini-corrida aislada: interseccion lista_corte ∩ abonos_rezagados.\n"
        f"Filas fuente: {len(source)}\n"
        f"Predios calculados: {len(report)}\n"
        f"Lotes objetivo: {', '.join(f'{m}-{l}' for m, l in sorted(keys))}\n"
        f"Pagos Yape filtrados: {len(py_keys)} claves\n"
        f"Pagos efectivo filtrados: {len(ef_keys)} claves\n"
        "La planilla es la misma fuente viva de 5_cobranza; no se mezcla con archivos historicos.\n"
        "No escribe el ledger real ni ejecuta 5_cobranza/main.py.\n",
        encoding="utf-8",
    )
    print(f"OK mini-corrida lista_corte&abonos: source={len(source)} predio_mes={len(report)}")
    print(f"TARGET_KEYS={','.join(f'{m}-{l}' for m, l in sorted(keys))}")
    pd.DataFrame(ledger_rows).to_excel(OUTPUTS / "mini_ledger_predicho.xlsx", index=False)
    print(f"OUTPUT={OUT}")


if __name__ == "__main__":
    main()
