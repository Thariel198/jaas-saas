"""Reporte de colisiones del universo de abonos normales de julio."""
from pathlib import Path

import pandas as pd

import main


ROOT = Path(__file__).resolve().parents[1]
MES_APLICA = "2026-07"
SOURCE = "rebuild_abonos_normales"


def _planilla(path: Path) -> dict[tuple[str, str], dict]:
    df = pd.read_excel(path, header=1)
    df.columns = main._norm_cols(df)
    out = {}
    for _, f in df.iterrows():
        mz = main._norm_mz(f.get("MZ"))
        lt = main._norm_lt(f.get("LT"))
        if mz and lt:
            out[(mz, lt)] = {
                "mes_anterior": main._float(f.get("MES_ANTERIOR")),
                "corte_reconexion": main._float(f.get("CORTE_RECONEXION")),
                "convenio": main._float(f.get("CONVENIO")),
                "multa": main._float(f.get("MULTA")),
                "acuerdos_asamblea": main._float(f.get("ACUERDOS_ASAMBLEA")),
                "pagado": main._float(f.get("MONTO_YAPE")) + main._float(f.get("MONTO_EFECTIVO")),
            }
    return out


def main_backfill() -> None:
    julio = _planilla(ROOT / "7_cierre" / "archivo" / "2026-07" / "planilla_cobrado.xlsx")
    junio = _planilla(ROOT / "7_cierre" / "archivo" / "2026-06" / "planilla_cobrado.xlsx")
    abonos = main._cargar_abonos_rezagados(MES_APLICA)
    campos = main._CAMPOS_WATERFALL_REIDENTIFICACION
    conceptos = ("CONVENIO", "ACUERDOS", "MULTA")
    total = 0.0
    colisiones = []
    aplicados = []

    for key, (cerrado, vigente) in abonos.items():
        monto = round(cerrado + vigente, 2)
        deuda = julio.get(key, {})
        if not deuda or sum(deuda.get(c, 0.0) for c in campos) <= main.TOL:
            deuda = junio.get(key, deuda)
        if not deuda:
            raise ValueError(f"Abono sin planilla: {key}")

        comps = [deuda[c] for c in campos]
        restante = deuda["pagado"]
        residual = []
        for comp in comps:
            cubierto = min(max(comp, 0.0), restante)
            residual.append(round(max(comp, 0.0) - cubierto, 2))
            restante = round(max(restante - cubierto, 0.0), 2)

        asignaciones = []
        tiene_colision = False
        for idx, campo in enumerate(campos):
            aplicar = min(residual[idx], monto)
            if idx >= 2 and aplicar > main.TOL:
                concepto = conceptos[idx - 2]
                saldo = main.repo.get_saldo(key[0], key[1], concepto, MES_APLICA)
                asignaciones.append((concepto, round(aplicar, 2)))
                if saldo < aplicar - main.TOL:
                    tiene_colision = True
                    colisiones.append({
                        "MZ": key[0], "LT": key[1], "CONCEPTO": concepto,
                        "MONTO_PREVISTO": round(aplicar, 2),
                        "SALDO_LEDGER": round(saldo, 2),
                        "PENDIENTE_CALCULAR_AL_CIERRE": True,
                        "MES_CICLO": "2026-06",
                        "MES_ANO_APLICA": MES_APLICA,
                    })
            monto = round(monto - aplicar, 2)
            if monto <= main.TOL:
                break
        if not tiene_colision:
            for concepto, aplicar in asignaciones:
                ref = f"abono_julio_{key[0]}_{key[1]}_{concepto.lower()}"
                result = main.repo.registrar_pago(
                    key[0], key[1], concepto, MES_APLICA, aplicar,
                    source=SOURCE, audit_ref=ref,
                    clase="ABONO_REZAGADO",
                    motivo="Abono normal fuera de ventana; BALDE=agua",
                )
                if not result["skipped"]:
                    aplicados.append((key, concepto, aplicar))
        total += cerrado + vigente

    print(f"abonos julio: {len(abonos)} predios · S/{total:.2f}")
    print(f"colisiones: {len(colisiones)}")
    print(f"pagos escritos: {len(aplicados)}")
    for r in colisiones:
        print(f"{r['MZ']}-{r['LT']} | {r['CONCEPTO']} | previsto S/{r['MONTO_PREVISTO']:.2f} | saldo ledger S/{r['SALDO_LEDGER']:.2f}")
    output = ROOT / "Pendiente" / "abonos_rezagados_pendientes_2026-07.md"
    lines = [
        "# Abonos rezagados pendientes — 2026-07",
        "",
        "Estos casos pertenecen exclusivamente a `abonos_rezagados.xlsx`. No se registran todavía en el ledger; deben recalcularse contra el saldo final del ciclo.",
        "",
        "| MZ | LT | CONCEPTO | MONTO_PREVISTO | SALDO_LEDGER | MES_CICLO | MES_ANO_APLICA |",
        "|---|---:|---|---:|---:|---|---|",
    ]
    lines.extend(
        f"| {r['MZ']} | {r['LT']} | {r['CONCEPTO']} | S/{r['MONTO_PREVISTO']:.2f} | S/{r['SALDO_LEDGER']:.2f} | {r['MES_CICLO']} | {r['MES_ANO_APLICA']} |"
        for r in colisiones
    )
    output.write_text("\n".join(lines) + "\n", encoding="utf-8")
    print(f"reporte pendiente: {output}")


if __name__ == "__main__":
    main_backfill()
