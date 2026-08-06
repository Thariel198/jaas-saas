# 4_pagos/efectivo/registrar_entrega.py — CLI de la tesorera para appendear a entregas.xlsx
#
# La tesorera dicta el monto; este CLI lo appendea vía entregas_repo (append-only).
# En SaaS, un endpoint POST llama a la misma función registrar_entrega().
#
# Uso interactivo (la tesorera dicta):
#   python registrar_entrega.py
#
# Uso por flags (scripts / tests / agente):
#   python registrar_entrega.py --fecha 03/07/2026 --cobrador "Pedro Quispe" --efectivo 450 --yape 120
#
# Corrección de una entrega ya registrada (se appendea el DELTA, nunca se edita):
#   python registrar_entrega.py --fecha 03/07/2026 --cobrador "Maria Garcia" \
#       --correccion --efectivo 540 --yape 60 --motivo "tipeo: efectivo era 540"
#   (--efectivo/--yape son el TOTAL CORRECTO; el CLI calcula el delta contra lo ya registrado)

import argparse
import sys
from datetime import datetime

import entregas_repo as repo


def _slug(cobrador: str) -> str:
    s = "".join(c for c in cobrador.lower() if c.isalnum())
    return s or "cobrador"


def main() -> None:
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass

    ap = argparse.ArgumentParser(description="Registrar una entrega de la tesorera (append-only)")
    ap.add_argument("--fecha", help="dd/mm/aaaa — día que pagó el usuario")
    ap.add_argument("--cobrador")
    ap.add_argument("--efectivo", type=float)
    ap.add_argument("--yape", type=float)
    ap.add_argument("--correccion", action="store_true",
                    help="los montos son el TOTAL correcto; se appendea el delta")
    ap.add_argument("--motivo", default="", help="obligatorio con --correccion")
    ap.add_argument("--source", default="cli_registrar_entrega")
    ap.add_argument("--audit-ref", default=None)
    a = ap.parse_args()

    fecha    = a.fecha    or input("FECHA (dd/mm/aaaa): ").strip()
    cobrador = a.cobrador or input("COBRADOR: ").strip()
    efectivo = a.efectivo if a.efectivo is not None else float(input("EFECTIVO recibido: ") or 0)
    yape     = a.yape     if a.yape     is not None else float(input("YAPE recibido: ") or 0)

    fecha_dt = repo._parse_fecha(fecha)     # valida el formato antes de nada
    cobrador = repo._norm_cobrador(cobrador)
    motivo   = a.motivo

    if a.correccion:
        motivo = motivo or input("MOTIVO de la corrección: ").strip()
        if not motivo:
            sys.exit("ERROR: una corrección requiere --motivo")
        actual = repo.suma_entregas().get(
            (repo._fecha_key(fecha_dt), cobrador.upper()),
            {"efectivo": 0.0, "yape": 0.0})
        d_efec = round(efectivo - actual["efectivo"], 2)
        d_yape = round(yape - actual["yape"], 2)
        print(f"  registrado hoy: efec={actual['efectivo']:.2f} yape={actual['yape']:.2f}")
        print(f"  correcto:       efec={efectivo:.2f} yape={yape:.2f}")
        print(f"  delta a appendear: efec={d_efec:+.2f} yape={d_yape:+.2f}")
        efectivo, yape = d_efec, d_yape
        if efectivo == 0 and yape == 0:
            print("  nada que corregir (delta 0) — no se appendea fila.")
            return

    audit_ref = a.audit_ref or (
        f"ENT-{fecha_dt:%Y%m%d}-{_slug(cobrador)}-{datetime.now():%H%M%S}")

    res = repo.registrar_entrega(fecha, cobrador, efectivo, yape,
                                 source=a.source, audit_ref=audit_ref, motivo=motivo)
    if res["skipped"]:
        print(f"skip — ya existe (source={a.source}, audit_ref={audit_ref})")
    else:
        print(f"OK — entrega {fecha_dt:%d/%m/%Y} {cobrador} "
              f"efec={efectivo:+.2f} yape={yape:+.2f} [{audit_ref}]")


if __name__ == "__main__":
    main()
