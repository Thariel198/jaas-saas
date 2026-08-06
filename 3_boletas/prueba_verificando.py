# Prueba puntual: genera SOLO las boletas de los predios con concepto en
# VERIFICANDO (pendientes_convenio_multas.xlsx), usando el mismo código de
# produccion (main.py). No borra Outputs, no consolida.
# Uso: py prueba_verificando.py [N]   (N = cuantos casos, default 10)
import sys

import main


def prueba_verificando(limit: int = 10):
    main.OUTPUT_DIR.mkdir(exist_ok=True)
    main.IMAGES_DIR.mkdir(exist_ok=True)

    df = main.load_data()
    main.validate_data(df)
    df = main.process_data(df)

    pendientes = main.load_pendientes_convenio_multas()
    claves = {(mz, lt) for (mz, lt, _concepto) in pendientes.keys()}
    df["_mz_n"] = df["MZ"].apply(main._norm_mz)
    df["_lt_n"] = df["LT"].apply(main._norm_lt)
    df_filtrado = df[df.apply(lambda r: (r["_mz_n"], r["_lt_n"]) in claves, axis=1)].drop(columns=["_mz_n", "_lt_n"])

    print(f"[PRUEBA VERIFICANDO] {len(claves)} predios en pendientes · {df_filtrado.shape[0]} filas encontradas en DATA_boletas · probando {limit}")

    grouped = main.group_data(df_filtrado)
    main.generate_boletas(grouped, limit=limit)

    print(f"\n[PRUEBA VERIFICANDO] Revisar formato en {main.OUTPUT_DIR.resolve()}")


if __name__ == "__main__":
    n = int(sys.argv[1]) if len(sys.argv) > 1 and sys.argv[1].isdigit() else 10
    prueba_verificando(n)
