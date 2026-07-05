# Prueba de formato: genera solo las primeras 3 boletas usando el MISMO código
# de producción (main.py) — sin merge, sin borrar Outputs.
# Uso: py Prueba.py [N]   (N = cuántas boletas, default 3)
import sys

import main


def prueba(limit: int = 3):
    main.OUTPUT_DIR.mkdir(exist_ok=True)
    main.IMAGES_DIR.mkdir(exist_ok=True)

    df = main.load_data()
    main.validate_data(df)
    df = main.process_data(df)
    grouped = main.group_data(df)
    main.generate_boletas(grouped, limit=limit)

    print(f"\n[PRUEBA] Revisar formato en {main.OUTPUT_DIR.resolve()}")
    print("[PRUEBA] Si el formato está bien: py main.py (genera todas y consolida)")


if __name__ == "__main__":
    n = int(sys.argv[1]) if len(sys.argv) > 1 and sys.argv[1].isdigit() else 3
    prueba(n)
