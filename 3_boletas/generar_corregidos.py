import shutil
import sys
import main as boletas_main

# Recibos corregidos a mano por reclamo (ver inputs/reclamos_2026-08-01.xlsx +
# inputs/reclamos_2026-08-01/README.md). Se van agregando conforme se confirma
# cada caso con la directiva.
RECIBOS_OBJETIVO = [
    # Bloque A -- 13 reclamos dictados de las fotos
    18114,  # G-14  Margarita Gomez Bonifacio
    18084,  # E-14B Juan Saavedra Saavedra
    18170,  # J-1   Comedor Popular Club de Madres
    18192,  # K-9   Fortunato Vargas Cabello
    18191,  # K-8   Victor Teodoro Flores Durand
    18342,  # T-14  Pedro Candacho Huarac
    18368,  # V-14  Leonardo Huamani Sotelo
    17992,  # B-8   Rosalina Olimpia Ciriaco Sotelo
    18096,  # F-10  Herminio Lucero Trujillo
    18086,  # F-1   Maria Godo Sifuentes
    18093,  # F-7   Victor Laurencio Valladares
    18056,  # D-6   Hermelinda Jara Trujillo
    18245,  # O-2   Carmen Ingaruca Julca
    # Bloque B -- 11 predios del lote de SALDO negativo
    17979,  # A-8   Victor Melgarejo Corcino
    17989,  # B-5   Pompeyo Celestino Lliuya
    18007,  # C-1   Odilon Cerna Romero
    18014,  # C-7   Victor Lopez Trujillo
    18080,  # E-12  Teofila Fernandez Reyes
    18162,  # I-11  Dominga Chacara Lopez
    18167,  # I-16  Adolfo Rosario Rojas
    18172,  # J-3   Vilma Celestino Villafana
    18200,  # K-17  Marcial Sanchez Araoz
    18185,  # K-2   Antonio Espinoza Sifuentes
    18285,  # P-12  Judith Venturo Rosales
    # Bloque C -- mismo patron, encontrado despues
    18144,  # H-16  Gregorio Tolentino Sanchez
    18104,  # G-4   Natalia Chinchay Collas
]

CORREGIDOS_DIR = boletas_main.OUTPUT_DIR.parent / "corregidos"

if CORREGIDOS_DIR.exists():
    shutil.rmtree(CORREGIDOS_DIR)
CORREGIDOS_DIR.mkdir(parents=True)

boletas_main.OUTPUT_DIR = CORREGIDOS_DIR
boletas_main.IMAGES_DIR = CORREGIDOS_DIR / "Imagenes"
boletas_main.IMAGES_DIR.mkdir(exist_ok=True)

df = boletas_main.load_data()
boletas_main.validate_data(df)
df = boletas_main.process_data(df)
df = df[df["NUMERO DE RECIBO"].isin(RECIBOS_OBJETIVO)]

if df.empty:
    print("[ERROR] Ningun recibo de RECIBOS_OBJETIVO encontrado en DATA_boletas.xlsx")
    sys.exit(1)

grouped = boletas_main.group_data(df)
boletas_main.generate_boletas(grouped, limit=None)
print(f"[OK] {len(RECIBOS_OBJETIVO)} recibo(s) corregido(s) en {CORREGIDOS_DIR}")
