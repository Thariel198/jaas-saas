# =========================IMPORTS===========================
from rapidfuzz import fuzz

# ========================CONFIGURACION======================
UMBRAL_NOMBRE = 50   # fuzzy origen vs nombre registrado
MIN_TOKEN     = 3    # longitud mínima de token significativo

# ========================UTILIDADES=========================
def similitud_nombre(origen: str, nombre: str) -> float:
    """
    Calcula similitud fuzzy entre un origen y el nombre registrado.
    Ignora orden de palabras y asteriscos.
    """
    return fuzz.token_sort_ratio(
        origen.upper().replace("*", "").strip(),
        nombre.upper().replace("*", "").strip()
    )

def tokens(s: str) -> set:
    """
    Extrae tokens significativos de un string.
    Ignora iniciales cortas y prefijos de banco.
    """
    ignorar = {"BCP", "YAPE", "PLIN", "DEL", "DE", "LA", "LAS", "LOS", "EL"}
    return set(
        t for t in s.upper()
                    .replace("*", "")
                    .replace("-", " ")
                    .replace(".", " ")
                    .split()
        if len(t) >= MIN_TOKEN and t not in ignorar
    )

def comparte_token(a: str, b: str) -> bool:
    """
    Retorna True si dos strings comparten al menos un token significativo.
    Detecta familia por apellido compartido.
    """
    return bool(tokens(a) & tokens(b))

# ========================FILTRADO===========================
def filtrar_origenes_validos(origenes: list, nombre: str) -> list:
    """
    Conserva solo los orígenes válidos descartando intrusos.

    Un origen es válido si cumple al menos una condición:
        1. Comparte token exacto con otro origen     → familia · apellido compartido
        2. Comparte token exacto con el nombre       → nombre registrado del usuario
        3. Similitud fuzzy >= UMBRAL_NOMBRE vs nombre → nombres truncados con *

    Intrusos — sin ninguna condición — se descartan siempre.
    Motor_matching usará solo los válidos para buscar.

    Args:
        origenes : lista de orígenes brutos
        nombre   : nombre registrado del usuario

    Returns:
        lista de orígenes válidos · vacía si ninguno pasa
    """
    if not origenes:
        return []

    if len(origenes) == 1:
        origen = origenes[0]
        if comparte_token(origen, nombre) or similitud_nombre(origen, nombre) >= UMBRAL_NOMBRE:
            return origenes
        return []

    validos = []
    for i, origen in enumerate(origenes):
        tiene_par    = any(i != j and comparte_token(origen, origenes[j])
                           for j in range(len(origenes)))
        token_nombre = comparte_token(origen, nombre)
        fuzzy_nombre = similitud_nombre(origen, nombre) >= UMBRAL_NOMBRE

        if tiene_par or token_nombre or fuzzy_nombre:
            validos.append(origen)

    return validos
