from typing import List, Dict, Set

# Este modulo se encarga de normalizar datos puros, pero no tiene la capacidad de trabajar con archivos.
# Los datos deberan ser entregados en formato de lista, y se entregaran resultados de normalizacion.

CAPITALIZATION = {
    "upper": str.upper,
    "lower": str.lower,
    "capitalize": str.capitalize,
    "titlecase": str.title,
    "none": lambda x: x,
}

def simple_norm(text: str, strip: bool = True, capitalization: str = "CC") -> str:
    # Aplicar capitalizacion de caracteres (Mayusculas y minusculas)
    result = CAPITALIZATION[capitalization](text)
    # Aplicar stripping
    if strip:
        result = result.strip()

    return result

def normalize_simple(data: List[str]) -> List[str]:
    result = list(map(normalize_simple, data))
    return result
