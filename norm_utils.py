from typing import Dict, List, Set, Tuple
from difflib import SequenceMatcher
import re

STRICT_RUT_PATTERN = re.compile(r'^((\d{1,3}(?:\.\d{3}){2})|(\d{7,9}))-[\dkK]$')
LAX_RUT_PATTERN = re.compile(r'^[\d.]{7,9}-?[\dkK]$')

def similarity(a: str, b: str) -> float:
    """
    Retorna un valor de similitud entre dos textos, de 0 a 1.
    """
    return SequenceMatcher(None, a.lower(), b.lower()).ratio()

def find_potential_matches(data : List[str], threshold : float = 0.8) -> Set[str]:
    """
    Retorna un set de tuplas que contienen valores que son parecidos (posiblemente el mismo)
    """
    matches = set()
    for i, w1 in enumerate(data):
        for j, w2 in enumerate(data):
            if i >= j:
                continue
            if similarity(w1, w2) >= threshold:
                pair = tuple(sorted([w1, w2]))
                matches.add(pair)
    return matches

def unify_by_user(data : List[str], threshold: float = 0.8) -> List[str]:
    matches = find_potential_matches(data, threshold = threshold)
    unified_data = data.copy()
    for w1, w2 in matches:
        resp_yn = ""
        while resp_yn != "y" and resp_yn != "n":
            resp_yn = input(f"Iguales? \"{w1}\" - \"{w2}\" (y/n)")
        if resp_yn == "y":
            resp_yn = "n"
            unified = ""
            while resp_yn == "n":
                unified = input("Ingrese unificación: ")
                resp_yn = ""
                while resp_yn != "y" and resp_yn != "n":
                    resp_yn = input(f"Seguro? \"{w1}\" - \"{w2}\" -> \"{unified}\" (y/n)")
            if w1 in unified_data:
                unified_data.remove(w1)
            if w2 in unified_data:
                unified_data.remove(w2)
            unified_data.append(unified)
    return unified_data

def calculate_dv(rut_num: str) -> str:
    """
    Calcula el dígito verificador (DV) de un RUT chileno.
    :param rut_num: string con solo dígitos (sin puntos ni guion)
    """
    reversed_digits = map(int, reversed(rut_num))
    factors = [2, 3, 4, 5, 6, 7]

    total = 0
    for i, d in enumerate(reversed_digits):
        total += d * factors[i % len(factors)]

    remainder = 11 - (total % 11)
    if remainder == 11:
        return "0"
    elif remainder == 10:
        return "k"
    else:
        return str(remainder)

def check_rut_normalize(rut: str, validation_mode: str = "lax", norm_mode: str = "standard") -> Tuple[bool, str]:
    """
    Valida un RUT y lo normaliza.
    Modo de uso:
    valido, normalizado = check_rut_normalize("19837745-7", validation_mode="strict", norm_mode="dotted")
    :param rut:
    Rut con dígito verificador
    :param validation_mode:
    Rigidez de la validación
    "lax" -> Permite cualquier estilo de rut, ignorando todos los puntos y guiones sin importar su ubicacion. NO permite ruts con el numero incorrecto de digitos.
    "strict" -> Solo permite ruts escritos en los siguientes formatos... XXXXXXX-X y X.XXX.XXX-X con entre 7 y 9 digitos + verificador
    :param norm_mode:
    "standard" -> Rut sin puntos y con guion
    "dotted" -> Rut con puntos y con guion
    "none" -> No normaliza el rut
    :return:
    Tuple[bool, str] -> Una tupla donde el primer elemento indica si el rut es valido, y el segundo elemento
    es el rut normalizado.
    """
    # Validacion de parametros de comportamiento
    if validation_mode not in ("strict", "lax"):
        raise ValueError(f"Invalid validation mode: \"{validation_mode}\"")
    if norm_mode not in ("standard", "dotted", "none"):
        raise ValueError(f"Invalid normalization mode \"{norm_mode}\".")

    # Fase de validacion de formato
    valid_format = True
    norm = rut.lower()
    if validation_mode == "strict":
        if not STRICT_RUT_PATTERN.match(rut):
            valid_format = False
    elif validation_mode == "lax":
        if not LAX_RUT_PATTERN.match(rut):
            valid_format = False

    # Fase de validacion de digito
    final_rut = rut
    valid = False
    if valid_format:
        norm = "".join(norm.split("."))
        dv = ""
        # Aquí tenemos garantía de que el rut es sin puntos y con o sin guion, y que de haber guion, solo
        # tiene uno y en el lugar correcto. (basado en regex pattern matching)
        if norm[-2] == "-":
            norm, dv = norm.split("-")
        else:
            norm, dv = norm[:-1], norm[-1]
        valid = dv == calculate_dv(norm)

        # También tenemos garantía de que el rut tiene entre 7 y 9 dígitos
        if norm_mode == "standard":
            final_rut = norm + "-" + dv
        elif norm_mode == "dotted":
            final_rut = f"{norm[:-6]}.{norm[-6:-3]}.{norm[-3:]}-{dv}"
    # Retornar resultados
    # NOTA: No podemos normalizar un rut que no tiene formato válido. (porque podría ser cualquier cosa)
    return valid, final_rut
