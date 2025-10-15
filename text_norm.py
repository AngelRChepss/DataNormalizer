from dataclasses import dataclass
from typing import List, Dict

# Este modulo se encarga de normalizar datos puros, pero no tiene la capacidad de trabajar con archivos.
# Los datos deberan ser entregados en formato de lista, y se entregaran resultados de normalizacion.


# Lista de todos los tipos de espacios raros que hay que eliminar
WEIRD_SPACES = [
    "\u00A0",  # NO-BREAK SPACE
    "\u2000",  # EN QUAD
    "\u2001",  # EM QUAD
    "\u2002",  # EN SPACE
    "\u2003",  # EM SPACE
    "\u2004",  # THREE-PER-EM SPACE
    "\u2005",  # FOUR-PER-EM SPACE
    "\u2006",  # SIX-PER-EM SPACE
    "\u2007",  # FIGURE SPACE
    "\u2008",  # PUNCTUATION SPACE
    "\u2009",  # THIN SPACE
    "\u200A",  # HAIR SPACE
    "\u200B",  # ZERO WIDTH SPACE
    "\u202F",  # NARROW NO-BREAK SPACE
    "\u205F",  # MEDIUM MATHEMATICAL SPACE
    "\u3000",  # IDEOGRAPHIC SPACE
    "\uFEFF",  # ZERO WIDTH NO-BREAK SPACE / BOM
]

# Diccionario de todos los tildes is sus respectivas letras sin tilde
TILDES = {
    "á": "a", "é": "e", "í": "i", "ó": "o", "ú": "u",
    "Á": "A", "É": "E", "Í": "I", "Ó": "O", "Ú": "U",
}

# Lista de caracteres invisibles que deben ser eliminados
INVISIBLES = [
    "\u200B",  # ZERO WIDTH SPACE
    "\u200C",  # ZERO WIDTH NON-JOINER
    "\u200D",  # ZERO WIDTH JOINER
    "\u200E",  # LEFT-TO-RIGHT MARK
    "\u200F",  # RIGHT-TO-LEFT MARK
    "\u202A",  # LEFT-TO-RIGHT EMBEDDING
    "\u202B",  # RIGHT-TO-LEFT EMBEDDING
    "\u202C",  # POP DIRECTIONAL FORMATTING
    "\u202D",  # LEFT-TO-RIGHT OVERRIDE
    "\u202E",  # RIGHT-TO-LEFT OVERRIDE
    "\u2060",  # WORD JOINER
    "\u2061",  # FUNCTION APPLICATION
    "\u2062",  # INVISIBLE TIMES
    "\u2063",  # INVISIBLE SEPARATOR
    "\u2064",  # INVISIBLE PLUS
    "\u2066",  # LEFT-TO-RIGHT ISOLATE
    "\u2067",  # RIGHT-TO-LEFT ISOLATE
    "\u2068",  # FIRST STRONG ISOLATE
    "\u2069",  # POP DIRECTIONAL ISOLATE
    "\uFEFF",  # ZERO WIDTH NO-BREAK SPACE (BOM)
]

def rmv_tildes(x : str) -> str:
    """
    Removes tildes from a string
    """
    for t in TILDES.keys():
        x = x.replace(t, TILDES[t])
    return x

def collapse(string: str, character: str) -> str:
    if len(character) != 1:
        raise ValueError("character must be a single character")
    print("collapsing")
    return character.join(string.split(character))

def repl_fixed(string: str, rep_list: List[str], replacement: str) -> str:
    for rep in rep_list:
        string = string.replace(rep, replacement)
    return string

def repl_list(string: str, rep_dict: Dict[str, str]) -> str:
    for k, v in rep_dict.items():
        string = string.replace(k, v)
    return string

def repl_words(string: str, rep_dict: Dict[str, str], word_separator = " ") -> str:
    if word_separator == "":
        return repl_list(string, rep_dict)
    string = f"{word_separator}{string}{word_separator}"
    for k, v in rep_dict.items():
        word = f"{word_separator}{k}{word_separator}"
        replacer = f"{word_separator}{v}{word_separator}"
        string = string.replace(word, replacer)
    return string[1:-1]

def rmv_simple(string: str, rem_list: List[str]) -> str:
    for rem in rem_list:
        string = string.replace(rem, "")
    return string

def rmv_list(string: str, rem_list: List[str], exhaust = True, surround : str = "") -> str:
    """
    Remueve todos los caracteres o palabras de una lista en x.
    :param string: Cadena de texto por normalizar
    :param rem_list: Lista de caracteres o palabras por remover de string
    :param exhaust: Si al hacer las remociones aparecen nuevas palabras por remover,
    controla si se deben también remover estas.
    :param surround: Caracter que se agrega a los lados de cada palabra que se debe remover.
    """
    removes = rem_list if surround == "" else [f"{surround}{word}{surround}" for word in rem_list]
    result = string
    removed = True
    while removed:
        removed = False
        for rem in removes:
            result = result.replace(rem, "")
        if exhaust:
            removed = len(result) != len(string)
        string = result
    return result

def patch_cap(text: str, patches: List[str]) -> str:
    """
    Recapitaliza todas las excepciones (palabras en la lista) dentro del texto, usando
    la capitalizacion entregada en la lista.
    """
    if not patches:
        return text

    result = text
    lower_text = result.lower()
    for rule in patches:
        # Comparamos la regla en minusculas con el texto en minusculas para encontrar posibles
        # igualdades
        lower_rule = rule.lower()
        start = 0
        while True:
            # Buscamos si existe la regla dentro del texto
            idx = lower_text.find(lower_rule, start)
            if idx == -1:
                break
            # si existe, entonces reemplazamos la regla capitalizada sobre el texto capitalizado
            end = idx + len(lower_rule)
            result = result[:idx] + rule + result[end:]
            start = idx + len(rule)
    return result

def naming_case(s : str):
    """
    Como Title Case, pero las palabras y, de, la, las, el, los, las, a, e, o se mantienen en minusculas.
    Si una es la primera palabra, sí se capitaliza.
    """
    connectors = ["y", "de", "la", "las", "el", "los", "a", "e", "o"]
    # Capitalizar primera letra de todas las palabras
    result = s.title()
    # Decapitalizar primera letra de todos los conectores
    for c in connectors:
        # Se agregan espacios en los costados para garantizar que es una palabra completa
        spaced = f" {c.title()} "
        replacer = f" {c} "
        result = result.replace(spaced, replacer)
    # Capitalizar primera letra aunque sea conector
    if len(result) > 0:
        result = result[0].upper() + result[1:]
    return result

CAPITALIZATION = {
    "upper": str.upper,
    "lower": str.lower,
    "capitalize": str.capitalize,
    "titlecase": str.title,
    "none": lambda x: x,
    "namingcase": naming_case,
}

def normalize_text(
        text: str,
        strip: bool = True,
        capitalization: str = "namingcase",
        remove_dots: bool = True,
        remove_tildes: bool = True,
        remove_invisibles: bool = True,
        remove_weird_spaces: bool = True,
        remove_multi_spaces: bool = True,
        cap_rules: List[str] = None
    ) -> str:
    if text is None:
        return ""
    result = text
    # Remover puntos
    if remove_dots:
        result = result.replace(".", " ")
    # Remueve tildes
    if remove_tildes:
        result = repl_list(result, TILDES)
    # Remover caracteres invisibles
    if remove_invisibles:
        result = rmv_simple(result, INVISIBLES)
    # Reemplazar espacios raros por espacios normales
    if remove_weird_spaces:
        result = repl_fixed(result, WEIRD_SPACES, " ")
    # Remover secciones con multiples espacios consecutivos
    if remove_multi_spaces:
        result = collapse(result, " ")
    # Aplicar stripping (Remover espacios adelante y detras del texto
    if strip:
        result = result.strip()
    # Aplicar capitalizacion de caracteres (Mayusculas y minusculas)
    result = CAPITALIZATION[capitalization](result)
    result = patch_cap(result, cap_rules)
    return result

@dataclass
class Normalizer:
    strip: bool = True
    capitalization: str = "namingcase"
    remove_dots: bool = True
    remove_tildes: bool = True
    remove_invisibles: bool = True
    remove_weird_spaces: bool = True
    remove_multi_spaces: bool = True
    cap_rules: List[str] = None

    def normalize(self, text: str) -> str:
        return normalize_text(text, **self.__dict__)