from typing import Dict, List, Set
from difflib import SequenceMatcher

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
