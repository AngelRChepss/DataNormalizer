from typing import List
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.cell.cell import Cell
from openpyxl.styles import PatternFill, Font
from text_norm import Normalizer


COLOR_NORMALIZED = PatternFill(fill_type="solid", fgColor="FFCCFFFF")
FONT_BASE = Font(bold=False)

def change_cell(cell: Cell, value, pattern: PatternFill = None, font: Font = None):
    cell.value = value
    if pattern:
        cell.fill = pattern
    if font:
        cell.font = font


def normalize_columns(ws: Worksheet, columns: List[str], normalizer: Normalizer) -> None:
    """
    Normaliza una lista de columnas de TEXTO en un Worksheet
    """
    for column in columns:
        cells = ws[column]
        for cell in cells:
            # Normalizer normaliza None a "", porque espera strings.
            # Pero nosotros preferimos quedarnos con None. Mas aun, textos vacios
            # tambien deben ser None.
            if not (cell.value is None):
                result = normalizer.normalize(str(cell.value))
                cell.value = result or None # Cambia texto vacio a None

def find_uniques(ws: Worksheet, column: str, exclude_empty: bool = True, sort: bool = False) -> List:
    """
    Encuentra todos los valores unicos en una columna de valores y retorna una lista con ellos.
    La columna puede tener cualquier tipo de datos.
    """
    cells = ws[column]

    values = {
        cell.value for cell in cells
        if not (exclude_empty and cell.value in (None, ""))
    }
    values = sorted(values) if sort else list(values)
    return values

