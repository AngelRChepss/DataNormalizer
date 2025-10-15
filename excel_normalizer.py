from typing import List, Dict
import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl import Workbook
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.cell.cell import Cell
from openpyxl.styles import PatternFill, Font
from text_norm import Normalizer


COLOR_NORMALIZED = PatternFill(fill_type="solid", fgColor="FFCCFFFF")
FONT_BASE = Font(bold=False)

def get_header_map(ws: Worksheet):
    header_map =  {
        cell.value : cell.column_letter
        for idx, cell in enumerate(ws[1], start=1)
        if cell.value
    }
    return header_map

def change_cell(cell: Cell, value, pattern: PatternFill = None, font: Font = None):
    cell.value = value
    if pattern:
        cell.fill = pattern
    if font:
        cell.font = font


def normalize_columns(ws: Worksheet, columns: List[str], normalizer: Normalizer, start_row: int = 2) -> None:
    """
    Normaliza una lista de columnas de TEXTO en un Worksheet
    """
    header_map = get_header_map(ws)
    for column in columns:
        cells = ws[header_map[column]]
        for cell in cells:
            if cell.row < start_row:
                continue
            # Normalizer normaliza None a "", porque espera strings.
            # Pero nosotros preferimos quedarnos con None. Mas aun, textos vacios
            # tambien deben ser None.
            if not (cell.value is None):
                result = normalizer.normalize(str(cell.value))
                if cell.value != result:
                    change_cell(cell, result or None, pattern=COLOR_NORMALIZED) # Cambia texto vacio a None

def find_uniques(ws: Worksheet, column: str, exclude_empty: bool = True, sort: bool = False, start_row : int = 2) -> List:
    """
    Encuentra todos los valores unicos en una columna de valores y retorna una lista con ellos.
    La columna puede tener cualquier tipo de datos.
    """
    header_map = get_header_map(ws)
    cells = ws[header_map[column]]

    values = {
        cell.value for cell in cells
        if cell.row >= start_row and not (exclude_empty and cell.value in (None, ""))
    }
    values = sorted(values) if sort else list(values)
    return values

def store_values_in_sheet(wb: Workbook, sheet_name : str, values : Dict) -> None:
    """
    Crea una nueva hoja y agrega todos los datos entregados como columnas.
    """
    ws = wb.create_sheet(title=sheet_name)
    col_idx = 1
    for column_name, column_values in values.items():
        col_letter = get_column_letter(col_idx)
        ws[col_letter + "1"].value = column_name
        row = 2
        for value in column_values:
            ws[col_letter + str(row)].value = value
            row += 1
        col_idx += 1
