from typing import List, Dict
import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl import Workbook
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.cell.cell import Cell
from openpyxl.styles import PatternFill, Font
from text_norm import Normalizer
from norm_utils import check_rut_normalize

class ExcelNormalizer:
    FILL_NORMALIZED = PatternFill(fill_type="solid", fgColor="FFCCFFFF")
    FILL_INVALID = PatternFill(fill_type="solid", fgColor="FFFF4444")
    FONT_BASE = Font(bold=False)

    def __init__(self, worksheet: Worksheet):
        self.ws = worksheet
        self._max_row = None
        self._header_map = None

    @property
    def header_map(self) -> Dict[str, str]:
        if self._header_map is None:
            self._header_map = {
                cell.value : cell.column_letter
                for idx, cell in enumerate(self.ws[1], start=1)
                if cell.value
            }
        return self._header_map

    @property
    def max_row(self) -> int:
        if self._max_row is None:
            max_row = self.ws.max_row
            while max_row > 0 and all(cell.value is None for cell in self.ws[max_row]):
                max_row -= 1
            self._max_row = max_row
        return self._max_row

    @staticmethod
    def change_cell(cell: Cell, value, pattern: PatternFill | None = None, font: Font | None = None):
        cell.value = value
        if pattern:
            cell.fill = pattern
        if font:
            cell.font = font

    def normalize_columns(self, columns: List[str], normalizer: Normalizer, start_row: int = 2) -> None:
        """
        Normaliza una lista de columnas de TEXTO en un Worksheet
        """
        for column in columns:
            cells = self.ws[self.header_map[column]]
            for cell in cells:
                if cell.row < start_row:
                    continue
                # Normalizer normaliza None a "", porque espera strings.
                # Pero nosotros preferimos quedarnos con None. Más aún, textos vacíos
                # también deben ser None.
                if not (cell.value is None):
                    result = normalizer.normalize(str(cell.value))
                    if cell.value != result:
                        # Cambia texto vacío a None
                        ExcelNormalizer.change_cell(cell, result or None, pattern=ExcelNormalizer.FILL_NORMALIZED)

    def find_uniques(self, column: str, exclude_empty: bool = True, sort: bool = False, start_row : int = 2) -> List:
        """
        Encuentra todos los valores unicos en una columna de valores y retorna una lista con ellos.
        La columna puede tener cualquier tipo de datos.
        """
        cells = self.ws[self.header_map[column]]

        values = {
            cell.value for cell in cells
            if cell.row >= start_row and not (exclude_empty and cell.value in (None, ""))
        }
        values = sorted(values) if sort else list(values)
        return values

    def highlight_invalid_ruts(self, column: str) -> int:
        cells = self.ws[self.header_map[column]]
        invalid_count = 0
        for cell in cells:
            if cell.value:
                valid, norm = check_rut_normalize(str(cell.value))
                if valid:
                    continue
            cell.fill = ExcelNormalizer.FILL_INVALID
            invalid_count += 1
        return invalid_count

    def normalize_ruts(self, column: str) -> int:
        cells = self.ws[self.header_map[column]]
        invalid_count = 0
        for cell in cells:
            if cell.value:
                valid, norm = check_rut_normalize(str(cell.value))
                if valid:
                    if norm != cell.value:
                        ExcelNormalizer.change_cell(cell, norm, pattern=ExcelNormalizer.FILL_NORMALIZED)
                    continue
            cell.fill = ExcelNormalizer.FILL_INVALID
            invalid_count += 1
        return invalid_count

    @staticmethod
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

    def join_columns(self, write_ws: Worksheet, columns: List[str], target_name: str, join_character: str = " ") -> None:
        """
        Combina varias columnas en una sola que escribe en otra hoja.
        """
        # TODO: FUNC
        pass

    def copy_column(self, read_column: str, write_column: str | None = None, write_ws: Worksheet | None = None) -> None:
        """
        Copia una columna de una hoja a otra hoja (o la misma hoja)
        """
        write_ws = write_ws or self.ws
        write_column = write_column or read_column
        # TODO: FUNC
        pass
