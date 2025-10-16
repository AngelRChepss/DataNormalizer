from __future__ import annotations
from typing import List, Dict
import openpyxl
from openpyxl import load_workbook, Workbook
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.cell.cell import Cell
from openpyxl.styles import PatternFill, Font
from text_norm import Normalizer
from norm_utils import check_rut_normalize

class SheetNormalizer:
    FILL_NORMALIZED = PatternFill(fill_type="solid", fgColor="FFCCFFFF")
    FILL_INVALID = PatternFill(fill_type="solid", fgColor="FFFF4444")
    FONT_BASE = Font(bold=False)

    def __init__(self, worksheet: Worksheet, wb_normalizer: BookNormalizer):
        self.ws = worksheet
        self._max_row: int = None
        self._header_map = None
        self._max_column: int = None
        self.wb_normalizer = wb_normalizer

    def recalculate_header_map(self):
        self._header_map = {
            cell.value: cell.column_letter
            for idx, cell in enumerate(self.ws[1], start=1)
            if cell.value
        }

    @property
    def header_map(self) -> Dict[str, str]:
        if self._header_map is None:
            self.recalculate_header_map()
        return self._header_map

    def recalculate_max_row(self):
        max_row = self.ws.max_row
        while max_row > 0 and all(cell.value is None for cell in self.ws[max_row]):
            max_row -= 1
        self._max_row = max_row

    @property
    def max_row(self) -> int:
        if self._max_row is None:
            self.recalculate_max_row()
        return self._max_row

    def recalculate_max_column(self):
        self._max_column = len([cell for cell in self.ws[1] if cell.value])

    @property
    def max_column(self) -> int:
        if self._max_column is None:
            self.recalculate_max_column()
        return self._max_column

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
                        SheetNormalizer.change_cell(cell, result or None, pattern=SheetNormalizer.FILL_NORMALIZED)

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
            cell.fill = SheetNormalizer.FILL_INVALID
            invalid_count += 1
        return invalid_count

    def normalize_ruts(self, column: str, start_row: int = 2) -> int:
        cells = self.ws[self.header_map[column]]
        invalid_count = 0
        for cell in cells:
            if cell.row < start_row:
                continue
            if cell.value:
                valid, norm = check_rut_normalize(str(cell.value))
                if valid:
                    if norm != cell.value:
                        SheetNormalizer.change_cell(cell, norm, pattern=SheetNormalizer.FILL_NORMALIZED)
                    continue
            cell.fill = SheetNormalizer.FILL_INVALID
            invalid_count += 1
        return invalid_count

    def write_values(self, values: Dict) -> None:
        """
        Agrega todos los datos entregados como columnas.
        """
        col_idx = self.max_column + 1
        for column_name, column_values in values.items():
            col_letter = get_column_letter(col_idx)
            self.ws[col_letter + "1"].value = column_name
            row = 2
            for value in column_values:
                self.ws[col_letter + str(row)].value = value
                row += 1
            col_idx += 1
        self.recalculate_max_row()
        self.recalculate_header_map()
        self.recalculate_max_column()

    def copy_column(self, read_column: str, write_column: str | None = None, write_ws: Worksheet | None = None) -> None:
        """
        Copia una columna de una hoja a otra hoja (o la misma hoja)
        """
        write_ws = write_ws or self.ws
        write_column = write_column or read_column
        # TODO: FUNC
        pass

class BookNormalizer:
    def __init__(self, file_name: str):
        self.wb: Workbook = load_workbook(file_name)
        self.ws_norms = {sheet: SheetNormalizer(self.wb[sheet], self) for sheet in self.wb.sheetnames}
        self.current_norm = self.ws_norms[self.wb.sheetnames[0]]

    def keep_sheets(self, sheets: List[str] | None = None) -> None:
        sheets = sheets or self.wb.sheetnames
        wb_sheets = self.wb.sheetnames.copy()
        for sheet in wb_sheets:
            if sheet not in sheets:
                del self.wb[sheet]

    def save(self, file_name: str) -> None:
        self.wb.save(file_name)

    def create_sheet(self, sheet_name: str) -> None:
        ws = self.wb.create_sheet(sheet_name)
        self.ws_norms[sheet_name] = SheetNormalizer(ws, self)

    def join_columns(self, target_worksheet: str, columns: List[str], target_name: str, join_character: str = " ") -> None:
        """
        Combina varias columnas en una sola que escribe en otra hoja.
        """
        cols = [self.header_map[col] for col in columns]
        max_row = self.max_row
        tgt_wsn = self.ws_norms[target_worksheet]

        values = []
        for row in range(2, max_row+1):
            data = []
            for col in cols:
                data.append(self.ws[f"{col}{row}"].value)
            values.append(join_character.join(str(x or "") for x in data if x).strip())

        tgt_wsn.write_values({target_name: values})

    def activate_sheet(self, sheet_name: str) -> None:
        self.current_norm = self.ws_norms[sheet_name]

    @property
    def sheet(self) -> SheetNormalizer:
        """Acceso explícito al SheetNormalizer actual (para autocompletado)."""
        return self.current_norm

    def __getattr__(self, name):
        """
        Delegar metodos y atributos al SheetNormalizer actual.
        Mejora mantenibilidad de las clases.
        """
        if self.current_norm and hasattr(self.current_norm, name):
            return getattr(self.current_norm, name)
        raise AttributeError(f"'BookNormalizer' object has no attribute '{name}'")
