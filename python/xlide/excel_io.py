"""Excel cell read/write via openpyxl."""
from __future__ import annotations

from typing import Any

import openpyxl
from openpyxl.utils.cell import coordinate_from_string, column_index_from_string


def _parse_cell(ref: str) -> tuple[int, int]:
    """Return (row, col) 1-based integers from a cell reference like 'B3'."""
    col_str, row = coordinate_from_string(ref)
    return row, column_index_from_string(col_str)


def read_cells(*, path: str, sheet: str, range: str) -> dict[str, Any]:
    """Return {data: [[...]]} for a cell range in A1 notation."""
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True, keep_vba=True)
    try:
        ws = wb[sheet]
        data: list[list[Any]] = [
            [cell.value for cell in row] for row in ws[range]
        ]
    finally:
        wb.close()
    return {"data": data}


def write_cells(
    *, path: str, sheet: str, startCell: str, data: list[list[Any]]
) -> dict[str, Any]:
    """Write a 2-D array of values starting at startCell and save."""
    start_row, start_col = _parse_cell(startCell)

    wb = openpyxl.load_workbook(path, keep_vba=True)
    try:
        ws = wb[sheet]
        for r_offset, row in enumerate(data):
            for c_offset, value in enumerate(row):
                ws.cell(
                    row=start_row + r_offset,
                    column=start_col + c_offset,
                    value=value,
                )
        wb.save(path)
    finally:
        wb.close()
    return {"ok": True}
