"""Utils."""

from typing import Optional
from collections import defaultdict
import re
from openpyxl.utils.cell import COORD_RE, SHEET_TITLE
from openpyxl.utils.cell import range_boundaries, get_column_letter


MAX_COL: int = 18278
MAX_ROW: int = 1048576


def resolve_sheet(sheet_name: str) -> str:
    """Resolve sheet name."""
    sheet_name = sheet_name.strip()
    sheet_match = re.match(SHEET_TITLE.strip(), f'{sheet_name}!')
    if sheet_match is None:
        # Internally, sheets are not properly quoted, so consider the entire string.
        return sheet_name
    return sheet_match.group('quoted') or sheet_match.group('notquoted')


def resolve_address(address: str) -> tuple[str, str, str]:
    """Resolve cell address."""
    sheet_name, address_name = address.split('!')
    sheet: str = resolve_sheet(sheet_name)
    coord_match: list[str] = COORD_RE.split(address_name)
    column, row = coord_match[1:3]
    return sheet, column, row


def resolve_ranges(ranges: str, default_sheet: str = 'Sheet1!') -> tuple[str, list[list[str]]]:
    sheet: Optional[str] = None
    range_cells = defaultdict(set)
    for rng in ranges.split(','):
        # Handle sheets in range.
        if '!' in rng:
            sheet_str, rng = rng.split('!')
            rng_sheet = resolve_sheet(sheet_str)
            if sheet is not None and sheet != rng_sheet:
                raise ValueError(
                    'Got multiple different sheets in ranges:'
                    f'{sheet}, {rng_sheet}'
                )
            sheet = rng_sheet
        min_col, min_row, max_col, max_row = range_boundaries(rng)

        min_col = min_col or 1
        min_row = min_row or 1
        max_col = max_col or MAX_COL
        max_row = max_row or MAX_ROW
        for row_idx in range(min_row, max_row + 1):
            for col_idx in range(min_col, max_col + 1):
                range_cells[row_idx].add(col_idx)
    sheet = default_sheet if sheet is None else sheet
    sheet_str = f'{sheet}!' if sheet else ''
    return sheet, [
        [
            f'{sheet_str}{get_column_letter(col_idx)}{row_idx}'
            for col_idx in sorted(row_cells)
        ] for row_idx, row_cells in sorted(range_cells.items())
    ]
