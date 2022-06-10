"""Module for utils."""

import re
import uuid
from collections import defaultdict
from string import ascii_uppercase
from typing import Optional

from openpyxl.utils.cell import COORD_RE, SHEET_TITLE
from openpyxl.utils.cell import get_column_letter, range_boundaries

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


def resolve_ranges(ranges: str, default_sheet: str = 'Sheet1') -> tuple[str, list[list[str]]]:
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


def col2num(col: Optional[str]) -> int:
    if not col:
        raise Exception('Column may not be empty')

    tot = 0
    for i, c in enumerate([c for c in col[::-1] if c != '$']):
        if c == '$':
            continue
        tot += (ord(c) - 64) * 26 ** i
    return tot


def num2col(num: int):
    if num < 1:
        raise Exception(f'Number must be larger than 0: {num}')
    s = ''
    q = num
    while q > 0:
        (q, r) = divmod(q, 26)
        if r == 0:
            q = q - 1
            r = 26
        s = ascii_uppercase[r - 1] + s
    return s


def init_uuid():
    """Default factory to initialise Formula.ranges."""
    return uuid.uuid4()
