from __future__ import annotations

from typing import Sequence

from openpyxl import Workbook
from openpyxl.utils import get_column_letter, quote_sheetname
from openpyxl.workbook.defined_name import DefinedName


def define_named_ranges_for_dict_table(
    *,
    book: Workbook,
    sheet_name: str,
    first_table_row: int,
    first_table_col: int,
    keys: Sequence[str | None],
    workbook_scope: bool,
) -> None:
    """
    For a table with two columns representing the keys and values of a dictionary, define single-cell named ranges for
    the cells in the `values` column.

    Args:
        book: The openpyxl Workbook in which to define the named ranges.
        sheet_name: The sheet on which the dict table exists.
        first_table_row: The number of the top row of the table.
        first_table_col: The number of the left-most column of the table (1=A)
        keys:
            The dictionary keys, in the same order as in the first table column.
            If a key is None, the corresponding row will be skipped, i.e., no named range will be defined for it.
        workbook_scope: Whether to make a workbook-scoped named range (True) or a sheet-scoped named range (False).
    """
    # Values are in the second table column
    col = first_table_col + 1

    for i, key in enumerate(keys):
        if key is None:
            continue

        row = first_table_row + 1 + i
        col_letter = get_column_letter(col)
        name = DefinedName(
            name=key,
            attr_text=f"{quote_sheetname(sheet_name)}!${col_letter}${row}:${col_letter}${row}",
        )

        if workbook_scope:
            book.defined_names.add(name)
        else:
            book[sheet_name].defined_names.add(name)
