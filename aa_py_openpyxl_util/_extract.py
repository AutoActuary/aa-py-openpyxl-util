from __future__ import annotations

import logging
import re
from collections import OrderedDict
from itertools import chain
from typing import Generator, Optional, List, Any, Tuple

from openpyxl import Workbook
from openpyxl.cell import Cell

from ._data_util import data_to_dicts, skip_empty_rows
from ._iter_tables import iter_list_object_tables, iter_named_range_tables
from ._typing import TableCells

logger = logging.Logger(__name__)


def extract_data_from_numbered_tables(
    book: Workbook,
    base_name: str,
    columns: Optional[List[str]] = None,
) -> Generator[OrderedDict[str, Any], None, None]:
    """
    Stack multiple numbered tables in order, and extract data from all of them.

    Args:
        book: The workbook, opened using openpyxl.
        base_name: See `get_numbered_tables`.
        columns: The columns to extract. If not given, all columns will be extracted.

    Returns:
        A generator of ordered, case-insensitive dictionaries.
    """
    for name, cells in get_numbered_tables(book=book, base_name=base_name):
        yield from skip_empty_rows(
            data_to_dicts(
                data=cells,
                columns=columns,
                value_callback=get_cell_value,
                header_callback=get_cell_value_as_str,
            )
        )


def get_cell_value_as_str(cell: Cell) -> Any:
    return str(get_cell_value(cell))


def get_cell_value(cell: Cell) -> Any:
    value = cell.value

    # Workaround for:
    # - https://github.com/AutoActuary/aa-py-autory-normalize/issues/4
    # - https://foss.heptapod.net/openpyxl/openpyxl/-/issues/1410
    # - https://foss.heptapod.net/openpyxl/openpyxl/-/issues/1975
    if isinstance(value, str) and "_x000D_\n" in value:
        value = value.replace("_x000D_\n", "\n")

        # Emit a warning, because the replacement is unsafe. The user should fix the Excel file.
        logger.warning(
            f"Cell {cell.coordinate} contains a carriage return. "
            f"This is not supported. Please replace the carriage return with a newline."
        )

    return value


def get_numbered_tables(
    book: Workbook,
    base_name: str,
) -> List[Tuple[str, TableCells]]:
    """
    Get a list of all tables that match ``name123`` where name is the given base name and 123 is any integer.
    The list is sorted in ascending numerical order.

    Args:
        book: The Excel workbook, opened by xlwings.
        base_name: The table base name.

    Returns:
        List of tables.

    Examples:

        get_numbered_tables(book, "MyTable")
        ["MyTable", "MyTable3", "MyTable15"]
    """
    re_name = re.escape(base_name.casefold())
    pattern = re.compile(rf"^{re_name}(\d+)?$")

    def gen() -> Generator[Tuple[int, str, TableCells], None, None]:
        lo_names_and_cells = (
            (table.name, sheet[table.ref])
            for sheet, table in iter_list_object_tables(
                book=book, exclude_list_objects=[], exclude_sheets=[]
            )
        )

        nr_names_and_cells = (
            (table_name, sheet[table_range])
            for sheet, table_name, table_range in iter_named_range_tables(
                book=book, exclude_names=[], exclude_sheets=[]
            )
        )

        for name, cells in chain(
            lo_names_and_cells,
            nr_names_and_cells,
        ):
            match = pattern.fullmatch(name.casefold())
            if match:
                num = int(match.group(1) or "-1")
                yield num, name, cells

    tables = sorted(list(gen()))

    return [(name, cells) for i, name, cells in tables]
