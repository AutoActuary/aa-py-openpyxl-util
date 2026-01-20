from __future__ import annotations

import re
from collections import OrderedDict
from itertools import chain
from logging import getLogger
from typing import Generator, Optional, List, Any, Tuple, Dict, TYPE_CHECKING, Literal

from ._data_util import data_to_dicts, skip_empty_rows
from ._find_table import find_table
from ._iter_tables import iter_list_object_tables, iter_named_range_tables

if TYPE_CHECKING:
    from ._typing import TableCells
    from openpyxl import Workbook
    from openpyxl.cell import Cell

logger = getLogger(__name__)


def read_table(
    *,
    book: "Workbook",
    table_name: str,
    columns: List[str] | None = None,
    ci: (
        bool | Literal["warn"]
    ) = False,  # TODO: Make this required in the next major version.
) -> Generator[Dict[str, Any], None, None]:
    """
    Read a table from a workbook and yield its rows as dictionaries.

    Args:
        book: The workbook, opened using openpyxl, from which to read the table.
        table_name: The name of the table (ListObject or named range) to read.
        columns:
            Optional list of column names to extract.
            If not given, all columns are extracted.
        ci:
            Whether the table name lookup should be case-insensitive.
            When this is "warn", a warning is logged when the provided case does not match the actual case.

    Returns:
        A generator of dictionaries mapping column names to cell values for
        each non-header row in the table.
    """
    sheet, table_range = find_table(book=book, name=table_name, ci=ci)
    return data_to_dicts(
        data=sheet[table_range],
        columns=columns,
        value_callback=get_cell_value,
        header_callback=get_cell_value_as_str,
    )


def read_dict_table(
    *,
    book: "Workbook",
    table_name: str,
    key_column: str = "Key",  # TODO: Make this required in the next major version.
    value_column: str = "Value",  # TODO: Make this required in the next major version.
    ci: (
        bool | Literal["warn"]
    ) = False,  # TODO: Make this required in the next major version.
) -> Dict[str, Any]:
    """
    Read a two-column table from a workbook and return it as a dictionary.

    Args:
        book: The openpyxl workbook object from which to read the table.
        table_name: The name of the table (ListObject or named range) to read.
        key_column: The name of the column whose values to use as dictionary keys.
        value_column: The name of the column whose values to use as dictionary values.
        ci:
            Whether the table name lookup should be case-insensitive.
            When this is "warn", a warning is logged when the provided case does not match the actual case.

    Returns:
        A dictionary that maps each value from `key_column` in the specified table
        to the corresponding value from `value_column`.
    """
    data = read_table(
        book=book,
        table_name=table_name,
        columns=[key_column, value_column],
        ci=ci,
    )
    return {row[key_column]: row[value_column] for row in data}


def extract_data_from_numbered_tables(
    book: "Workbook",
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


def get_cell_value_as_str(cell: "Cell") -> Any:
    return str(get_cell_value(cell))


def get_cell_value(cell: "Cell") -> Any:
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
    book: "Workbook",
    base_name: str,
) -> List[Tuple[str, "TableCells"]]:
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

    def gen() -> Generator[Tuple[int, str, "TableCells"], None, None]:
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
