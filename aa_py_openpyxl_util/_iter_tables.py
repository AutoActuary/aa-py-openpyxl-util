from typing import Collection, Generator, Tuple

from openpyxl import Workbook
from openpyxl.utils import range_boundaries
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.table import Table
from openpyxl.worksheet.worksheet import Worksheet


def iter_named_range_tables(
    *,
    book: Workbook,
    exclude_names: Collection[str],
    exclude_sheets: Collection[str],
) -> Generator[Tuple[Worksheet, str, str], None, None]:
    """
    Iterate over named range tables in the workbook.

    Args:
        book: The workbook containing the named range tables to iterate over.
        exclude_names: Tables with these names will be excluded. Must be provided in lower case.
        exclude_sheets: Tables in these sheets will be excluded. Must be provided in lower case.

    Returns:
        A generator of tuples like (sheet, name, range).
    """
    defined_name: DefinedName
    for defined_name in book.defined_names.values():

        try:
            destinations = list(defined_name.destinations)
        except AttributeError:
            destinations = []

        for sheet_name, table_range in destinations:
            if sheet_name.casefold() in exclude_sheets:
                continue

            if defined_name.name.casefold() in exclude_names:
                continue

            if not is_table_range(table_range):
                continue

            yield book[sheet_name], defined_name.name, table_range


def iter_list_object_tables(
    *,
    book: Workbook,
    exclude_list_objects: Collection[str],
    exclude_sheets: Collection[str],
) -> Generator[Tuple[Worksheet, Table], None, None]:
    """
    Iterate over list object tables in the workbook.

    Args:
        book: The workbook containing the list object tables to iterate over.
        exclude_list_objects: Tables with these names will be excluded. Must be provided in lower case.
        exclude_sheets: Tables in these sheets will be excluded. Must be provided in lower case.

    Returns:
        A generator of tuples like (name, cells).
    """
    sheet: Worksheet
    for sheet in book.worksheets:
        if sheet.title.casefold() in exclude_sheets:
            continue

        for table_name in sheet.tables.keys():
            if table_name.casefold() in exclude_list_objects:
                continue

            table = sheet.tables[table_name]
            if not is_table_range(table.ref):
                continue

            yield sheet, table


def is_table_range(table_range: str) -> bool:
    """
    Check if the given range can be a table, i.e. it has at least one column and at least two rows.
    """
    min_col, min_row, max_col, max_row = range_boundaries(table_range)
    n_rows = max_row - min_row + 1
    n_cols = max_col - min_col + 1

    return bool(n_rows >= 2) and bool(n_cols >= 1)
