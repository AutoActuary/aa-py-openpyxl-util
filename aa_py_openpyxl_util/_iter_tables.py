from typing import Collection, Generator, Tuple, Pattern, Union

from openpyxl import Workbook
from openpyxl.utils import range_boundaries
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.table import Table
from openpyxl.worksheet.worksheet import Worksheet


def iter_named_range_tables(
    *,
    book: Workbook,
    exclude_names: Collection[Union[Pattern[str], str]],
    exclude_sheets: Collection[Union[Pattern[str], str]],
) -> Generator[Tuple[Worksheet, str, str], None, None]:
    """
    Iterate over named range tables in the workbook.

    TODO: Change `Union[Pattern[str], str]` to `Pattern[str]` in the next major version.

    Args:
        book: The workbook containing the named range tables to iterate over.
        exclude_names: Tables with names matching any of these patterns will be excluded.
        exclude_sheets: Tables in sheets having names matching any of these patterns will be excluded.

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
            if any_match(exclude_sheets, sheet_name.casefold()):
                continue

            if any_match(exclude_names, defined_name.name.casefold()):
                continue

            if not is_table_range(table_range):
                continue

            yield book[sheet_name], defined_name.name, table_range


def iter_list_object_tables(
    *,
    book: Workbook,
    exclude_list_objects: Collection[Union[Pattern[str], str]],
    exclude_sheets: Collection[Union[Pattern[str], str]],
) -> Generator[Tuple[Worksheet, Table], None, None]:
    """
    Iterate over list object tables in the workbook.

    TODO: Change `Union[Pattern[str], str]` to `Pattern[str]` in the next major version.

    Args:
        book: The workbook containing the list object tables to iterate over.
        exclude_list_objects: Tables with names matching any of these patterns will be excluded.
        exclude_sheets: Tables in sheets having names matching any of these patterns will be excluded.

    Returns:
        A generator of tuples like (sheet, table).
    """
    sheet: Worksheet
    for sheet in book.worksheets:
        if any_match(exclude_sheets, sheet.title):
            continue

        for table_name in sheet.tables.keys():
            if any_match(exclude_list_objects, table_name):
                continue

            table = sheet.tables[table_name]
            if not is_table_range(table.ref):
                continue

            yield sheet, table


def any_match(patterns: Collection[Union[Pattern[str], str]], string: str) -> bool:
    """
    Check if any of the given patterns match the given string.

    TODO: Change `Union[Pattern[str], str]` to `Pattern[str]` in the next major version.

    Args:
        patterns: The patterns to check.
        string: The string to check.

    Returns:
        True if any of the patterns match the string, False otherwise.
    """
    return any(
        (
            pattern.search(string)
            if isinstance(pattern, Pattern)
            else string.casefold() == pattern.casefold()
        )
        for pattern in patterns
    )


def is_table_range(table_range: str) -> bool:
    """
    Check if the given range can be a table, i.e. it has at least one column and at least two rows.

    Examples:
        >>> is_table_range('A1:A1')
        False

        >>> is_table_range('A1:A2')
        True

        >>> is_table_range('A1:B1')
        False

        >>> is_table_range('A1:B2')
        True

        >>> is_table_range('A1')
        False

        >>> is_table_range('')
        False

        >>> is_table_range('???')
        False
    """
    try:
        min_col, min_row, max_col, max_row = range_boundaries(table_range)
    except ValueError:
        return False

    if min_col is None or min_row is None or max_col is None or max_row is None:
        return False

    n_rows = max_row - min_row + 1
    n_cols = max_col - min_col + 1

    return bool(n_rows >= 2) and bool(n_cols >= 1)
