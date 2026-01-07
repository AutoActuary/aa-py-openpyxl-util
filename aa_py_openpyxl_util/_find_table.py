from typing import Tuple, TYPE_CHECKING

if TYPE_CHECKING:
    from openpyxl import Workbook
    from openpyxl.workbook.defined_name import DefinedName
    from openpyxl.worksheet.table import Table
    from openpyxl.worksheet.worksheet import Worksheet


def find_table(*, book: "Workbook", name: str) -> Tuple["Worksheet", str]:
    """
    Find a table in the given workbook. The table can be a named range or a ListObject.
    """
    try:
        sheet, table_range = find_named_range_by_name(book=book, name=name)
        table_type = "Named range"
    except KeyError as e1:
        try:
            sheet, table = find_list_object_by_name(book=book, name=name)
            table_range = table.ref
            table_type = "ListObject"
        except KeyError as e2:
            raise KeyError(f"Table `{name}` not found: {e1.args[0]} {e2.args[0]}")

    from openpyxl.utils import range_boundaries

    min_col, min_row, max_col, max_row = range_boundaries(table_range)
    n_rows = max_row - min_row + 1
    n_cols = max_col - min_col + 1

    if n_rows < 2:
        raise KeyError(f"{table_type} `{name}` found, but it has fewer than 2 rows.")
    if n_cols < 1:
        raise KeyError(f"{table_type} `{name}` found, but it has no columns.")

    return sheet, table_range


def find_named_range_by_name(*, book: "Workbook", name: str) -> Tuple["Worksheet", str]:
    """
    Find a named range by name.

    Args:
        book: The book to search through.
        name: The name of the ListObject.

    Returns:
        A tuple of (sheet, range).
    """
    try:
        defined_name: "DefinedName" = book.defined_names[name]
    except KeyError as e:
        raise KeyError(f"Named range `{name}` not found.") from e

    try:
        destinations = list(defined_name.destinations)
    except AttributeError:
        raise KeyError(f"Named range `{name}` found, but it has no destinations.")

    if len(destinations) != 1:
        raise KeyError(f"Named range `{name}` found, but it has multiple destinations.")

    sheet_name, table_range = destinations[0]
    return book[sheet_name], table_range


def find_list_object_by_name(
    *, book: "Workbook", name: str
) -> Tuple["Worksheet", "Table"]:
    """
    Find a list object by name.

    Args:
        book: The book to search through.
        name: The name of the ListObject.

    Returns:
        A tuple of (sheet, range).
    """
    sheet: "Worksheet"
    for sheet in book.worksheets:
        if name in sheet.tables:
            return sheet, sheet.tables[name]

    raise KeyError(f"ListObject `{name}` not found.")
