"""
Utilities that build on top of `openpyxl`.
"""

from typing import TYPE_CHECKING

from ._cells import process_cells, get_cell_values
from ._context import safe_load_workbook, changed_builtin_number_formats
from ._data_validation import set_data_validation_input_message
from ._extract import extract_data_from_numbered_tables, read_table, read_dict_table
from ._find_table import find_table
from ._iter_tables import iter_named_range_tables, iter_list_object_tables
from ._named_ranges import define_named_ranges_for_dict_table
from ._workarounds import save_workbook_workaround, remove_atexit_permission_error
from ._write_only import (
    FormattedCell,
    TableInfo,
    write_tables_side_by_side,
    write_tables_side_by_side_over_multiple_sheets,
)

if TYPE_CHECKING:
    from ._typing import (
        TableCells,
        WrittenTables,
        WrittenTablesInSheet,
        WrittenTableInfo,
        WrittenTableCoordinates,
    )
