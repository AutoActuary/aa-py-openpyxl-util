"""
Utilities that build on top of `openpyxl`.
"""

from ._cells import process_cells, get_cell_values
from ._context import safe_load_workbook, changed_builtin_number_formats
from ._extract import extract_data_from_numbered_tables
from ._find_table import find_table
from ._iter_tables import iter_named_range_tables, iter_list_object_tables
from ._typing import TableCells
from ._workaround import save_workbook_workaround
from ._write_only import (
    FormattedCell,
    TableInfo,
    write_tables_side_by_side,
    write_tables_side_by_side_over_multiple_sheets,
    define_named_ranges_for_dict_table,
)
from ._written_tables_types import (
    WrittenTables,
    WrittenTablesInSheet,
    WrittenTableInfo,
    WrittenTableCoordinates,
)
