"""
Type definitions to make code more readable.
"""

from __future__ import annotations

from typing import Dict, Tuple, TYPE_CHECKING

if TYPE_CHECKING:
    from openpyxl.cell import Cell
    from openpyxl.worksheet.table import Table

    TableCells = Tuple[Tuple[Cell, ...], ...]
    """
    A 2D tuple of table cells.
    """

    WrittenTableCoordinates = Tuple[int, int]
    """
    The co-ordinates of the top-left cell of the table.
    """

    WrittenTableInfo = Tuple[WrittenTableCoordinates, Table]
    """
    Info about a table that has been written using openpyxl.
    """

    WrittenTablesInSheet = Dict[str, WrittenTableInfo]
    """
    Tables that have been written to the same sheet, keyed by table name.
    """

    WrittenTables = Dict[str, WrittenTablesInSheet]
    """
    Tables that have been written to the same workbook, keyed by sheet name.
    """
