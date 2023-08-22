"""
Type definitions for the tables that have been written using openpyxl.
This is mainly to make the code more readable.
"""

from typing import Dict, Tuple

import openpyxl.worksheet.table

WrittenTableCoordinates = Tuple[int, int]
"""
The co-ordinates of the top-left cell of the table.
"""

WrittenTableInfo = Tuple[WrittenTableCoordinates, openpyxl.worksheet.table.Table]
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
