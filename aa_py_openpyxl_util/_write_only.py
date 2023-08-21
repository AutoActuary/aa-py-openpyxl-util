"""
Utilities for working with write-only openpyxl workbooks.
"""
from __future__ import annotations

import logging
import warnings
from dataclasses import dataclass, field
from itertools import zip_longest
from typing import Optional, Any, Sequence, Generator, List, Tuple, Dict, Iterable

from openpyxl import Workbook
from openpyxl.cell import WriteOnlyCell, Cell
from openpyxl.utils import get_column_letter, quote_sheetname
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet._write_only import WriteOnlyWorksheet
from openpyxl.worksheet.formula import ArrayFormula
from openpyxl.worksheet.table import Table, TableStyleInfo

logger = logging.getLogger(__name__)


@dataclass
class FormattedCell:
    """
    Custom class to hold cell data separate from a sheet.

    This is required because:

        - openpyxl does not allow creating a WriteOnlyCell with formatting if it's not attached to a sheet.
        - openpyxl's `ArrayFormula` class needs a `ref` value, which can only be derived when we know the cell position.
    """

    value: Any
    """
    The cell value. If it's a string starting with `=`, openpyxl interprets it as a formula.
    """

    number_format: Optional[str] = None
    """
    The cell's number format. Optional.
    """

    array: bool = False
    """
    Whether this is an array formula. This sets `t="array"` in the `f` tag in XML.
    """

    def create_openpyxl_cell(
        self,
        sheet: WriteOnlyWorksheet,
        ref: str,
    ) -> Cell:
        """

        Args:
            sheet:
                The sheet into which the cell will be written.
            ref:
                The value of the `ref` attribute for the `f` tag in XML. Required when `self.array==True`. For array
                formulas with scalar results, this should refer to the cell containing the formula. For array formulas
                with multi-celled results, this should refer to the entire range of cells that will contain the results.

        Returns:
            An openpyxl `WriteOnlyCell` instance.
        """
        value = (
            ArrayFormula(
                ref=ref,
                text=self.value,
            )
            if self.array
            else self.value
        )

        cell: Cell = WriteOnlyCell(ws=sheet, value=value)

        if self.number_format:
            # noinspection PyUnresolvedReferences,PyDunderSlots
            cell.number_format = self.number_format

        return cell


default_table_style = TableStyleInfo(
    name="TableStyleMedium2",
    showFirstColumn=False,
    showLastColumn=False,
    showRowStripes=True,
    showColumnStripes=False,
)


@dataclass
class TableInfo:
    name: str
    """
    The table name
    """

    column_names: Sequence[str]
    """
    The column names
    """

    rows: Sequence[Sequence[FormattedCell]]
    """
    The table rows.
    """

    pre_rows: Sequence[Sequence[FormattedCell]] = field(default_factory=list)
    """
    Rows to write outside the table, above the header, but below the name and description.
    """

    style: Optional[TableStyleInfo] = field(default=default_table_style)
    """
    The table style
    """

    description: str = field(default="")
    """
    A table description to write below the table name.
    """


def write_tables_side_by_side(
    *,
    book: Workbook,
    sheet_name: str,
    tables: Sequence[TableInfo],
    row_margin: int,
    col_margin: int,
    write_captions: bool,
    write_pre_rows: bool,
) -> Dict[str, Tuple[Tuple[int, int], Table]]:
    """
    Create a new sheet containing one or more tables, stacked horizontally.

    See https://openpyxl.readthedocs.io/en/stable/worksheet_tables.html#creating-a-table
    See https://openpyxl.readthedocs.io/en/stable/worksheet_tables.html#manually-adding-column-headings

    Args:
        book: A write-only workbook in which to create the sheet and table.
        sheet_name: The name of the new sheet. Also the name of the table.
        tables: A sequence of table info objects.
        row_margin: The number of empty rows to leave above each table.
        col_margin: The number of empty columns to leave to the left of each table.
        write_captions: Whether to write the table name and description above the table. This shifts the table down.
        write_pre_rows: Whether to write the pre_rows (below the name and description, but above the table header).

    Returns:
        Dict of info about each written table, keyed by table name.
        Each value is a tuple containing:
        - coordinates (e.g. (2,3) which means cell C2)
        - openpyxl table object
    """
    sheet = book.create_sheet(title=sheet_name)

    # Write rows
    for i_row, row in enumerate(
        stack_table_rows_side_by_side(
            tables=tables,
            row_margin=row_margin,
            col_margin=col_margin,
            write_captions=write_captions,
            write_pre_rows=write_pre_rows,
        ),
        start=1,
    ):
        sheet.append(
            (
                cell.create_openpyxl_cell(
                    sheet=sheet,
                    ref=f"{get_column_letter(i_col)}{i_row}",
                )
                for i_col, cell in enumerate(row, start=1)
            )
        )

    # Define ListObjects
    results: Dict[str, Tuple[Tuple[int, int], Table]] = {}

    if not len(tables):
        return results

    first_row = (
        1
        + row_margin
        + (2 if write_captions else 0)
        + (max(len(t.pre_rows) for t in tables) if write_pre_rows else 0)
    )
    first_column = 1 + col_margin
    for t in tables:
        width = len(t.column_names)
        if width < 1:
            raise ValueError(f"Can't create table '{t.name}' with zero columns.")

        coords = first_row, first_column
        lo = define_list_object(
            sheet=sheet,
            first_column=first_column,
            first_row=first_row,
            name=t.name,
            column_names=t.column_names,
            n_data_rows=len(t.rows),
            style=t.style,
        )
        results[t.name] = (coords, lo)

        first_column += col_margin + width

    return results


def distribute_tables_over_multiple_sheets(
    *,
    tables: Iterable[TableInfo],
    max_sheet_width: int,
    left_margin: int,
    gutter: int,
) -> List[List[TableInfo]]:
    """
    Distribute tables over multiple sheets, so that the total width of each sheet is within bounds.

    Args:
        tables: The table to distribute.
        max_sheet_width: The maximum number of columns in each sheet.
        left_margin: The number of columns to leave open left of the first table.
        gutter: The number of columns to keep open between tables.

    Returns:
        A list of lists of tables. Each inner list represents a sheet.
    """
    result: List[List[TableInfo]] = []
    current_column = 1 + left_margin  # 1-indexed, like the Excel, i.e., column A is 1.
    current_sheet = 0
    for table in tables:
        if len(result) < 1:
            result.append([])

        table_width = len(table.column_names)
        if len(result[current_sheet]) != 0:
            current_column += gutter

        if current_column + table_width - 1 > max_sheet_width:
            # This table does not fit in the current sheet.

            if len(result[current_sheet]) == 0:
                # This is the first table in the sheet.
                # This table will never fit.
                raise ValueError(
                    f"Table `{table.name}` is too wide to fit in a sheet with maximum width {max_sheet_width} "
                    f"and a left margin of {left_margin} columns."
                )

            # Move to the next sheet.
            current_sheet += 1
            current_column = 1 + left_margin
            result.append([])

        # Add the table to the current sheet.
        result[current_sheet].append(table)
        current_column += table_width

    return result


def stack_table_rows_side_by_side(
    tables: Sequence[TableInfo],
    row_margin: int,
    col_margin: int,
    write_captions: bool,
    write_pre_rows: bool,
) -> Generator[List[FormattedCell], None, None]:
    """
    Iterate over the cells of multiple tables simultaneously, and yield one row at a time,

    Args:
        tables: A sequence of table info objects.
        row_margin: The number of empty rows to leave above each table.
        col_margin: The number of empty columns to leave to the left of each table.
        write_captions: Whether to write the table name and description above the table.
        write_pre_rows: Whether to write the pre_rows (below the name and description, but above the table header).

    Returns:
        A generator that yields one row at a time, in such a way that a sheet can be written from this data,
        top to bottom, without ever going back to a previous row.
    """
    if row_margin < 0:
        raise ValueError("Row margin must be a positive integer.")
    if col_margin < 0:
        raise ValueError("Column margin must be a positive integer.")

    def table_name_row() -> Generator[FormattedCell, None, None]:
        for t in tables:
            yield from [FormattedCell(None)] * col_margin
            yield FormattedCell(t.name)
            yield from [FormattedCell(None)] * (len(t.column_names) - 1)

    def table_description_row() -> Generator[FormattedCell, None, None]:
        for t in tables:
            yield from [FormattedCell(None)] * col_margin
            yield FormattedCell(t.description)
            yield from [FormattedCell(None)] * (len(t.column_names) - 1)

    def header_row() -> Generator[FormattedCell, None, None]:
        for t in tables:
            yield from [FormattedCell(None)] * col_margin
            for c in t.column_names:
                yield FormattedCell(c)

    widths = [len(t.column_names) for t in tables]

    def row(
        data: Sequence[Optional[Sequence[FormattedCell]]],
    ) -> Generator[FormattedCell, None, None]:
        for w, d in zip(widths, data):
            yield from [FormattedCell(None)] * col_margin

            if d is None:
                yield from [FormattedCell(None)] * w
            else:
                if len(d) != w:
                    raise ValueError(f"Table row has {len(d)} columns. Expected {w}.")
                for value in d:
                    yield value

    for _ in range(row_margin):
        yield []

    if write_captions:
        yield list(table_name_row())
        yield list(table_description_row())

    if write_pre_rows:
        for row_data in zip_longest(*(t.pre_rows for t in tables), fillvalue=None):
            yield list(row(row_data))

    yield list(header_row())

    n_data_rows = 0
    for row_data in zip_longest(*(t.rows for t in tables), fillvalue=None):
        yield list(row(row_data))
        n_data_rows += 1

    if n_data_rows < 1:
        # Tables are not allowed to have zero rows. Add an empty row.
        yield []


def define_list_object(
    sheet: WriteOnlyWorksheet,
    first_column: int,
    first_row: int,
    name: str,
    column_names: Sequence[str],
    n_data_rows: int,
    style: Optional[TableStyleInfo],
) -> Table:
    last_column = first_column - 1 + len(column_names)
    last_row = first_row + max(n_data_rows, 1)

    table = Table(
        displayName=name,
        ref=f"{get_column_letter(first_column)}{first_row}:{get_column_letter(last_column)}{last_row}",
    )
    # noinspection PyProtectedMember
    table._initialise_columns()
    for column, value in zip(table.tableColumns, column_names):
        column.name = value

    if style:
        table.tableStyleInfo = style

    with warnings.catch_warnings():
        # See https://foss.heptapod.net/openpyxl/openpyxl/-/issues/1760
        warnings.simplefilter(action="ignore", category=UserWarning)
        sheet.add_table(table)

    return table


def define_named_ranges_for_dict_table(
    *,
    book: Workbook,
    sheet_name: str,
    first_table_row: int,
    first_table_col: int,
    keys: Sequence[str],
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
        keys: The dictionary keys, in the same order as in the first table column.
        workbook_scope: Whether to make a workbook-scoped named range (True) or a sheet-scoped named range (False).
    """
    # Values are in the second table column
    col = first_table_col + 1

    for i, key in enumerate(keys):
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
