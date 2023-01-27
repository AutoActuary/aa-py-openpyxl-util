from typing import Tuple, List, Any, Callable

from openpyxl.cell import Cell


def get_cell_values(cells: Tuple[Tuple[Cell, ...], ...]) -> List[List[Any]]:
    """
    Get the values of the cells in the given table.

    Args:
        cells: 2D Tuples of Cell objects.

    Returns:
        2D list of cell values
    """
    return process_cells(
        cells=cells,
        callback=lambda cell: cell.value,
    )


def process_cells(
    *,
    cells: Tuple[Tuple[Cell, ...], ...],
    callback: Callable[[Cell], Any],
) -> List[List[Any]]:
    """
    Process the cells in the given table.

    Args:
        cells: 2D Tuples of Cell objects.
        callback: The callback to process each cell.

    Returns:
        2D list of processed cell values
    """
    return [[callback(cell) for cell in row] for row in cells]
