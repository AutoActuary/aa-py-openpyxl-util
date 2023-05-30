from datetime import datetime
from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED

from openpyxl.workbook import Workbook
from openpyxl.writer.excel import ExcelWriter


def save_workbook_workaround(*, book: Workbook, p: Path) -> None:
    """
    Workaround for https://foss.heptapod.net/openpyxl/openpyxl/-/issues/2042 .
    Use this instead of the `Workbook.save` method.

    Args:
        book: The openpyxl book to save.
        p: The path to which to write the book.
    """
    if book.read_only:
        raise TypeError("""Workbook is read-only""")
    if book.write_only and not book.worksheets:
        book.create_sheet()

    book.properties.modified = datetime.utcnow()
    with ZipFile(
        file=p,
        mode="w",
        compression=ZIP_DEFLATED,
        allowZip64=True,
    ) as archive:
        ExcelWriter(book, archive).write_data()
