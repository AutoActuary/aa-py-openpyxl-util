import os
from contextlib import suppress
from datetime import datetime
from functools import cache
from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED
from openpyxl.workbook import Workbook
from openpyxl.writer.excel import ExcelWriter
import tempfile
import atexit
import time
import warnings


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


@cache
def remove_atexit_permission_error() -> None:
    """
    Registers a new atexit function to preemptively handle PermissionError in openpyxl.

    This function registers a new atexit function to preemptively handle the PermissionError in openpyxl.
    It checks if the necessary module and attributes are available and proceeds to register the new function,
    which aims to handle the PermissionError scenario. Additionally, it removes stale temporary files
    associated with openpyxl, which might have been left over by previous PermissionError scenarios.

    Warnings:
        RuntimeWarning: If the required module or attribute is not found, a warning is raised.

    Returns:
        None
    """
    try:
        import openpyxl.worksheet._writer
    except ModuleNotFoundError:
        warnings.warn(
            "ModuleNotFoundError in aa_py_openpyxl_util.remove_atexit_permission_error: cannot import openpyxl.worksheet._writer",
            RuntimeWarning,
        )
        return

    if not hasattr(openpyxl.worksheet._writer, "ALL_TEMP_FILES"):
        warnings.warn(
            "AttributeError in aa_py_openpyxl_util.remove_atexit_permission_error: openpyxl.worksheet._writer does not have attribute 'ALL_TEMP_FILES'",
            RuntimeWarning,
        )
        return

    # Remove old openpyxl temp files (likely due to PermissionError)
    now = time.time()
    for file_path in Path(tempfile.gettempdir()).glob("openpyxl.*"):
        with suppress(PermissionError, FileNotFoundError):
            if now - os.stat(file_path).st_mtime > (60 * 60 * 24 * 30):
                os.remove(file_path)

    @atexit.register
    def _openpyxl_shutdown_fix() -> None:
        temp_files_copy = openpyxl.worksheet._writer.ALL_TEMP_FILES.copy()
        openpyxl.worksheet._writer.ALL_TEMP_FILES.clear()

        for path in temp_files_copy:
            with suppress(PermissionError, FileNotFoundError):
                os.remove(path)
