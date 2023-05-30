import unittest
from contextlib import contextmanager
from pathlib import Path
from tempfile import TemporaryDirectory
from typing import Generator

from openpyxl.workbook import Workbook
from portalocker import Lock, LockFlags

# noinspection PyProtectedMember
from aa_py_openpyxl_util._workaround import save_workbook_workaround


@contextmanager
def create_locked_file(
    file_path: Path,
) -> Generator[Path, None, None]:
    """
    This function is used only for testing on a locked file.
    """
    if file_path.exists():
        raise FileExistsError("Cannot create locked file, because it already exists.")
    file_path = file_path.resolve()

    try:
        with Lock(
            filename=file_path,
            mode="w",
            flags=LockFlags.EXCLUSIVE,
        ):
            yield file_path
    finally:
        # This should NOT fail. It will fail if there is a leaked file handle.
        file_path.unlink(missing_ok=True)


class TestSaveWorkbookWorkaround(unittest.TestCase):
    def test_save_over_locked_file(self) -> None:
        book = Workbook(write_only=True)
        with TemporaryDirectory() as tmp_dir_str:
            tmp_dir = Path(tmp_dir_str)
            file_path = tmp_dir / "test.xlsx"
            self.assertFalse(file_path.exists())
            with create_locked_file(file_path):
                self.assertTrue(file_path.exists())
                with self.assertRaises(PermissionError):
                    save_workbook_workaround(book=book, p=file_path)
                self.assertTrue(file_path.exists())
            self.assertFalse(file_path.exists())

    def test_save_new_file(self) -> None:
        book = Workbook(write_only=True)
        with TemporaryDirectory() as tmp_dir_str:
            tmp_dir = Path(tmp_dir_str)
            file_path = tmp_dir / "test.xlsx"
            self.assertFalse(file_path.exists())
            save_workbook_workaround(book=book, p=file_path)
            self.assertTrue(file_path.exists())
