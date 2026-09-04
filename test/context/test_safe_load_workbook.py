import unittest
from io import BytesIO

from locate import this_dir

from aa_py_openpyxl_util import safe_load_workbook

workbook_path = this_dir().parent.parent.joinpath("test_data", "number_formats.xlsx")


class TestSafeLoadWorkbook(unittest.TestCase):
    def test_string_path(self) -> None:
        with safe_load_workbook(
            path=str(workbook_path),
            read_only=True,
            data_only=True,
        ) as book:
            self.assertEqual(["Sheet1"], book.sheetnames)

    def test_binary_file_like_object(self) -> None:
        workbook_file = BytesIO(workbook_path.read_bytes())

        with safe_load_workbook(
            path=workbook_file,
            read_only=True,
            data_only=True,
        ) as book:
            self.assertEqual(["Sheet1"], book.sheetnames)

        self.assertFalse(workbook_file.closed)
