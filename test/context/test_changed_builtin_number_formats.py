import unittest

from locate import this_dir
from openpyxl.cell import Cell

from aa_py_openpyxl_util import changed_builtin_number_formats, safe_load_workbook

data_dir = this_dir().parent.parent.joinpath("test_data")


class TestChangedBuiltinNumberFormats(unittest.TestCase):
    def test_1(self) -> None:
        self._test("yyyy/mm/dd")
        self._test(r"mm\-dd\-yy")

    def _test(self, regional_date_format: str) -> None:
        with changed_builtin_number_formats({14: regional_date_format}):
            with safe_load_workbook(
                path=data_dir.joinpath("number_formats.xlsx"),
                read_only=True,
                data_only=True,
            ) as book:
                sheet = book["Sheet1"]
                regional_date: Cell = sheet["B1"]
                explicit_date: Cell = sheet["B2"]

                # This is affected by `changed_builtin_number_formats`:
                self.assertEqual(regional_date_format, regional_date.number_format)

                # This is not affected by `changed_builtin_number_formats`:
                self.assertEqual(r"yyyy\-mm\-dd", explicit_date.number_format)


if __name__ == "__main__":
    unittest.main(
        failfast=True,
    )
