import unittest

from locate import this_dir

from aa_py_openpyxl_util import (
    safe_load_workbook,
    find_table,
)

# noinspection PyProtectedMember
from aa_py_openpyxl_util._find_table import (
    find_named_range_by_name,
    find_list_object_by_name,
)

data_dir = this_dir().parent.joinpath("test_data")


class TestFindTable(unittest.TestCase):
    def test_tables(self) -> None:
        with safe_load_workbook(
            path=data_dir.joinpath("tables.xlsx"),
            read_only=False,
            data_only=False,
        ) as book:
            sheet, table_range = find_table(book=book, name="Table1")
            self.assertEqual("Sheet1", sheet.title)
            self.assertEqual("$B$2:$C$4", table_range)

            sheet, table_range = find_table(book=book, name="Table2")
            self.assertEqual("Sheet1", sheet.title)
            self.assertEqual("E2:F3", table_range)

    def test_not_tables(self) -> None:
        with safe_load_workbook(
            path=data_dir.joinpath("tables.xlsx"),
            read_only=False,
            data_only=False,
        ) as book:
            # Does not exist
            with self.assertRaises(KeyError):
                find_table(book=book, name="Table999")

            # Single cell is not a table
            with self.assertRaises(KeyError):
                find_table(book=book, name="SingleCell1")

            # Single row is not a table
            with self.assertRaises(KeyError):
                find_table(book=book, name="SingleRow1")


class TestFindNamedRangeTable(unittest.TestCase):
    def test_table1(self) -> None:
        with safe_load_workbook(
            path=data_dir.joinpath("tables.xlsx"),
            read_only=False,
            data_only=False,
        ) as book:
            sheet, table_range = find_named_range_by_name(book=book, name="Table1")
            self.assertEqual("Sheet1", sheet.title)
            self.assertEqual("$B$2:$C$4", table_range)

    def test_single_cell(self) -> None:
        with safe_load_workbook(
            path=data_dir.joinpath("tables.xlsx"),
            read_only=False,
            data_only=False,
        ) as book:
            sheet, table_range = find_named_range_by_name(book=book, name="SingleCell1")
            self.assertEqual("Sheet1", sheet.title)
            self.assertEqual("$H$2", table_range)

    def test_single_row(self) -> None:
        with safe_load_workbook(
            path=data_dir.joinpath("tables.xlsx"),
            read_only=False,
            data_only=False,
        ) as book:
            sheet, table_range = find_named_range_by_name(book=book, name="SingleRow1")
            self.assertEqual("Sheet1", sheet.title)
            self.assertEqual("$J$2:$L$2", table_range)

    def test_not_found(self) -> None:
        with safe_load_workbook(
            path=data_dir.joinpath("tables.xlsx"),
            read_only=False,
            data_only=False,
        ) as book:
            with self.assertRaises(KeyError):
                find_named_range_by_name(book=book, name="Table999")


class TestFindListObjectTable(unittest.TestCase):
    def test_1(self) -> None:
        with safe_load_workbook(
            path=data_dir.joinpath("tables.xlsx"),
            read_only=False,
            data_only=False,
        ) as book:
            sheet, table = find_list_object_by_name(book=book, name="Table2")
            self.assertEqual("Sheet1", sheet.title)
            self.assertEqual("E2:F3", table.ref)

    def test_not_found(self) -> None:
        with safe_load_workbook(
            path=data_dir.joinpath("tables.xlsx"),
            read_only=False,
            data_only=False,
        ) as book:
            with self.assertRaises(KeyError):
                find_list_object_by_name(book=book, name="Table999")


if __name__ == "__main__":
    unittest.main(
        failfast=True,
    )
