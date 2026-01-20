import unittest
from logging import WARNING

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
    def test_tables__case_sensitive(self) -> None:
        with safe_load_workbook(
            path=data_dir.joinpath("tables.xlsx"),
            read_only=False,
            data_only=False,
        ) as book:
            sheet, table_range = find_table(book=book, name="Table1", ci=False)
            self.assertEqual("Sheet1", sheet.title)
            self.assertEqual("$B$2:$C$4", table_range)

            for name in ["table1", "TABLE1"]:
                with self.assertRaises(KeyError):
                    find_table(book=book, name=name, ci=False)

            sheet, table_range = find_table(book=book, name="Table2", ci=False)
            self.assertEqual("Sheet1", sheet.title)
            self.assertEqual("E2:F3", table_range)

            for name in ["table2", "TABLE2"]:
                with self.assertRaises(KeyError):
                    find_table(book=book, name=name, ci=False)

    def test_tables__case_insensitive(self) -> None:
        with self.assertNoLogs(level=WARNING):
            with safe_load_workbook(
                path=data_dir.joinpath("tables.xlsx"),
                read_only=False,
                data_only=False,
            ) as book:
                for name in ["Table1", "table1", "TABLE1"]:
                    sheet, table_range = find_table(book=book, name=name, ci=True)
                    self.assertEqual("Sheet1", sheet.title)
                    self.assertEqual("$B$2:$C$4", table_range)

                for name in ["Table2", "table2", "TABLE2"]:
                    sheet, table_range = find_table(book=book, name=name, ci=True)
                    self.assertEqual("Sheet1", sheet.title)
                    self.assertEqual("E2:F3", table_range)

    def test_tables__case_insensitive_warn(self) -> None:
        with self.assertLogs(level=WARNING) as logs:
            with safe_load_workbook(
                path=data_dir.joinpath("tables.xlsx"),
                read_only=False,
                data_only=False,
            ) as book:
                for name in ["Table1", "table1", "TABLE1"]:
                    sheet, table_range = find_table(book=book, name=name, ci="warn")
                    self.assertEqual("Sheet1", sheet.title)
                    self.assertEqual("$B$2:$C$4", table_range)

                for name in ["Table2", "table2", "TABLE2"]:
                    sheet, table_range = find_table(book=book, name=name, ci="warn")
                    self.assertEqual("Sheet1", sheet.title)
                    self.assertEqual("E2:F3", table_range)

        self.assertEqual(
            [
                "Table with exact name `table1` not found. Using case-insensitive match `Table1` instead.",
                "Table with exact name `TABLE1` not found. Using case-insensitive match `Table1` instead.",
                "Table with exact name `table2` not found. Using case-insensitive match `Table2` instead.",
                "Table with exact name `TABLE2` not found. Using case-insensitive match `Table2` instead.",
            ],
            [r.message for r in logs.records],
        )

    def test_not_tables(self) -> None:
        with safe_load_workbook(
            path=data_dir.joinpath("tables.xlsx"),
            read_only=False,
            data_only=False,
        ) as book:
            # Does not exist
            with self.assertRaises(KeyError):
                find_table(book=book, name="Table999", ci=False)

            # Single cell is not a table
            with self.assertRaises(KeyError):
                find_table(book=book, name="SingleCell1", ci=False)

            # Single row is not a table
            with self.assertRaises(KeyError):
                find_table(book=book, name="SingleRow1", ci=False)


class TestFindNamedRangeTable(unittest.TestCase):

    def test_case_sensitive(self) -> None:
        with safe_load_workbook(
            path=data_dir.joinpath("tables.xlsx"),
            read_only=False,
            data_only=False,
        ) as book:
            sheet, table_range = find_named_range_by_name(
                book=book, name="Table1", ci=False
            )
            self.assertEqual("Sheet1", sheet.title)
            self.assertEqual("$B$2:$C$4", table_range)

            for name in ["table1", "TABLE1"]:
                with self.assertRaises(KeyError):
                    find_named_range_by_name(book=book, name=name, ci=False)

    def test_case_insensitive(self) -> None:
        with self.assertNoLogs(level=WARNING):
            with safe_load_workbook(
                path=data_dir.joinpath("tables.xlsx"),
                read_only=False,
                data_only=False,
            ) as book:
                for name in ["Table1", "table1", "TABLE1"]:
                    sheet, table_range = find_named_range_by_name(
                        book=book, name=name, ci=True
                    )
                    self.assertEqual("Sheet1", sheet.title)
                    self.assertEqual("$B$2:$C$4", table_range)

    def test_case_insensitive_warn(self) -> None:
        with self.assertLogs(level=WARNING) as logs:
            with safe_load_workbook(
                path=data_dir.joinpath("tables.xlsx"),
                read_only=False,
                data_only=False,
            ) as book:
                for name in ["Table1", "table1", "TABLE1"]:
                    sheet, table_range = find_named_range_by_name(
                        book=book, name=name, ci="warn"
                    )
                    self.assertEqual("Sheet1", sheet.title)
                    self.assertEqual("$B$2:$C$4", table_range)

        self.assertEqual(
            [
                "Table with exact name `table1` not found. Using case-insensitive match `Table1` instead.",
                "Table with exact name `TABLE1` not found. Using case-insensitive match `Table1` instead.",
            ],
            [r.message for r in logs.records],
        )

    def test_single_cell(self) -> None:
        with safe_load_workbook(
            path=data_dir.joinpath("tables.xlsx"),
            read_only=False,
            data_only=False,
        ) as book:
            sheet, table_range = find_named_range_by_name(
                book=book, name="SingleCell1", ci=False
            )
            self.assertEqual("Sheet1", sheet.title)
            self.assertEqual("$H$2", table_range)

    def test_single_row(self) -> None:
        with safe_load_workbook(
            path=data_dir.joinpath("tables.xlsx"),
            read_only=False,
            data_only=False,
        ) as book:
            sheet, table_range = find_named_range_by_name(
                book=book, name="SingleRow1", ci=False
            )
            self.assertEqual("Sheet1", sheet.title)
            self.assertEqual("$J$2:$L$2", table_range)

    def test_not_found(self) -> None:
        with safe_load_workbook(
            path=data_dir.joinpath("tables.xlsx"),
            read_only=False,
            data_only=False,
        ) as book:
            with self.assertRaises(KeyError):
                find_named_range_by_name(book=book, name="Table999", ci=False)


class TestFindListObjectTable(unittest.TestCase):
    def test_case_sensitive(self) -> None:
        with safe_load_workbook(
            path=data_dir.joinpath("tables.xlsx"),
            read_only=False,
            data_only=False,
        ) as book:
            sheet, table = find_list_object_by_name(book=book, name="Table2", ci=False)
            self.assertEqual("Sheet1", sheet.title)
            self.assertEqual("E2:F3", table.ref)

            for name in ["table2", "TABLE2"]:
                with self.assertRaises(KeyError):
                    find_list_object_by_name(book=book, name=name, ci=False)

    def test_case_insensitive(self) -> None:
        with self.assertNoLogs(level=WARNING):
            with safe_load_workbook(
                path=data_dir.joinpath("tables.xlsx"),
                read_only=False,
                data_only=False,
            ) as book:
                for name in ["Table2", "table2", "TABLE2"]:
                    sheet, table = find_list_object_by_name(
                        book=book, name=name, ci=True
                    )
                    self.assertEqual("Sheet1", sheet.title)
                    self.assertEqual("E2:F3", table.ref)

    def test_case_insensitive_warn(self) -> None:
        with self.assertLogs(level=WARNING) as logs:
            with safe_load_workbook(
                path=data_dir.joinpath("tables.xlsx"),
                read_only=False,
                data_only=False,
            ) as book:
                for name in ["Table2", "table2", "TABLE2"]:
                    sheet, table = find_list_object_by_name(
                        book=book, name=name, ci="warn"
                    )
                    self.assertEqual("Sheet1", sheet.title)
                    self.assertEqual("E2:F3", table.ref)

        self.assertEqual(
            [
                "Table with exact name `table2` not found. Using case-insensitive match `Table2` instead.",
                "Table with exact name `TABLE2` not found. Using case-insensitive match `Table2` instead.",
            ],
            [r.message for r in logs.records],
        )

    def test_not_found(self) -> None:
        with safe_load_workbook(
            path=data_dir.joinpath("tables.xlsx"),
            read_only=False,
            data_only=False,
        ) as book:
            with self.assertRaises(KeyError):
                find_list_object_by_name(book=book, name="Table999", ci=False)


if __name__ == "__main__":
    unittest.main(
        failfast=True,
    )
