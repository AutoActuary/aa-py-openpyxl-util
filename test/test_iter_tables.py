import re
import unittest

from locate import this_dir

from aa_py_openpyxl_util import (
    safe_load_workbook,
    iter_named_range_tables,
    iter_list_object_tables,
)

data_dir = this_dir().parent.joinpath("test_data")


class TestIterNamedRangeTables(unittest.TestCase):
    def test_1(self) -> None:
        with safe_load_workbook(
            path=data_dir.joinpath("tables.xlsx"),
            read_only=False,
            data_only=False,
        ) as book:
            self.assertEqual(
                [
                    ("Sheet2", "FooBar2", "$F$2:$H$5"),
                    ("Sheet1", "Table1", "$B$2:$C$4"),
                ],
                list(
                    (
                        (s.title, t, r)
                        for s, t, r in iter_named_range_tables(
                            book=book,
                            exclude_names=[],
                            exclude_sheets=[],
                        )
                    )
                ),
            )

    def test_exclude_sheets(self) -> None:
        with safe_load_workbook(
            path=data_dir.joinpath("tables.xlsx"),
            read_only=False,
            data_only=False,
        ) as book:
            self.assertEqual(
                [
                    ("Sheet2", "FooBar2", "$F$2:$H$5"),
                ],
                list(
                    (
                        (s.title, t, r)
                        for s, t, r in iter_named_range_tables(
                            book=book,
                            exclude_names=[],
                            exclude_sheets=[re.compile(r"sheet1", re.IGNORECASE)],
                        )
                    )
                ),
            )

    def test_exclude_tables(self) -> None:
        with safe_load_workbook(
            path=data_dir.joinpath("tables.xlsx"),
            read_only=False,
            data_only=False,
        ) as book:
            self.assertEqual(
                [
                    ("Sheet2", "FooBar2", "$F$2:$H$5"),
                ],
                list(
                    (
                        (s.title, t, r)
                        for s, t, r in iter_named_range_tables(
                            book=book,
                            exclude_names=[re.compile(r"table\d*", re.IGNORECASE)],
                            exclude_sheets=[],
                        )
                    )
                ),
            )


class TestIterListObjectTables(unittest.TestCase):
    def test_1(self) -> None:
        with safe_load_workbook(
            path=data_dir.joinpath("tables.xlsx"),
            read_only=False,
            data_only=False,
        ) as book:
            self.assertEqual(
                [
                    ("Sheet1", "Table2", "E2:F3"),
                    ("Sheet2", "FooBar1", "B2:D5"),
                ],
                list(
                    (
                        (s.title, t.name, t.ref)
                        for s, t in iter_list_object_tables(
                            book=book,
                            exclude_list_objects=[],
                            exclude_sheets=[],
                        )
                    )
                ),
            )

    def test_exclude_sheets(self) -> None:
        with safe_load_workbook(
            path=data_dir.joinpath("tables.xlsx"),
            read_only=False,
            data_only=False,
        ) as book:
            self.assertEqual(
                [
                    ("Sheet1", "Table2", "E2:F3"),
                ],
                list(
                    (
                        (s.title, t.name, t.ref)
                        for s, t in iter_list_object_tables(
                            book=book,
                            exclude_list_objects=[],
                            exclude_sheets=[re.compile(r"\w*2", re.IGNORECASE)],
                        )
                    )
                ),
            )

    def test_exclude_tables(self) -> None:
        with safe_load_workbook(
            path=data_dir.joinpath("tables.xlsx"),
            read_only=False,
            data_only=False,
        ) as book:
            self.assertEqual(
                [
                    ("Sheet1", "Table2", "E2:F3"),
                ],
                list(
                    (
                        (s.title, t.name, t.ref)
                        for s, t in iter_list_object_tables(
                            book=book,
                            exclude_list_objects=[
                                re.compile(r"FooBar\d*", re.IGNORECASE)
                            ],
                            exclude_sheets=[],
                        )
                    )
                ),
            )


if __name__ == "__main__":
    unittest.main(
        failfast=True,
    )
