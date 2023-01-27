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
                [("Sheet1", "Table1", "$B$2:$C$4")],
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


class TestIterListObjectTables(unittest.TestCase):
    def test_1(self) -> None:
        with safe_load_workbook(
            path=data_dir.joinpath("tables.xlsx"),
            read_only=False,
            data_only=False,
        ) as book:
            self.assertEqual(
                [("Sheet1", "Table2", "E2:F3")],
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


if __name__ == "__main__":
    unittest.main(
        failfast=True,
    )
