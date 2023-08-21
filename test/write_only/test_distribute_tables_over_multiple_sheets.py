import unittest
from typing import List

from aa_py_openpyxl_util import TableInfo

# noinspection PyProtectedMember
from aa_py_openpyxl_util._write_only import distribute_tables_over_multiple_sheets


def get_names(tables: List[List[TableInfo]]) -> List[List[str]]:
    # Just to make comparing test results easier.
    return [[table.name for table in sheet] for sheet in tables]


class TestDistributeTablesOverMultipleSheets(unittest.TestCase):
    def test_empty(self) -> None:
        self.assertEqual(
            [],
            distribute_tables_over_multiple_sheets(
                tables=[], max_sheet_width=10, left_margin=1, gutter=2
            ),
        )

    def test_one_table(self) -> None:
        table1 = TableInfo(name="Table1", column_names=["a", "b"], rows=[])

        self.assertEqual(
            [[table1]],
            distribute_tables_over_multiple_sheets(
                tables=[table1], max_sheet_width=10, left_margin=1, gutter=2
            ),
        )

    def test_table_too_wide(self) -> None:
        table1 = TableInfo(
            name="Table1",
            column_names=["a", "b", "c", "d", "e", "f", "g", "h", "j", "k"],
            rows=[],
        )

        with self.assertRaises(ValueError):
            distribute_tables_over_multiple_sheets(
                tables=[table1], max_sheet_width=10, left_margin=3, gutter=2
            )

    def test_three_tables_same_length(self) -> None:
        table_a = TableInfo(name="Table A", column_names=["a", "b", "c", "d"], rows=[])
        table_b = TableInfo(name="Table B", column_names=["a", "b", "c", "d"], rows=[])
        table_c = TableInfo(name="Table C", column_names=["a", "b", "c", "d"], rows=[])

        cases = {
            # Everything fits on one sheet:
            16: [[table_a, table_b, table_c]],  # One column open on the right.
            15: [[table_a, table_b, table_c]],  # No columns open on the right.
            # Two tables per sheet:
            14: [[table_a, table_b], [table_c]],
            13: [[table_a, table_b], [table_c]],
            12: [[table_a, table_b], [table_c]],
            11: [[table_a, table_b], [table_c]],  # One column open on the right.
            10: [[table_a, table_b], [table_c]],  # No columns open on the right.
            # One table per sheet:
            9: [[table_a], [table_b], [table_c]],
            8: [[table_a], [table_b], [table_c]],
            7: [[table_a], [table_b], [table_c]],
            6: [[table_a], [table_b], [table_c]],  # One column open on the right.
            5: [[table_a], [table_b], [table_c]],  # No columns open on the right.
        }

        def get_names(tables: List[List[TableInfo]]) -> List[List[str]]:
            # Just to make comparing test results easier.
            return [[table.name for table in sheet] for sheet in tables]

        for max_sheet_width, expected_result in cases.items():
            with self.subTest(max_sheet_width=max_sheet_width):
                self.assertEqual(
                    get_names(expected_result),
                    get_names(
                        distribute_tables_over_multiple_sheets(
                            tables=[table_a, table_b, table_c],
                            max_sheet_width=max_sheet_width,
                            left_margin=1,
                            gutter=1,
                        )
                    ),
                )

    def test_single_column_no_margins(self) -> None:
        table_a = TableInfo(name="Table A", column_names=["a"], rows=[])
        table_b = TableInfo(name="Table B", column_names=["a"], rows=[])
        table_c = TableInfo(name="Table C", column_names=["a"], rows=[])

        cases = {
            4: [[table_a, table_b, table_c]],
            3: [[table_a, table_b, table_c]],
            2: [[table_a, table_b], [table_c]],
            1: [[table_a], [table_b], [table_c]],
        }

        for max_sheet_width, expected_result in cases.items():
            with self.subTest(max_sheet_width=max_sheet_width):
                self.assertEqual(
                    get_names(expected_result),
                    get_names(
                        distribute_tables_over_multiple_sheets(
                            tables=[table_a, table_b, table_c],
                            max_sheet_width=max_sheet_width,
                            left_margin=0,
                            gutter=0,
                        )
                    ),
                )

    def test_wide_then_narrow(self) -> None:
        table_wide = TableInfo(
            name="Wide", column_names=["a", "b", "c", "d", "e", "f"], rows=[]
        )
        table_narrow = TableInfo(name="Narrow", column_names=["a", "b"], rows=[])

        cases = {
            12: [[table_wide, table_narrow]],  # One column open on the right.
            11: [[table_wide, table_narrow]],  # No columns open on the right.
            10: [[table_wide], [table_narrow]],
            9: [[table_wide], [table_narrow]],
            8: [[table_wide], [table_narrow]],  # One column open on the right.
            7: [[table_wide], [table_narrow]],  # No columns open on the right.
        }

        for max_sheet_width, expected_result in cases.items():
            with self.subTest(max_sheet_width=max_sheet_width):
                self.assertEqual(
                    get_names(expected_result),
                    get_names(
                        distribute_tables_over_multiple_sheets(
                            tables=[table_wide, table_narrow],
                            max_sheet_width=max_sheet_width,
                            left_margin=1,
                            gutter=2,
                        )
                    ),
                )

    def test_narrow_then_wide(self) -> None:
        table_wide = TableInfo(
            name="Wide", column_names=["a", "b", "c", "d", "e", "f"], rows=[]
        )
        table_narrow = TableInfo(name="Narrow", column_names=["a", "b"], rows=[])

        cases = {
            12: [[table_narrow, table_wide]],  # One column open on the right.
            11: [[table_narrow, table_wide]],  # No columns open on the right.
            10: [[table_narrow], [table_wide]],
            9: [[table_narrow], [table_wide]],
            8: [[table_narrow], [table_wide]],
            7: [[table_narrow], [table_wide]],
        }

        for max_sheet_width, expected_result in cases.items():
            with self.subTest(max_sheet_width=max_sheet_width):
                self.assertEqual(
                    get_names(expected_result),
                    get_names(
                        distribute_tables_over_multiple_sheets(
                            tables=[table_narrow, table_wide],
                            max_sheet_width=max_sheet_width,
                            left_margin=1,
                            gutter=2,
                        )
                    ),
                )
