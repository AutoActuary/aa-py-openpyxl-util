import unittest

from aa_py_openpyxl_util import FormattedCell, TableInfo

# noinspection PyProtectedMember
from aa_py_openpyxl_util._write_only import stack_table_rows_side_by_side


class TestStackTableRowsSideBySide(unittest.TestCase):
    def test_empty(self) -> None:
        self.assertEqual(
            [
                [],
                [
                    FormattedCell(None),
                    FormattedCell("a"),
                    FormattedCell("b"),
                    FormattedCell("c"),
                    FormattedCell(None),
                    FormattedCell("d"),
                    FormattedCell("e"),
                ],
                [],
            ],
            list(
                stack_table_rows_side_by_side(
                    tables=[
                        TableInfo(
                            name="Table1",
                            column_names=["a", "b", "c"],
                            rows=[],
                        ),
                        TableInfo(
                            name="Table2",
                            column_names=["d", "e"],
                            rows=[],
                        ),
                    ],
                    row_margin=1,
                    col_margin=1,
                    write_captions=False,
                    write_pre_rows=False,
                )
            ),
        )

    def test_same_length(self) -> None:
        self.assertEqual(
            [
                [],
                [
                    FormattedCell(None),
                    FormattedCell("a"),
                    FormattedCell("b"),
                    FormattedCell("c"),
                    FormattedCell(None),
                    FormattedCell("d"),
                    FormattedCell("e"),
                ],
                [
                    FormattedCell(None),
                    FormattedCell(1),
                    FormattedCell(2),
                    FormattedCell(3),
                    FormattedCell(None),
                    FormattedCell(7),
                    FormattedCell(8),
                ],
                [
                    FormattedCell(None),
                    FormattedCell(4),
                    FormattedCell(5),
                    FormattedCell(6),
                    FormattedCell(None),
                    FormattedCell(9),
                    FormattedCell(0),
                ],
            ],
            list(
                stack_table_rows_side_by_side(
                    tables=[
                        TableInfo(
                            name="Table1",
                            column_names=["a", "b", "c"],
                            rows=[
                                [FormattedCell(1), FormattedCell(2), FormattedCell(3)],
                                [FormattedCell(4), FormattedCell(5), FormattedCell(6)],
                            ],
                        ),
                        TableInfo(
                            name="Table2",
                            column_names=["d", "e"],
                            rows=[
                                [FormattedCell(7), FormattedCell(8)],
                                [FormattedCell(9), FormattedCell(0)],
                            ],
                        ),
                    ],
                    row_margin=1,
                    col_margin=1,
                    write_captions=False,
                    write_pre_rows=False,
                )
            ),
        )

    def test_different_lengths(self) -> None:
        self.assertEqual(
            [
                [],
                [
                    FormattedCell(None),
                    FormattedCell("a"),
                    FormattedCell("b"),
                    FormattedCell("c"),
                    FormattedCell(None),
                    FormattedCell("d"),
                    FormattedCell("e"),
                    FormattedCell(None),
                    FormattedCell("f"),
                    FormattedCell("g"),
                ],
                [
                    FormattedCell(None),
                    FormattedCell(1),
                    FormattedCell(2),
                    FormattedCell(3),
                    FormattedCell(None),
                    FormattedCell(None),
                    FormattedCell(None),
                    FormattedCell(None),
                    FormattedCell(7),
                    FormattedCell(8),
                ],
                [
                    FormattedCell(None),
                    FormattedCell(2),
                    FormattedCell(3),
                    FormattedCell(4),
                    FormattedCell(None),
                    FormattedCell(None),
                    FormattedCell(None),
                    FormattedCell(None),
                    FormattedCell(9),
                    FormattedCell(0),
                ],
                [
                    FormattedCell(None),
                    FormattedCell(3),
                    FormattedCell(4),
                    FormattedCell(5),
                    FormattedCell(None),
                    FormattedCell(None),
                    FormattedCell(None),
                    FormattedCell(None),
                    FormattedCell(None),
                    FormattedCell(None),
                ],
            ],
            list(
                stack_table_rows_side_by_side(
                    tables=[
                        TableInfo(
                            name="Table1",
                            column_names=["a", "b", "c"],
                            rows=[
                                [FormattedCell(1), FormattedCell(2), FormattedCell(3)],
                                [FormattedCell(2), FormattedCell(3), FormattedCell(4)],
                                [FormattedCell(3), FormattedCell(4), FormattedCell(5)],
                            ],
                        ),
                        TableInfo(
                            name="Table2",
                            column_names=["d", "e"],
                            rows=[],
                        ),
                        TableInfo(
                            name="Table3",
                            column_names=["f", "g"],
                            rows=[
                                [FormattedCell(7), FormattedCell(8)],
                                [FormattedCell(9), FormattedCell(0)],
                            ],
                        ),
                    ],
                    row_margin=1,
                    col_margin=1,
                    write_captions=False,
                    write_pre_rows=False,
                )
            ),
        )

    def test_caption(self) -> None:
        self.assertEqual(
            [
                [],
                [
                    FormattedCell(None),
                    FormattedCell("Table1"),
                    FormattedCell(None),
                    FormattedCell(None),
                    FormattedCell(None),
                    FormattedCell("Table2"),
                    FormattedCell(None),
                    FormattedCell(None),
                    FormattedCell("Table3"),
                    FormattedCell(None),
                ],
                [
                    FormattedCell(None),
                    FormattedCell("This is the first table"),
                    FormattedCell(None),
                    FormattedCell(None),
                    FormattedCell(None),
                    FormattedCell(""),
                    FormattedCell(None),
                    FormattedCell(None),
                    FormattedCell("This is the last table"),
                    FormattedCell(None),
                ],
                [
                    FormattedCell(None),
                    FormattedCell("a"),
                    FormattedCell("b"),
                    FormattedCell("c"),
                    FormattedCell(None),
                    FormattedCell("d"),
                    FormattedCell("e"),
                    FormattedCell(None),
                    FormattedCell("f"),
                    FormattedCell("g"),
                ],
                [
                    FormattedCell(None),
                    FormattedCell(1),
                    FormattedCell(2),
                    FormattedCell(3),
                    FormattedCell(None),
                    FormattedCell(None),
                    FormattedCell(None),
                    FormattedCell(None),
                    FormattedCell(7),
                    FormattedCell(8),
                ],
                [
                    FormattedCell(None),
                    FormattedCell(2),
                    FormattedCell(3),
                    FormattedCell(4),
                    FormattedCell(None),
                    FormattedCell(None),
                    FormattedCell(None),
                    FormattedCell(None),
                    FormattedCell(9),
                    FormattedCell(0),
                ],
                [
                    FormattedCell(None),
                    FormattedCell(3),
                    FormattedCell(4),
                    FormattedCell(5),
                    FormattedCell(None),
                    FormattedCell(None),
                    FormattedCell(None),
                    FormattedCell(None),
                    FormattedCell(None),
                    FormattedCell(None),
                ],
            ],
            list(
                stack_table_rows_side_by_side(
                    tables=[
                        TableInfo(
                            name="Table1",
                            column_names=["a", "b", "c"],
                            rows=[
                                [FormattedCell(1), FormattedCell(2), FormattedCell(3)],
                                [FormattedCell(2), FormattedCell(3), FormattedCell(4)],
                                [FormattedCell(3), FormattedCell(4), FormattedCell(5)],
                            ],
                            description="This is the first table",
                        ),
                        TableInfo(
                            name="Table2",
                            column_names=["d", "e"],
                            rows=[],
                        ),
                        TableInfo(
                            name="Table3",
                            column_names=["f", "g"],
                            rows=[
                                [FormattedCell(7), FormattedCell(8)],
                                [FormattedCell(9), FormattedCell(0)],
                            ],
                            description="This is the last table",
                        ),
                    ],
                    row_margin=1,
                    col_margin=1,
                    write_captions=True,
                    write_pre_rows=False,
                )
            ),
        )

    def test_pre_rows(self) -> None:
        self.assertEqual(
            [
                [],
                [
                    FormattedCell(None),
                    FormattedCell("AA"),
                    FormattedCell("BB"),
                    FormattedCell("CC"),
                    FormattedCell(None),
                    FormattedCell(None),
                    FormattedCell(None),
                    FormattedCell(None),
                    FormattedCell("FF"),
                    FormattedCell("GG"),
                ],
                [
                    FormattedCell(None),
                    FormattedCell("aa"),
                    FormattedCell("bb"),
                    FormattedCell("cc"),
                    FormattedCell(None),
                    FormattedCell(None),
                    FormattedCell(None),
                    FormattedCell(None),
                    FormattedCell(None),
                    FormattedCell(None),
                ],
                [
                    FormattedCell(None),
                    FormattedCell("a"),
                    FormattedCell("b"),
                    FormattedCell("c"),
                    FormattedCell(None),
                    FormattedCell("d"),
                    FormattedCell("e"),
                    FormattedCell(None),
                    FormattedCell("f"),
                    FormattedCell("g"),
                ],
                [
                    FormattedCell(None),
                    FormattedCell(1),
                    FormattedCell(2),
                    FormattedCell(3),
                    FormattedCell(None),
                    FormattedCell(None),
                    FormattedCell(None),
                    FormattedCell(None),
                    FormattedCell(7),
                    FormattedCell(8),
                ],
                [
                    FormattedCell(None),
                    FormattedCell(2),
                    FormattedCell(3),
                    FormattedCell(4),
                    FormattedCell(None),
                    FormattedCell(None),
                    FormattedCell(None),
                    FormattedCell(None),
                    FormattedCell(9),
                    FormattedCell(0),
                ],
                [
                    FormattedCell(None),
                    FormattedCell(3),
                    FormattedCell(4),
                    FormattedCell(5),
                    FormattedCell(None),
                    FormattedCell(None),
                    FormattedCell(None),
                    FormattedCell(None),
                    FormattedCell(None),
                    FormattedCell(None),
                ],
            ],
            list(
                stack_table_rows_side_by_side(
                    tables=[
                        TableInfo(
                            name="Table1",
                            column_names=["a", "b", "c"],
                            rows=[
                                [FormattedCell(1), FormattedCell(2), FormattedCell(3)],
                                [FormattedCell(2), FormattedCell(3), FormattedCell(4)],
                                [FormattedCell(3), FormattedCell(4), FormattedCell(5)],
                            ],
                            pre_rows=[
                                [
                                    FormattedCell("AA"),
                                    FormattedCell("BB"),
                                    FormattedCell("CC"),
                                ],
                                [
                                    FormattedCell("aa"),
                                    FormattedCell("bb"),
                                    FormattedCell("cc"),
                                ],
                            ],
                        ),
                        TableInfo(
                            name="Table2",
                            column_names=["d", "e"],
                            rows=[],
                        ),
                        TableInfo(
                            name="Table3",
                            column_names=["f", "g"],
                            rows=[
                                [FormattedCell(7), FormattedCell(8)],
                                [FormattedCell(9), FormattedCell(0)],
                            ],
                            pre_rows=[
                                [FormattedCell("FF"), FormattedCell("GG")],
                            ],
                            description="Description",
                        ),
                    ],
                    row_margin=1,
                    col_margin=1,
                    write_captions=False,
                    write_pre_rows=True,
                )
            ),
        )


if __name__ == "__main__":
    unittest.main(
        failfast=True,
    )
