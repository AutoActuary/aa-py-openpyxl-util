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
        # Keep all these variables names the same length for readability.
        c___ = FormattedCell(None)

        cel1 = FormattedCell(1)
        cel2 = FormattedCell(2)
        cel3 = FormattedCell(3)
        cel4 = FormattedCell(4)
        cel5 = FormattedCell(5)
        cel7 = FormattedCell(7)
        cel8 = FormattedCell(8)
        cel9 = FormattedCell(9)
        cel0 = FormattedCell(0)

        pre1 = FormattedCell("AA")
        pre2 = FormattedCell("BB")
        pre3 = FormattedCell("CC")
        pre4 = FormattedCell("aa")
        pre5 = FormattedCell("bb")
        pre6 = FormattedCell("cc")
        pre7 = FormattedCell("FF")
        pre8 = FormattedCell("GG")

        col1 = FormattedCell("a")
        col2 = FormattedCell("b")
        col3 = FormattedCell("c")
        col4 = FormattedCell("d")
        col5 = FormattedCell("e")
        col6 = FormattedCell("f")
        col7 = FormattedCell("g")

        self.assertEqual(
            [
                [],
                [c___, pre1, pre2, pre3, c___, c___, c___, c___, pre7, pre8],
                [c___, pre4, pre5, pre6, c___, c___, c___, c___, c___, c___],
                [c___, col1, col2, col3, c___, col4, col5, c___, col6, col7],
                [c___, cel1, cel2, cel3, c___, c___, c___, c___, cel7, cel8],
                [c___, cel2, cel3, cel4, c___, c___, c___, c___, cel9, cel0],
                [c___, cel3, cel4, cel5, c___, c___, c___, c___, c___, c___],
            ],
            list(
                stack_table_rows_side_by_side(
                    tables=[
                        TableInfo(
                            name="Table1",
                            column_names=["a", "b", "c"],
                            rows=[
                                [cel1, cel2, cel3],
                                [cel2, cel3, cel4],
                                [cel3, cel4, cel5],
                            ],
                            pre_rows=[
                                [pre1, pre2, pre3],
                                [pre4, pre5, pre6],
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
                                [cel7, cel8],
                                [cel9, cel0],
                            ],
                            pre_rows=[
                                [pre7, pre8],
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

    def test_pre_rows_varying_widths(self) -> None:
        """
        Test stacking tables where the `pre_rows` are not the same width as the table.
        """
        # Keep all these variables names the same length for readability.
        c___ = FormattedCell(None)

        cel1 = FormattedCell(1)
        cel2 = FormattedCell(2)
        cel3 = FormattedCell(3)
        cel4 = FormattedCell(4)

        pre1 = FormattedCell("AA")
        pre2 = FormattedCell("BB")
        pre3 = FormattedCell("CC")
        pre4 = FormattedCell("aa")

        col1 = FormattedCell("a")
        col2 = FormattedCell("b")

        nam1 = FormattedCell("Table1")
        nam2 = FormattedCell("Table2")
        des1 = FormattedCell("Description1")
        des2 = FormattedCell("Description2")

        self.assertEqual(
            [
                [],
                [c___, nam1, c___, c___, nam2, c___],
                [c___, des1, c___, c___, des2, c___],
                [c___, pre1, pre2, c___, pre1, c___],
                [c___, pre3, pre4, c___, pre2, c___],
                [c___, col1, c___, c___, col1, col2],
                [c___, cel1, c___, c___, cel1, cel2],
                [c___, cel2, c___, c___, cel2, cel3],
                [c___, cel3, c___, c___, cel3, cel4],
            ],
            list(
                stack_table_rows_side_by_side(
                    tables=[
                        # Pre-rows wider than table.
                        TableInfo(
                            name="Table1",
                            description="Description1",
                            column_names=["a"],
                            rows=[
                                [cel1],
                                [cel2],
                                [cel3],
                            ],
                            pre_rows=[
                                [pre1, pre2],
                                [pre3, pre4],
                            ],
                        ),
                        # Pre-rows narrower than table.
                        TableInfo(
                            name="Table2",
                            description="Description2",
                            column_names=["a", "b"],
                            rows=[
                                [cel1, cel2],
                                [cel2, cel3],
                                [cel3, cel4],
                            ],
                            pre_rows=[
                                [pre1],
                                [pre2],
                            ],
                        ),
                    ],
                    row_margin=1,
                    col_margin=1,
                    write_captions=True,
                    write_pre_rows=True,
                )
            ),
        )


if __name__ == "__main__":
    unittest.main(
        failfast=True,
    )
