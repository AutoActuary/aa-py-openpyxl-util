import unittest

from aa_py_openpyxl_util import FormattedCell


class TestFormattedCell(unittest.TestCase):
    def test_short_formula(self) -> None:
        fc = FormattedCell(value="=SUM(A1:A10)", number_format="0.00")
        checked = fc.check()
        self.assertIs(fc, checked)

    def test_maximum_length_formula(self) -> None:
        # Create a formula of the maximum length that Excel can support.
        s = '"abcdefghijklmnopqrstuvwxyz"'
        concat_args = ", ".join([s] * 136)
        concat = f"_xlfn.CONCAT({concat_args})"
        f = f'{concat}&{concat}&"abcd"'
        self.assertEqual(8192, len(f))

        fc = FormattedCell(value=f"={f}", number_format="0.00")
        checked = fc.check()
        self.assertIs(fc, checked)

        # Check manually with Excel to verify that this is indeed a valid workbook:
        # from openpyxl.workbook import Workbook
        # book = Workbook(write_only=True)
        # ws = book.create_sheet()
        # ws.append([fc.create_openpyxl_cell(ws, "A1")])
        # book.save("max.xlsx")

    def test_too_long_formula(self) -> None:
        # Create a formula that is too long for Excel.
        s = '"abcdefghijklmnopqrstuvwxyz"'
        concat_args = ", ".join([s] * 136)
        concat = f"_xlfn.CONCAT({concat_args})"
        f = f'{concat}&{concat}&"abcde"'
        self.assertEqual(8193, len(f))

        fc = FormattedCell(value=f"={f}", number_format="0.00")
        with self.assertRaises(ValueError):
            fc.check()

        # Check manually with Excel to verify that this is indeed a corrupted workbook:
        # from openpyxl.workbook import Workbook
        # book = Workbook(write_only=True)
        # ws = book.create_sheet()
        # ws.append([fc.create_openpyxl_cell(ws, "A1")])
        # book.save("too_long.xlsx")
