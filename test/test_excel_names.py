import unittest

from aa_py_openpyxl_util import (
    is_a1_reference_like,
    is_r1c1_reference_like,
    is_valid_excel_defined_name,
    is_valid_excel_sheet_title,
    is_valid_excel_table_name,
    make_safe_excel_defined_name,
    make_safe_excel_sheet_title,
    make_safe_excel_table_name,
    validate_excel_defined_name,
    validate_excel_sheet_title,
    validate_excel_table_name,
    validate_unique_excel_names,
)
from aa_py_openpyxl_util._excel_names import (
    DuplicateExcelNameError,
    InvalidExcelNameError,
    InvalidExcelSheetTitleError,
)


class TestExcelNameValidation(unittest.TestCase):
    def test_valid_excel_names(self) -> None:
        names = [
            "Table1",
            "Foo",
            "_Foo",
            "\\Foo",
            "Foo_Bar",
            "Foo.Bar",
            "XFE1",
            "A1048577",
            "Sales2024",
            "\u00e9Foo",
            "\u540d\u5b57",
            "A" * 255,
        ]

        for name in names:
            with self.subTest(name=name):
                self.assertEqual(name, validate_excel_table_name(name))
                self.assertEqual(name, validate_excel_defined_name(name))
                self.assertTrue(is_valid_excel_table_name(name))
                self.assertTrue(is_valid_excel_defined_name(name))

    def test_invalid_excel_names(self) -> None:
        names = [
            "",
            None,
            "A1",
            "A01",
            "XFD1048576",
            "key1",
            "May01",
            "ROI2016",
            "R1C1",
            "R1C",
            "RC1",
            "RC",
            "R",
            "C",
            "1Foo",
            ".Foo",
            "Foo Bar",
            "Foo-Bar",
            "Foo#Bar",
            "_xl",
            "_xlcn.LinkedTable_Table11",
            "A" * 256,
        ]

        for name in names:
            with self.subTest(name=name):
                with self.assertRaises(InvalidExcelNameError):
                    validate_excel_table_name(name)
                with self.assertRaises(InvalidExcelNameError):
                    validate_excel_defined_name(name)
                self.assertFalse(is_valid_excel_table_name(name))
                self.assertFalse(is_valid_excel_defined_name(name))

    def test_invalid_name_error_message_is_specific(self) -> None:
        with self.assertRaises(InvalidExcelNameError) as cm:
            validate_excel_table_name("A1")

        self.assertIn("A1-style cell reference", str(cm.exception))
        self.assertIn("'A1'", str(cm.exception))

    def test_a1_reference_like_edges(self) -> None:
        self.assertTrue(is_a1_reference_like("A1"))
        self.assertTrue(is_a1_reference_like("A01"))
        self.assertTrue(is_a1_reference_like("May01"))
        self.assertTrue(is_a1_reference_like("ROI2016"))
        self.assertTrue(is_a1_reference_like("XFD1048576"))
        self.assertFalse(is_a1_reference_like("A0"))
        self.assertFalse(is_a1_reference_like("XFE1"))
        self.assertFalse(is_a1_reference_like("A1048577"))
        self.assertFalse(is_a1_reference_like("Sales2024"))

    def test_r1c1_reference_like_edges(self) -> None:
        self.assertTrue(is_r1c1_reference_like("R1C1"))
        self.assertTrue(is_r1c1_reference_like("R1C"))
        self.assertTrue(is_r1c1_reference_like("RC1"))
        self.assertTrue(is_r1c1_reference_like("RC"))
        self.assertTrue(is_r1c1_reference_like("r1c1"))
        self.assertFalse(is_r1c1_reference_like("R0C1"))
        self.assertFalse(is_r1c1_reference_like("RC0"))
        self.assertFalse(is_r1c1_reference_like("R1048577C1"))

    def test_make_safe_excel_table_name(self) -> None:
        self.assertEqual("Table_A1", make_safe_excel_table_name("A1"))
        self.assertEqual("Foo_Bar", make_safe_excel_table_name("Foo Bar"))
        self.assertEqual("Foo_Bar", make_safe_excel_table_name("Foo-Bar"))
        self.assertEqual("Table_1Foo", make_safe_excel_table_name("1Foo"))
        self.assertEqual("Table__xlFoo", make_safe_excel_table_name("_xlFoo"))
        self.assertEqual(
            "Table_2",
            make_safe_excel_table_name("Table", existing_names=["table"]),
        )
        self.assertEqual(255, len(make_safe_excel_table_name("A" * 256)))
        self.assertTrue(is_valid_excel_table_name(make_safe_excel_table_name("A1")))

    def test_make_safe_excel_defined_name(self) -> None:
        self.assertEqual("Name_R1C1", make_safe_excel_defined_name("R1C1"))
        self.assertEqual(
            "Name_2", make_safe_excel_defined_name("Name", existing_names=["name"])
        )
        self.assertTrue(is_valid_excel_defined_name(make_safe_excel_defined_name("A1")))

    def test_validate_unique_excel_names(self) -> None:
        with self.assertRaises(DuplicateExcelNameError) as cm:
            validate_unique_excel_names(["Foo", "foo"], scope_label="workbook")

        self.assertIn("case-insensitive", str(cm.exception))


class TestExcelSheetTitleValidation(unittest.TestCase):
    def test_valid_sheet_title(self) -> None:
        self.assertEqual("Sheet1", validate_excel_sheet_title("Sheet1"))
        self.assertEqual("A" * 31, validate_excel_sheet_title("A" * 31))
        self.assertTrue(is_valid_excel_sheet_title("Sheet1"))

    def test_invalid_sheet_title(self) -> None:
        for title in ["", None, "A" * 32, "'Foo", "Foo'", "Foo:Bar"]:
            with self.subTest(title=title):
                with self.assertRaises(InvalidExcelSheetTitleError):
                    validate_excel_sheet_title(title)
                self.assertFalse(is_valid_excel_sheet_title(title))

    def test_make_safe_sheet_title(self) -> None:
        self.assertEqual("Foo_Bar", make_safe_excel_sheet_title("Foo:Bar"))
        self.assertEqual("Foo", make_safe_excel_sheet_title("'Foo'"))
        self.assertEqual("Sheet", make_safe_excel_sheet_title(""))
        self.assertEqual(31, len(make_safe_excel_sheet_title("A" * 32)))
        self.assertEqual(
            "Sheet_2",
            make_safe_excel_sheet_title("Sheet", existing_titles=["sheet"]),
        )


if __name__ == "__main__":
    unittest.main(failfast=True)
