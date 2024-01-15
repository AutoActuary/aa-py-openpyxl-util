import unittest
from datetime import datetime
from pathlib import Path

from aa_py_openpyxl_util import safe_load_workbook, extract_data_from_numbered_tables

repo_dir = Path(__file__).parent.parent
data_dir = repo_dir / "test_data/extract"


class TestExtractDataFromNumberedTables(unittest.TestCase):
    def test_missing_table(self) -> None:
        with safe_load_workbook(
            path=data_dir / "empty.xlsx",
            read_only=False,  # Because openpyxl can't find ListObjects in read-only mode!
            data_only=True,
        ) as book:
            results = list(
                extract_data_from_numbered_tables(book=book, base_name="Table")
            )
            self.assertEqual([], results)

    def test_dates(self) -> None:
        with safe_load_workbook(
            path=data_dir / "dates.xlsx",
            read_only=False,  # Because openpyxl can't find ListObjects in read-only mode!
            data_only=True,
        ) as book:
            results = list(
                extract_data_from_numbered_tables(book=book, base_name="Table")
            )
            self.assertEqual(
                [
                    {
                        "DateValues": datetime(2024, 1, 15, 0, 0, tzinfo=None),
                        "DateFormulas": datetime(2024, 1, 15, 0, 0, tzinfo=None),
                    },
                    {
                        "DateValues": datetime(2024, 1, 15, 10, 30, tzinfo=None),
                        "DateFormulas": datetime(2024, 1, 15, 10, 30, tzinfo=None),
                    },
                ],
                results,
            )
