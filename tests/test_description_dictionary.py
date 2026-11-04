import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import MagicMock, patch

import xlrd
import xlwt

from cvetopt.invoice.description_dictionary import (
    _append_rows_xlwings,
    _dictionary_next_row,
    append_missing_descriptions,
    is_cargo_summary_row,
    row_values_in_order,
)


class DictionaryAppendTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.path = Path(self.temp.name) / "dictionary.xls"

    def make_dictionary(self, *, later_translation=False):
        book = xlwt.Workbook()
        sheet = book.add_sheet("Список")
        sheet.write(0, 1, "Description")
        sheet.write(0, 2, "Перевод")
        sheet.write(1, 1, "Known Flower")
        sheet.write(1, 2, "Цветок")
        # Actual empty text records, as in the user's workbook, not BLANK cells.
        for row in range(2, 30):
            sheet.row(row).set_cell_text(2, "" if row % 2 else " \t ")
        sheet.write(35, 0, "Other column data")
        if later_translation:
            sheet.write(31, 2, "Сохранить перевод")
        book.save(str(self.path))

    def test_ignores_empty_text_tail_and_other_columns(self):
        self.make_dictionary()
        self.assertEqual(_dictionary_next_row(self.path), 3)

    def test_keeps_internal_gaps_and_translation_only_rows(self):
        self.make_dictionary(later_translation=True)
        self.assertEqual(_dictionary_next_row(self.path), 33)

    def test_append_fallback_reuses_empty_tail_and_deduplicates(self):
        self.make_dictionary()
        with patch("cvetopt.invoice.description_dictionary.sys.platform", "darwin"):
            self.assertEqual(append_missing_descriptions(self.path, ["New Flower"]), 1)
            self.assertEqual(append_missing_descriptions(self.path, ["New Flower"]), 0)
        sheet = xlrd.open_workbook(str(self.path)).sheet_by_index(0)
        self.assertEqual(sheet.cell_value(2, 1), "New Flower")
        self.assertEqual(sheet.cell_value(2, 2), "")
        self.assertEqual(sheet.cell_value(1, 2), "Цветок")
        self.assertEqual(sheet.cell_value(35, 0), "Other column data")

    def test_excel_writer_starts_after_actual_content(self):
        self.make_dictionary()
        app = MagicMock()
        excel = SimpleNamespace(App=MagicMock(return_value=app))
        with patch.dict("sys.modules", {"xlwings": excel}):
            _append_rows_xlwings(self.path, [("New Flower", "")])
        sheet = app.books.open.return_value.sheets[0]
        sheet.range.assert_called_once_with("B3")
        self.assertEqual(sheet.range.return_value.resize.return_value.value, [["New Flower", ""]])
        app.books.open.return_value.save.assert_called_once()


class CargoSummaryRowTests(unittest.TestCase):
    def test_customs_code_first_is_cargo(self):
        # Как в инвойсе 27.10.26: строка сводки Cargo внизу.
        self.assertTrue(is_cargo_summary_row(["'06031970'", "", "", "", 80.0, "Eustoma"]))
        self.assertTrue(is_cargo_summary_row(["", 6031300.0, "Orchids"]))  # Excel съел ноль

    def test_product_row_with_code_further_right_is_not_cargo(self):
        row = ["117213", "", 1.0, 60.0, 60.0, "Chr G Magnum White", "", 1.3] + [""] * 9 + ["'06031400'"]
        self.assertFalse(is_cargo_summary_row(row))
        self.assertFalse(is_cargo_summary_row([110380.0, "Chr T Babe"]))

    def test_xlsx_row_order(self):
        row = {"AB": "x", "B": "'06031970'", "C": "Hydrangea"}
        self.assertEqual(row_values_in_order(row), ["'06031970'", "Hydrangea", "x"])


if __name__ == "__main__":
    unittest.main()
