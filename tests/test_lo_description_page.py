import os
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest
from unittest.mock import patch

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtWidgets import QApplication

from pages.lo_description import LoDescriptionPage
from core.word_description_formatter import DocumentInfo


class LoDescriptionPageTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QApplication.instance() or QApplication([])

    def test_page_has_editable_address_hints_and_limited_popups(self):
        page = LoDescriptionPage()
        try:
            self.assertTrue(page.district_combo.isEditable())
            self.assertTrue(page.settlement_combo.isEditable())
            self.assertEqual(page.district_combo.maxVisibleItems(), 12)
            self.assertEqual(page.district_combo.currentText(), "")
            self.assertIn("Гатчинский муниципальный округ", page.addresses)

            page.district_combo.setEditText("Кингисеппский муниципальный район")
            page._load_settlements(preserve=False)
            options = [
                page.settlement_combo.itemText(index)
                for index in range(page.settlement_combo.count())
            ]
            self.assertIn("Большелуцкое сельское поселение", options)
        finally:
            page.close()

    def test_preview_uses_all_non_empty_address_parts(self):
        page = LoDescriptionPage()
        try:
            page.region_edit.setText("Ленинградская область")
            page.district_combo.setEditText("Гатчинский муниципальный округ")
            page.settlement_combo.setEditText("")
            page.locality_edit.setText("деревня Пудость")
            self.app.processEvents()
            self.assertEqual(
                page.preview_label.text(),
                "Ленинградская область, Гатчинский муниципальный округ, деревня Пудость",
            )
        finally:
            page.close()

    def test_source_selection_suggests_separate_output_file(self):
        page = LoDescriptionPage()
        try:
            with TemporaryDirectory() as temporary:
                source = Path(temporary) / "Описание.docx"
                source.write_bytes(b"example")
                with patch.object(page, "_start_inspection") as inspect:
                    page._source_selected(str(source))

                self.assertEqual(
                    page.target_edit.text(),
                    str(source.with_name("Описание_оформлено.docx")),
                )
                inspect.assert_called_once_with(source)
        finally:
            page.close()

    def test_inspection_does_not_treat_locality_as_settlement_in_urban_district(self):
        page = LoDescriptionPage()
        try:
            source = Path("C:/documents/example.docx")
            page.source_edit.setText(str(source))
            page._inspection_generation = 1
            page._inspection_done(
                1,
                source,
                DocumentInfo(
                    "Ленинградская область, Гатчинский муниципальный округ, деревня Пудость",
                    "деревня Пудость",
                    2,
                ),
            )

            self.assertEqual(page.settlement_combo.currentText(), "")
            self.assertEqual(page.locality_edit.text(), "деревня Пудость")
        finally:
            page.close()


if __name__ == "__main__":
    unittest.main()
