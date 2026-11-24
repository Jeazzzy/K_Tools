import os
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest
from unittest.mock import patch

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtCore import Qt
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

    def test_folder_selection_enables_batch_mode_and_ignores_archive_documents(self):
        page = LoDescriptionPage()
        try:
            with TemporaryDirectory() as temporary:
                root = Path(temporary)
                locality = root / "деревня Выбье"
                locality.mkdir()
                (locality / "описание.docx").write_bytes(b"example")
                archive = root / "Архив"
                archive.mkdir()
                (archive / "старое.docx").write_bytes(b"old")

                page.source_edit.setText(str(root))
                page._source_selected(str(root))

                self.assertTrue(page._is_batch_mode())
                self.assertEqual(len(page._batch_documents), 1)
                self.assertEqual(page._batch_documents[0].locality, "деревня Выбье")
                self.assertTrue(page.batch_preview_card.isVisibleTo(page))
                self.assertTrue(page.log_text.isVisibleTo(page))
                self.assertEqual(page.batch_table.rowCount(), 1)
                self.assertEqual(page.batch_table.item(0, 2).text(), "НП")
                self.assertIn("деревня Выбье", page.batch_table.item(0, 3).text())
                self.assertFalse(page.target_edit.isEnabled())
                self.assertFalse(page.locality_edit.isEnabled())
                self.assertIn("Выбрано для обработки: 1", page.status_label.text())

                page.batch_table.item(0, 0).setCheckState(Qt.CheckState.Unchecked)
                self.assertEqual(page._selected_batch_documents(), ())
                self.assertFalse(page.run_button.isEnabled())
                page._set_all_batch_checked(True)
                self.assertEqual(len(page._selected_batch_documents()), 1)
                page._batch_message("[1/1] готово. Примерно осталось: 0 с.")
                self.assertIn("Примерно осталось", page.status_label.text())
        finally:
            page.close()

    def test_batch_preview_shows_zone_and_outside_addresses(self):
        page = LoDescriptionPage()
        try:
            with TemporaryDirectory() as temporary:
                root = Path(temporary)
                inside = root / "_д. Выбье" / "Ж1"
                outside = root / "_Вне НП" / "И1"
                inside.mkdir(parents=True)
                outside.mkdir(parents=True)
                (inside / "внутри.docx").write_bytes(b"inside")
                (outside / "снаружи.docx").write_bytes(b"outside")

                page.region_edit.setText("Ленинградская область")
                page.district_combo.setEditText("Кингисеппский м.р-н")
                page.settlement_combo.setEditText("Усть-Лужское с.п.")
                page.source_edit.setText(str(root))
                page._source_selected(str(root))

                rows = {
                    page.batch_table.item(row, 2).text(): page.batch_table.item(row, 3).text()
                    for row in range(page.batch_table.rowCount())
                }
                self.assertEqual(
                    rows["ТЗ Ж1"],
                    "Ленинградская область, Кингисеппский м.р-н, "
                    "Усть-Лужское с.п., деревня Выбье",
                )
                self.assertEqual(
                    rows["Вне НП · ТЗ И1"],
                    "Ленинградская область, Кингисеппский м.р-н, Усть-Лужское с.п.",
                )
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
