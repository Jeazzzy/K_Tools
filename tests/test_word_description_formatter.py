from pathlib import Path
from tempfile import TemporaryDirectory
import unittest
from unittest.mock import patch

from core import word_description_formatter as formatter


class WordDescriptionFormatterTests(unittest.TestCase):
    def test_normalizes_short_locality_types(self):
        self.assertEqual(formatter.normalize_locality("д. Буркин Буерак"), "деревня Буркин Буерак")
        self.assertEqual(formatter.normalize_locality("гп Кузнечное"), "городской посёлок Кузнечное")

    def test_extracts_locality_from_address_or_title(self):
        self.assertEqual(
            formatter.extract_locality(
                "Ленинградская область, район, поселение, д. Большой Луцк",
                "",
            ),
            "деревня Большой Луцк",
        )
        self.assertEqual(
            formatter.extract_locality("", "село Путилово Ленинградской области"),
            "село Путилово",
        )

    def test_builds_address_without_empty_settlement_for_district(self):
        self.assertEqual(
            formatter.build_full_address(
                "Ленинградская область",
                "Гатчинский муниципальный округ",
                "",
                "деревня Пудость",
            ),
            "Ленинградская область, Гатчинский муниципальный округ, деревня Пудость",
        )

    def test_loads_packaged_lenoblast_address_hints(self):
        addresses = formatter.load_lo_addresses()
        self.assertIn("Кингисеппский муниципальный район", addresses)
        self.assertIn(
            "Большелуцкое сельское поселение",
            addresses["Кингисеппский муниципальный район"],
        )
        self.assertEqual(addresses["Гатчинский муниципальный округ"], [])

    def test_rejects_source_as_target_before_starting_word(self):
        with TemporaryDirectory() as temporary:
            source = Path(temporary) / "описание.docx"
            source.write_bytes(b"not-a-real-docx")
            with self.assertRaisesRegex(ValueError, "должны отличаться"):
                formatter.format_document(
                    source,
                    source,
                    "Ленинградская область",
                    "Гатчинский муниципальный округ",
                    "",
                    "деревня Пудость",
                )

    def test_failure_preserves_source_and_existing_target(self):
        with TemporaryDirectory() as temporary:
            folder = Path(temporary)
            source = folder / "исходник.docx"
            target = folder / "результат.docx"
            source.write_bytes(b"source")
            target.write_bytes(b"old-target")
            fake_pythoncom = type(
                "FakePythonCom",
                (),
                {"CoInitialize": staticmethod(lambda: None), "CoUninitialize": staticmethod(lambda: None)},
            )()

            with (
                patch.object(formatter, "pythoncom", fake_pythoncom),
                patch.object(formatter, "win32com", object()),
                patch.object(formatter, "_open_word", side_effect=RuntimeError("Word недоступен")),
            ):
                with self.assertRaisesRegex(RuntimeError, "Word недоступен"):
                    formatter.format_document(
                        source,
                        target,
                        "Ленинградская область",
                        "Гатчинский муниципальный округ",
                        "",
                        "деревня Пудость",
                    )

            self.assertEqual(source.read_bytes(), b"source")
            self.assertEqual(target.read_bytes(), b"old-target")
            self.assertFalse(list(folder.glob(".ktools-description-*")))


if __name__ == "__main__":
    unittest.main()
