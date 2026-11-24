from pathlib import Path
from tempfile import TemporaryDirectory
import unittest
from unittest.mock import Mock, patch

from core import word_description_formatter as formatter


class WordDescriptionFormatterTests(unittest.TestCase):
    def test_reads_locality_only_from_explicit_folder_name(self):
        self.assertEqual(
            formatter.locality_from_folder_name("деревня Выбье"),
            "деревня Выбье",
        )
        self.assertEqual(
            formatter.locality_from_folder_name("п. Усть-Луга"),
            "посёлок Усть-Луга",
        )
        self.assertEqual(
            formatter.locality_from_folder_name("_д. Выбье"),
            "деревня Выбье",
        )
        self.assertEqual(formatter.locality_from_folder_name("Архив"), "")
        self.assertTrue(formatter.is_zone_folder_name("Ж1"))
        self.assertTrue(formatter.is_zone_folder_name("П-2.1"))
        self.assertFalse(formatter.is_zone_folder_name("РН"))
        self.assertFalse(formatter.is_zone_folder_name("!архив"))
        self.assertTrue(formatter.is_outside_locality_folder("_Вне НП"))
        self.assertFalse(formatter.is_outside_locality_folder("деревня Выбье"))

    def test_batch_discovery_skips_non_locality_folders_but_scans_their_children(self):
        with TemporaryDirectory() as temporary:
            root = Path(temporary)
            direct = root / "_д. Выбье"
            direct.mkdir()
            (direct / "описание.docx").write_bytes(b"word")
            (direct / "описание_оформлено.docx").write_bytes(b"result")
            (direct / "~$описание.docx").write_bytes(b"lock")
            nested_archive = direct / "Архив"
            nested_archive.mkdir()
            (nested_archive / "старое.docx").write_bytes(b"old")
            zone = direct / "Ж1"
            zone.mkdir()
            (zone / "графическое описание.docx").write_bytes(b"zone")
            service = direct / "РН"
            service.mkdir()
            (service / "служебное.docx").write_bytes(b"ignore")

            archive = root / "Архив"
            nested_locality = archive / "поселок Усть-Луга"
            nested_locality.mkdir(parents=True)
            (archive / "чужое.docx").write_bytes(b"ignore")
            (nested_locality / "границы.doc").write_bytes(b"word")
            outside_zone = root / "_Вне НП" / "И1"
            outside_zone.mkdir(parents=True)
            (outside_zone / "вне населённого пункта.docx").write_bytes(b"word")

            documents = formatter.find_batch_documents(root)

            self.assertEqual(
                {document.source.relative_to(root).as_posix() for document in documents},
                {
                    "_д. Выбье/описание.docx",
                    "_д. Выбье/Ж1/графическое описание.docx",
                    "Архив/поселок Усть-Луга/границы.doc",
                    "_Вне НП/И1/вне населённого пункта.docx",
                },
            )
            localities = {document.source.name: document.locality for document in documents}
            self.assertEqual(localities["описание.docx"], "деревня Выбье")
            self.assertEqual(localities["графическое описание.docx"], "деревня Выбье")
            self.assertEqual(localities["границы.doc"], "посёлок Усть-Луга")
            self.assertEqual(localities["вне населённого пункта.docx"], "")

    def test_explicit_empty_locality_means_outside_settlement(self):
        address = (
            "Ленинградская область, Кингисеппский м.р-н, "
            "Усть-Лужское с.п., д. Выбье"
        )
        self.assertEqual(
            formatter.resolve_locality(None, address, ""),
            "деревня Выбье",
        )
        self.assertEqual(formatter.resolve_locality("", address, ""), "")
        self.assertEqual(
            formatter.build_full_address(
                "Ленинградская область",
                "Кингисеппский м.р-н",
                "Усть-Лужское с.п.",
                "",
            ),
            "Ленинградская область, Кингисеппский м.р-н, Усть-Лужское с.п.",
        )

    def test_batch_formatting_continues_after_one_document_fails(self):
        with TemporaryDirectory() as temporary:
            root = Path(temporary)
            locality = root / "деревня Выбье"
            locality.mkdir()
            (locality / "a.docx").write_bytes(b"a")
            (locality / "b.docx").write_bytes(b"b")
            progress = []
            messages = []
            fake_pythoncom = Mock()
            fake_word = object()

            def format_one(_word, source, *_args):
                if Path(source).name == "b.docx":
                    raise RuntimeError("повреждённый документ")
                return formatter.DocumentInfo("адрес", "деревня Выбье", 2)

            with (
                patch.object(formatter, "pythoncom", fake_pythoncom),
                patch.object(formatter, "win32com", object()),
                patch.object(formatter, "_open_word", return_value=fake_word) as open_word,
                patch.object(formatter, "_finish_word") as finish_word,
                patch.object(
                    formatter,
                    "_format_document_in_word",
                    side_effect=format_one,
                ),
            ):
                result = formatter.format_description_folder(
                    root,
                    "Ленинградская область",
                    "Кингисеппский муниципальный район",
                    "Большелуцкое сельское поселение",
                    lambda current, total: progress.append((current, total)),
                    messages.append,
                )

            open_word.assert_called_once_with()
            finish_word.assert_called_once_with(fake_word)
            fake_pythoncom.CoInitialize.assert_called_once_with()
            fake_pythoncom.CoUninitialize.assert_called_once_with()
            self.assertEqual(len(result.documents), 2)
            self.assertEqual(len(result.completed), 1)
            self.assertEqual(len(result.failures), 1)
            self.assertEqual(result.failures[0].source.name, "b.docx")
            self.assertEqual(progress, [(0, 2), (1, 2), (2, 2)])
            self.assertTrue(any("[1/2] Начало" in message for message in messages))
            self.assertTrue(any("Примерно осталось" in message for message in messages))

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
