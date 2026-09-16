import os
import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import MagicMock

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from pages.mdb_operations import (
    MdbCopyPage,
    _location_components,
    _locality_phrase,
    _replace_locality_text,
)


class MdbAddressTextTests(unittest.TestCase):
    SAMPLE = (
        "ВИ1. Зона виноградников в границах Вышестеблиевского "
        "сельского поселения Темрюкского района Краснодарского края"
    )

    def test_formats_fias_administrative_type(self):
        self.assertEqual(
            _locality_phrase("Ахтанизовское", "с.п."),
            "Ахтанизовского сельского поселения",
        )

    def test_replaces_real_release_wording_and_keeps_higher_levels(self):
        value, matched = _replace_locality_text(
            self.SAMPLE, "Ахтанизовское", "с.п."
        )
        self.assertTrue(matched)
        self.assertEqual(
            value,
            "ВИ1. Зона виноградников в границах Ахтанизовского "
            "сельского поселения Темрюкского района Краснодарского края",
        )

    def test_outside_mode_removes_settlement_but_keeps_district(self):
        value, matched = _replace_locality_text(self.SAMPLE, outside=True)
        self.assertTrue(matched)
        self.assertEqual(
            value,
            "ВИ1. Зона виноградников в границах муниципального образования "
            "Темрюкского района Краснодарского края",
        )

    def test_repeated_replacement_is_idempotent(self):
        expected, _ = _replace_locality_text(
            self.SAMPLE, "Ахтанизовское", "с.п."
        )
        repeated, matched = _replace_locality_text(
            expected, "Ахтанизовское", "с.п."
        )
        self.assertTrue(matched)
        self.assertEqual(repeated, expected)

    def test_prefers_locality_and_keeps_parent_settlement(self):
        components = _location_components(
            {
                "city_name": "Курчанское",
                "city_type": "с.п.",
                "locality_name": "Красный Октябрь",
                "locality_type": "п.",
            }
        )

        self.assertEqual(
            components,
            ("Красный Октябрь", "п.", "Курчанское", "с.п."),
        )

    def test_adds_locality_before_parent_settlement_and_is_idempotent(self):
        expected = (
            "Ж1. Зона застройки индивидуальными жилыми домами в границах "
            "поселка Красный Октябрь Курчанского сельского поселения "
            "Темрюкского района Краснодарского края"
        )
        original = (
            "Ж1. Зона застройки индивидуальными жилыми домами в границах "
            "Курчанского сельского поселения Темрюкского района "
            "Краснодарского края"
        )

        updated, matched = _replace_locality_text(
            original,
            "Красный Октябрь",
            "п.",
            parent_name="Курчанское",
            parent_kind="с.п.",
        )
        repeated, repeated_matched = _replace_locality_text(
            updated,
            "Красный Октябрь",
            "п.",
            parent_name="Курчанское",
            parent_kind="с.п.",
        )

        self.assertTrue(matched)
        self.assertTrue(repeated_matched)
        self.assertEqual(updated, expected)
        self.assertEqual(repeated, expected)


class MdbPathTests(unittest.TestCase):
    def test_source_mdb_can_be_inside_nested_rn_folder(self):
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            source = root / "ВИ1" / "РН" / "Проект.mdb"
            source.parent.mkdir(parents=True)
            source.touch()
            result = MdbCopyPage._collect_source_by_index(root, {"ВИ1"})
            self.assertEqual(Path(result["ВИ1"]), source)

    def test_ambiguous_source_is_reported(self):
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            folder = root / "ВИ1" / "РН"
            folder.mkdir(parents=True)
            (folder / "Первый.mdb").touch()
            (folder / "Второй.mdb").touch()
            with self.assertRaisesRegex(ValueError, "несколько source MDB"):
                MdbCopyPage._collect_source_by_index(root, {"ВИ1"})


class MdbTransactionTests(unittest.TestCase):
    def test_failed_insert_rolls_back_delete(self):
        source_cursor = MagicMock()
        source_cursor.description = [("ID",), ("Value",)]
        source_cursor.fetchmany.side_effect = [[(1, "a")], []]

        target_cursor = MagicMock()
        target_cursor.description = [("ID",), ("Value",)]
        target_cursor.executemany.side_effect = RuntimeError("insert failed")

        source_connection = MagicMock()
        source_connection.cursor.return_value = source_cursor
        target_connection = MagicMock()
        target_connection.cursor.return_value = target_cursor

        connections = iter((source_connection, target_connection))
        page = SimpleNamespace(
            _same_file=lambda *_: False,
            _get_conn=lambda *_: next(connections),
            _quote_identifier=MdbCopyPage._quote_identifier,
        )

        with self.assertRaisesRegex(RuntimeError, "insert failed"):
            MdbCopyPage._copy_table(page, "source.mdb", "target.mdb", "Table")

        target_connection.rollback.assert_called_once_with()
        target_connection.commit.assert_not_called()
        source_connection.close.assert_called_once_with()
        target_connection.close.assert_called_once_with()

    def test_fias_mode_uses_relation_safe_address_copier(self):
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            source = root / "source.mdb"
            target = root / "target"
            source.touch()
            target.mkdir()
            target_mdb = target / "target.mdb"
            target_mdb.touch()

            safe_copier = MagicMock()
            copy_runner = MagicMock(return_value="ok")
            page = SimpleNamespace(
                _collect_all_mdb=lambda *_: [str(target_mdb)],
                _same_file=lambda *_: False,
                _copy_to_targets=copy_runner,
                _copy_locations_address=safe_copier,
            )

            result = MdbCopyPage._execute(
                page, MagicMock(), 3, (str(source), str(target))
            )

            self.assertEqual(result, "ok")
            copier = copy_runner.call_args.kwargs["copier"]
            copier("source.mdb", "target.mdb")
            safe_copier.assert_called_once_with(
                "source.mdb",
                "target.mdb",
                repair_missing_links=False,
            )

    def test_fias_updates_locations_through_link_table_not_guid_parameter(self):
        source_cursor = MagicMock()
        source_cursor.description = [
            ("ID",),
            ("Code_FIAS",),
            ("City_Name",),
            ("Document_ID",),
        ]
        source_cursor.fetchall.return_value = [
            ("SOURCE-ID", "FIAS-NEW", "Новый адрес", "SOURCE-DOCUMENT")
        ]

        target_cursor = MagicMock()
        target_cursor.fetchall.side_effect = [
            [("TARGET-ID",)],
            [("TARGET-ID",)],
            [("TARGET-ID", "FIAS-NEW", "Новый адрес")],
        ]
        source_connection = MagicMock()
        source_connection.cursor.return_value = source_cursor
        target_connection = MagicMock()
        target_connection.cursor.return_value = target_cursor
        connections = iter((source_connection, target_connection))
        page = SimpleNamespace(
            _same_file=lambda *_: False,
            _get_conn=lambda *_: next(connections),
            _quote_identifier=MdbCopyPage._quote_identifier,
        )

        updated = MdbCopyPage._copy_locations_address(
            page, "source.mdb", "target.mdb"
        )

        self.assertEqual(updated, (1, 0))
        update_calls = [
            call for call in target_cursor.execute.call_args_list
            if str(call.args[0]).startswith("UPDATE")
        ]
        self.assertEqual(len(update_calls), 1)
        update_sql = update_calls[0].args[0]
        self.assertIn("INNER JOIN [Местоположения_картаплан]", update_sql)
        self.assertIn("L.[ID]=M.[Location_ID]", update_sql)
        self.assertNotIn("WHERE [ID]=?", update_sql)
        self.assertEqual(update_calls[0].args[1], ("FIAS-NEW", "Новый адрес"))
        target_connection.commit.assert_called_once_with()
        target_connection.rollback.assert_not_called()


class MdbLocationsAuditTests(unittest.TestCase):
    def test_classifies_orphans_dangling_links_and_exact_duplicates(self):
        columns = ["ID", "City_Name", "Document_ID", "Insert_Date"]
        rows = [
            ("USED", "Адрес А", "DOC-1", "DATE-1"),
            ("ORPHAN-DUP", "Адрес А", "DOC-2", "DATE-2"),
            ("ORPHAN-OTHER", "Адрес Б", "DOC-3", "DATE-3"),
        ]
        references = {
            "used": {"value": "USED", "sources": {"Table.Location_ID"}},
            "missing": {
                "value": "MISSING",
                "sources": {"Broken.Location_ID"},
            },
        }

        report = MdbCopyPage._classify_location_rows(
            columns,
            rows,
            references,
        )

        self.assertEqual(report["total"], 3)
        self.assertEqual(report["used"], 1)
        self.assertEqual(
            {item[0] for item in report["orphaned"]},
            {"ORPHAN-DUP", "ORPHAN-OTHER"},
        )
        self.assertEqual(set(report["dangling"]), {"missing"})
        self.assertEqual(report["duplicates"], [["USED", "ORPHAN-DUP"]])

    def test_report_explicitly_describes_read_only_findings(self):
        report = {
            "total": 2,
            "used": 1,
            "orphaned": [("ORPHAN", "DOC")],
            "dangling": {
                "missing": {
                    "value": "MISSING",
                    "sources": {"Table.Location_ID"},
                }
            },
            "duplicates": [["A", "B"]],
            "scan_errors": [],
        }

        text = MdbCopyPage._format_locations_report("test.mdb", report)

        self.assertIn("Лишняя запись", text)
        self.assertIn("Оборванная ссылка", text)
        self.assertIn("Совпадающий адрес", text)

    def test_cleanup_deletes_only_locations_not_used_by_map_plan(self):
        cursor = MagicMock()
        cursor.fetchall.side_effect = [[], []]
        cursor.fetchone.side_effect = [(2,), (0,), (1,)]
        connection = MagicMock()
        connection.cursor.return_value = cursor
        page = SimpleNamespace(_get_conn=lambda *_: connection)

        result = MdbCopyPage._delete_unused_locations(page, "target.mdb")

        self.assertEqual(result, (2, 1))
        delete_calls = [
            call for call in cursor.execute.call_args_list
            if str(call.args[0]).startswith("DELETE L.* FROM [Locations]")
        ]
        self.assertEqual(len(delete_calls), 1)
        self.assertIn("M.[Location_ID] IS NULL", delete_calls[0].args[0])
        connection.commit.assert_called_once_with()
        connection.rollback.assert_not_called()

    def test_cleanup_refuses_database_with_dangling_map_plan_link(self):
        cursor = MagicMock()
        cursor.fetchall.return_value = [("BROKEN-ID",)]
        connection = MagicMock()
        connection.cursor.return_value = cursor
        page = SimpleNamespace(_get_conn=lambda *_: connection)

        with self.assertRaisesRegex(ValueError, "Сначала восстановите"):
            MdbCopyPage._delete_unused_locations(page, "target.mdb")

        connection.rollback.assert_called_once_with()
        connection.commit.assert_not_called()
        self.assertFalse(any(
            str(call.args[0]).startswith("DELETE")
            for call in cursor.execute.call_args_list
        ))

    def test_fias_emergency_mode_repairs_dangling_location_link(self):
        source_cursor = MagicMock()
        source_cursor.description = [
            ("ID",),
            ("Code_FIAS",),
            ("City_Name",),
            ("Document_ID",),
        ]
        source_cursor.fetchall.return_value = [
            ("SOURCE-ID", "FIAS-NEW", "Новый адрес", "SOURCE-DOCUMENT")
        ]

        target_cursor = MagicMock()
        target_cursor.rowcount = 1
        target_cursor.fetchall.side_effect = [
            [("BROKEN-ID",)],
            [],
            [("SOURCE-ID",)],
            [("SOURCE-ID",)],
            [("SOURCE-ID", "FIAS-NEW", "Новый адрес")],
        ]
        source_connection = MagicMock()
        source_connection.cursor.return_value = source_cursor
        target_connection = MagicMock()
        target_connection.cursor.return_value = target_cursor
        connections = iter((source_connection, target_connection))
        page = SimpleNamespace(
            _same_file=lambda *_: False,
            _get_conn=lambda *_: next(connections),
            _quote_identifier=MdbCopyPage._quote_identifier,
        )

        result = MdbCopyPage._copy_locations_address(
            page,
            "source.mdb",
            "target.mdb",
            repair_missing_links=True,
        )

        self.assertEqual(result, (1, 1))
        repair_calls = [
            call for call in target_cursor.execute.call_args_list
            if str(call.args[0]).startswith(
                "UPDATE [Местоположения_картаплан] SET [Location_ID]"
            )
        ]
        self.assertEqual(len(repair_calls), 1)
        self.assertEqual(repair_calls[0].args[1], ("SOURCE-ID",))
        target_connection.commit.assert_called_once_with()
        target_connection.rollback.assert_not_called()


if __name__ == "__main__":
    unittest.main()
