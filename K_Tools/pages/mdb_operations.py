import os
import re
import shutil
import tempfile

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QCheckBox,
    QComboBox,
    QFileDialog,
    QFrame,
    QHBoxLayout,
    QLabel,
    QListView,
    QMessageBox,
    QPushButton,
    QSizePolicy,
    QSplitter,
    QTabWidget,
    QVBoxLayout,
    QWidget,
)

from core import BasePage, IndexSelector, PathEdit
from core.settlement_names import SETTLEMENT_TYPE_GENITIVE

try:
    import pyodbc
    PYODBC_AVAILABLE = True
except ImportError:
    pyodbc = None
    PYODBC_AVAILABLE = False


_ADMINISTRATIVE_TYPES = {
    "с.п.": "сельского поселения",
    "с.п": "сельского поселения",
    "сельское поселение": "сельского поселения",
    "г.п.": "городского поселения",
    "г.п": "городского поселения",
    "городское поселение": "городского поселения",
    "г.о.": "городского округа",
    "г.о": "городского округа",
    "городской округ": "городского округа",
    "м.о.": "муниципального округа",
    "м.о": "муниципального округа",
    "муниципальный округ": "муниципального округа",
    "м.р-н": "муниципального района",
    "м.р-н.": "муниципального района",
    "муниципальный район": "муниципального района",
}

_LOCALITY_SPAN = re.compile(
    r"(?P<prefix>\bв\s+границах\s+)"
    r"(?P<locality>"
    r"[^,;\r\n]*?(?:сельского|городского)\s+поселения"
    r"|[^,;\r\n]*?(?:городского|муниципального)\s+округа"
    r"|[^,;\r\n]*?(?=\s+муниципального\s+образования\b)"
    r"|[^,;\r\n]*?(?=\s+[А-ЯЁ][А-ЯЁа-яё-]+\s+(?:муниципального\s+)?района\b)"
    r")",
    re.IGNORECASE,
)


def _normalise_type(value):
    return " ".join(str(value or "").strip().casefold().replace("ё", "е").split())


def _adjective_to_genitive(value):
    """Склоняет типичные прилагательные в названиях МО."""
    text = " ".join(str(value or "").split())
    endings = (
        ("ское", "ского"), ("цкое", "цкого"),
        ("ое", "ого"), ("ее", "его"),
        ("ский", "ского"), ("цкий", "цкого"),
        ("ый", "ого"), ("ой", "ого"), ("ий", "его"),
        ("ая", "ой"), ("яя", "ей"),
    )
    lowered = text.casefold()
    for ending, replacement in endings:
        if lowered.endswith(ending):
            return text[:-len(ending)] + replacement
    return text


def _locality_phrase(name, kind):
    name = " ".join(str(name or "").split())
    kind_key = _normalise_type(kind)
    administrative = _ADMINISTRATIVE_TYPES.get(kind_key)
    if administrative:
        return f"{_adjective_to_genitive(name)} {administrative}".strip()
    genitive = SETTLEMENT_TYPE_GENITIVE.get(kind_key)
    if genitive is None:
        genitive = SETTLEMENT_TYPE_GENITIVE.get(
            kind_key.rstrip("."),
            str(kind or "").strip(),
        )
    return f"{genitive} {name}".strip()


def _location_components(values):
    """Возвращает НП и его родительское поселение из строки Locations."""

    locality_name = values.get("locality_name")
    locality_kind = values.get("locality_type")
    city_name = values.get("city_name")
    city_kind = values.get("city_type")
    if locality_name and locality_kind:
        return locality_name, locality_kind, city_name, city_kind
    if city_name and city_kind:
        return city_name, city_kind, None, None
    return None, None, None, None


def _replace_locality_text(
    value,
    name=None,
    kind=None,
    outside=False,
    parent_name=None,
    parent_kind=None,
):
    """Меняет только НП после «в границах», сохраняя район и регион."""
    if not isinstance(value, str):
        return value, False
    match = _LOCALITY_SPAN.search(value)
    if match is None:
        return value, False
    suffix = value[match.end():]
    if outside:
        replacement = "" if re.match(
            r"\s+муниципального\s+образования\b", suffix, re.IGNORECASE
        ) else "муниципального образования"
    else:
        replacement = _locality_phrase(name, kind)
        parent = _locality_phrase(parent_name, parent_kind)
        if parent and _normalise_type(parent) != _normalise_type(replacement):
            replacement = f"{replacement} {parent}".strip()
    new_value = value[:match.start("locality")] + replacement + value[match.end("locality"):]
    return new_value, True


class TableComboBox(QComboBox):
    """Выпадающий список, который не меняет таблицу случайным колесом."""

    def wheelEvent(self, event):
        if self.view().isVisible():
            super().wheelEvent(event)
        else:
            event.ignore()

    def keyPressEvent(self, event):
        if (
            event.key() in (Qt.Key.Key_Up, Qt.Key.Key_Down)
            and not self.view().isVisible()
        ):
            self.showPopup()
            event.accept()
            return
        super().keyPressEvent(event)

    def showPopup(self):
        """Ограничивает высоту popup-окна независимо от стиля Windows."""
        super().showPopup()
        view = self.view()
        visible_rows = min(max(self.count(), 1), 12)
        row_height = max(view.sizeHintForRow(0), view.fontMetrics().height() + 8)
        popup_height = visible_rows * row_height + 8
        view.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        view.setMaximumHeight(popup_height)
        view.window().setMaximumHeight(popup_height)


class MdbCopyPage(BasePage):
    TYPE_GENITIVE = SETTLEMENT_TYPE_GENITIVE

    def __init__(self, controller=None, parent=None):
        super().__init__(controller, parent)
        root = self.page_layout(
            "Работа с MDB",
            "Массовая замена файлов и таблиц Microsoft Access.",
        )
        root.setSpacing(12)
        root.setAlignment(Qt.AlignmentFlag.AlignTop)
        if not PYODBC_AVAILABLE:
            warning = QLabel("pyodbc не установлен. Добавьте зависимость перед работой с MDB.")
            warning.setObjectName("warningBanner")
            root.addWidget(warning)

        self.content_splitter = QSplitter(Qt.Orientation.Vertical)
        self.content_splitter.setObjectName("mdbContentSplitter")
        self.content_splitter.setChildrenCollapsible(False)
        self.content_splitter.setHandleWidth(8)
        root.addWidget(self.content_splitter, 1)

        controls = QWidget()
        controls_layout = QVBoxLayout(controls)
        controls_layout.setContentsMargins(0, 0, 0, 0)
        controls_layout.setSpacing(12)

        self.tabs = QTabWidget()
        controls_layout.addWidget(self.tabs)
        self._build_replace_mdb()
        self._build_vri()
        self._build_replace_table()
        self._build_fias()
        self._build_text()
        self.tabs.setSizePolicy(
            QSizePolicy.Policy.Expanding,
            QSizePolicy.Policy.Fixed,
        )
        self.run_btn = QPushButton("Запустить операцию")
        self.run_btn.setProperty("primary", True)
        self.run_btn.clicked.connect(self.run_operation)
        controls_layout.addWidget(self.run_btn)
        controls_layout.addWidget(self.setup_progress_bar())
        self.content_splitter.addWidget(controls)

        log_card = QFrame()
        log_card.setObjectName("card")
        logs = QVBoxLayout(log_card)
        log_card.setSizePolicy(
            QSizePolicy.Policy.Expanding,
            QSizePolicy.Policy.Expanding,
        )
        logs.setContentsMargins(14, 10, 14, 10)
        logs.setSpacing(5)
        log_title = QLabel("Журнал")
        log_title.setObjectName("cardTitle")
        logs.addWidget(log_title)
        log_text = self.setup_log_area()
        log_text.setMinimumHeight(96)
        log_text.setMaximumHeight(16777215)
        log_text.setTextInteractionFlags(
            Qt.TextInteractionFlag.TextSelectableByMouse
            | Qt.TextInteractionFlag.TextSelectableByKeyboard
        )
        logs.addWidget(log_text)
        self.content_splitter.addWidget(log_card)
        self.content_splitter.setStretchFactor(0, 0)
        self.content_splitter.setStretchFactor(1, 1)
        self._fit_tabs()
        self.content_splitter.setSizes([controls.sizeHint().height(), 180])

    @staticmethod
    def _tab():
        """Обычная вкладка: вся форма видна без внутренней прокрутки."""
        widget = QWidget()
        layout = QVBoxLayout(widget)
        layout.setContentsMargins(14, 10, 14, 10)
        layout.setSpacing(7)
        layout.setAlignment(Qt.AlignmentFlag.AlignTop)
        return widget, layout

    @staticmethod
    def _note(text):
        note = QLabel(text)
        note.setWordWrap(True)
        return note

    def _fit_tabs(self):
        """Держит все режимы MDB одной высоты и без внутренней прокрутки."""
        if self.tabs.count() == 0:
            return
        available_width = max(320, self.tabs.width() - 8, self.width() - 64)
        content_heights = []
        for index in range(self.tabs.count()):
            widget = self.tabs.widget(index)
            content_layout = widget.layout()
            if content_layout is not None and content_layout.hasHeightForWidth():
                content_heights.append(content_layout.heightForWidth(available_width))
            else:
                content_heights.append(widget.sizeHint().height())
        height = max(content_heights) + self.tabs.tabBar().sizeHint().height() + 12
        self.tabs.setFixedHeight(max(150, height))
        self.layout().invalidate()
        self.updateGeometry()

    def resizeEvent(self, event):
        super().resizeEvent(event)
        self._fit_tabs()

    def _directory_row(self, layout, label, source_selector=None):
        layout.addWidget(self._note(label))
        row = QHBoxLayout()
        edit = PathEdit()
        button = QPushButton("Выбрать")
        if source_selector is None:
            button.clicked.connect(lambda: self._pick_dir(edit))
        else:
            button.clicked.connect(lambda: self._pick_source_dir(edit, source_selector))
        row.addWidget(edit, 1)
        row.addWidget(button)
        layout.addLayout(row)
        return edit

    def _file_row(self, layout, label, callback):
        layout.addWidget(self._note(label))
        row = QHBoxLayout()
        edit = PathEdit()
        button = QPushButton("Выбрать MDB")
        button.clicked.connect(lambda: callback(edit))
        row.addWidget(edit, 1)
        row.addWidget(button)
        layout.addLayout(row)
        return edit

    def _build_replace_mdb(self):
        tab, layout = self._tab()
        note = self._note("Полностью заменяет target MDB файлом из папки соответствующего индекса.")
        layout.addWidget(note)
        self.replace_selector = IndexSelector()
        self.replace_selector.listbox.setMaximumHeight(90)
        self.replace_source = self._directory_row(
            layout, "Source: папки-индексы с MDB", self.replace_selector
        )
        layout.addWidget(QLabel("Индексы для обработки"))
        layout.addWidget(self.replace_selector)
        self.replace_target = self._directory_row(
            layout, "Target: структура МО / НП / индекс / … / MDB"
        )
        self.tabs.addTab(tab, "Замена MDB")

    def _build_vri(self):
        tab, layout = self._tab()
        note = self._note("Копирует таблицу Utilizations_KP между MDB одинакового индекса.")
        layout.addWidget(note)
        self.vri_selector = IndexSelector()
        self.vri_selector.listbox.setMaximumHeight(90)
        self.vri_source = self._directory_row(
            layout, "Source: папки-индексы с MDB", self.vri_selector
        )
        layout.addWidget(QLabel("Индексы для обработки"))
        layout.addWidget(self.vri_selector)
        self.vri_target = self._directory_row(layout, "Target: папка с MDB")
        self.tabs.addTab(tab, "ВРИ")

    def _build_replace_table(self):
        tab, layout = self._tab()
        self.table_source = self._file_row(
            layout, "Source MDB", self._pick_table_source
        )
        self.table_target = self._directory_row(layout, "Target: папка с MDB")
        layout.addWidget(QLabel("Имя таблицы"))
        row = QHBoxLayout()
        self.table_combo = TableComboBox()
        self.table_combo.setView(QListView(self.table_combo))
        self.table_combo.setEditable(False)
        self.table_combo.setMaxVisibleItems(12)
        self.table_combo.setPlaceholderText("Выберите таблицу из списка")
        refresh = QPushButton("Обновить список")
        refresh.clicked.connect(self._load_tables)
        row.addWidget(self.table_combo, 1)
        row.addWidget(refresh)
        layout.addLayout(row)
        self.tabs.addTab(tab, "Одна таблица")

    def _build_fias(self):
        tab, layout = self._tab()
        layout.addWidget(self._note(
            "Переносит адрес из связанной строки Locations во все target MDB, "
            "сохраняя их ID, Document_ID и связи."
        ))
        self.fias_source = self._file_row(layout, "Source MDB", self._pick_mdb)
        self.fias_target = self._directory_row(layout, "Target: папка с MDB")
        self.fias_check_btn = QPushButton("Проверить Locations — без изменений")
        self.fias_check_btn.setToolTip(
            "Сравнить только Locations и Местоположения_картаплан. "
            "Данные MDB не изменяются."
        )
        self.fias_check_btn.clicked.connect(self.check_locations)
        self.fias_cleanup_btn = QPushButton("Удалить лишние адреса Locations")
        self.fias_cleanup_btn.setProperty("danger", True)
        self.fias_cleanup_btn.setToolTip(
            "Удалить строки Locations, ID которых не используется в "
            "Местоположения_картаплан.Location_ID."
        )
        self.fias_cleanup_btn.clicked.connect(self.cleanup_locations)
        address_actions = QHBoxLayout()
        address_actions.addWidget(self.fias_check_btn)
        address_actions.addWidget(self.fias_cleanup_btn)
        layout.addLayout(address_actions)
        self.fias_repair_links = QCheckBox(
            "Крайний случай — восстановить повреждённые связи Location_ID"
        )
        self.fias_repair_links.setToolTip(
            "Включайте только для MDB, в которых Location_ID ссылается на "
            "уже отсутствующую строку Locations. Перед запуском сделайте копию."
        )
        layout.addWidget(self.fias_repair_links)
        self.fias_repair_warning = self._note(
            "Аварийный режим перепривяжет Местоположения_картаплан к адресу "
            "из source MDB. Используйте только если обычный режим сообщил "
            "о повреждённой связи."
        )
        self.fias_repair_warning.setObjectName("warningBanner")
        self.fias_repair_warning.setVisible(False)
        layout.addWidget(self.fias_repair_warning)
        self.fias_repair_links.toggled.connect(
            lambda checked: (
                self.fias_repair_warning.setVisible(checked),
                self._fit_tabs(),
            )
        )
        self.tabs.addTab(tab, "Адрес по ФИАС")

    def _build_text(self):
        tab, layout = self._tab()
        layout.addWidget(
            self._note("Обновляет поле Объект_ЗУ таблицы Титульный_картаплан по Locations.")
        )
        self.text_folder = self._directory_row(layout, "Папка с MDB")
        self.text_outside = QCheckBox("Вне НП — без названия НП")
        self.text_outside.setToolTip(
            "Формировать текст без названия населённого пункта."
        )
        layout.addWidget(self.text_outside)
        self.tabs.addTab(tab, "Адрес в тексте")

    def _pick_dir(self, edit):
        path = QFileDialog.getExistingDirectory(self, "Выберите папку")
        if path:
            edit.setText(path)

    def _pick_source_dir(self, edit, selector):
        self._pick_dir(edit)
        root = edit.text().strip()
        if os.path.isdir(root):
            selector.load(sorted(
                name for name in os.listdir(root)
                if os.path.isdir(os.path.join(root, name))
            ))

    def _pick_mdb(self, edit):
        path, _ = QFileDialog.getOpenFileName(self, "Выберите MDB", "", "MDB (*.mdb)")
        if path:
            edit.setText(path)

    def _pick_table_source(self, edit):
        self._pick_mdb(edit)
        if edit.text():
            self._load_tables()

    def _load_tables(self):
        path = self.table_source.text().strip()
        if not PYODBC_AVAILABLE:
            QMessageBox.critical(self, "Ошибка", "pyodbc не установлен.")
            return
        if not os.path.isfile(path):
            QMessageBox.warning(self, "Внимание", "Сначала выберите MDB.")
            return
        connection = None
        try:
            connection = self._get_conn(path)
            tables = [row.table_name for row in connection.cursor().tables(tableType="TABLE")]
            self.table_combo.clear()
            self.table_combo.addItems(tables)
            self.log(f"Загружено таблиц: {len(tables)}")
        except Exception as error:
            QMessageBox.critical(self, "Ошибка", str(error))
        finally:
            if connection is not None:
                connection.close()

    @staticmethod
    def _get_conn(path):
        return pyodbc.connect(
            rf"DRIVER={{Microsoft Access Driver (*.mdb, *.accdb)}};DBQ={path};",
            autocommit=False,
        )

    @staticmethod
    def _collect_source_by_index(root, selected):
        result = {}
        for name in sorted(os.listdir(root), key=str.casefold):
            folder = os.path.join(root, name)
            if name not in selected or not os.path.isdir(folder):
                continue
            candidates = MdbCopyPage._collect_all_mdb(folder)
            if len(candidates) > 1:
                listed = "\n".join(f"  - {item}" for item in candidates)
                raise ValueError(
                    f"Для индекса {name} найдено несколько source MDB. "
                    f"Оставьте один файл:\n{listed}"
                )
            if candidates:
                result[name] = candidates[0]
        return result

    @staticmethod
    def _find_target_mdb_by_index(root, index_name):
        result = []
        for folder, _, files in os.walk(root):
            parts = {part.casefold() for part in os.path.normpath(folder).split(os.sep)}
            if index_name.casefold() in parts:
                result.extend(
                    os.path.join(folder, name)
                    for name in sorted(files, key=str.casefold)
                    if name.lower().endswith(".mdb")
                )
        return sorted(result, key=str.casefold)

    @staticmethod
    def _collect_all_mdb(root):
        return sorted([
            os.path.join(folder, name)
            for folder, _, files in os.walk(root)
            for name in files if name.lower().endswith(".mdb")
        ], key=str.casefold)

    @staticmethod
    def _same_file(first, second):
        try:
            return os.path.samefile(first, second)
        except (FileNotFoundError, OSError):
            normalise = lambda value: os.path.normcase(os.path.realpath(value))
            return normalise(first) == normalise(second)

    @staticmethod
    def _quote_identifier(value):
        value = str(value or "")
        if not value or "\x00" in value:
            raise ValueError("Пустое или недопустимое имя таблицы/поля")
        return f"[{value.replace(']', ']]')}]"

    @staticmethod
    def _replace_mdb_file(source, target):
        if MdbCopyPage._same_file(source, target):
            raise ValueError("Source и target указывают на один MDB")
        target_dir = os.path.dirname(os.path.abspath(target))
        descriptor, temporary = tempfile.mkstemp(
            prefix=".ktools-mdb-", suffix=".tmp", dir=target_dir
        )
        os.close(descriptor)
        try:
            shutil.copy2(source, temporary)
            if os.path.getsize(temporary) <= 0:
                raise OSError("Скопированный MDB пуст")
            os.replace(temporary, target)
        finally:
            if os.path.exists(temporary):
                os.unlink(temporary)

    def _copy_table(self, source, target, table):
        if self._same_file(source, target):
            raise ValueError("Source и target указывают на один MDB")
        source_conn = target_conn = None
        try:
            source_conn = self._get_conn(source)
            target_conn = self._get_conn(target)
            source_cursor, target_cursor = source_conn.cursor(), target_conn.cursor()
            quoted_table = self._quote_identifier(table)
            source_cursor.execute(f"SELECT * FROM {quoted_table}")
            source_columns = [description[0] for description in source_cursor.description]
            target_cursor.execute(f"SELECT * FROM {quoted_table} WHERE 1=0")
            target_columns = [description[0] for description in target_cursor.description]
            if [item.casefold() for item in source_columns] != [
                item.casefold() for item in target_columns
            ]:
                raise ValueError(
                    f"Схемы таблицы {table} различаются: "
                    f"source полей {len(source_columns)}, target полей {len(target_columns)}"
                )

            names = ", ".join(self._quote_identifier(name) for name in target_columns)
            placeholders = ", ".join("?" for _ in target_columns)
            insert_sql = f"INSERT INTO {quoted_table} ({names}) VALUES ({placeholders})"
            target_cursor.execute(f"DELETE FROM {quoted_table}")
            while True:
                rows = source_cursor.fetchmany(500)
                if not rows:
                    break
                target_cursor.executemany(insert_sql, rows)
            target_conn.commit()
        except Exception:
            if target_conn is not None:
                target_conn.rollback()
            raise
        finally:
            if source_conn is not None:
                source_conn.close()
            if target_conn is not None:
                target_conn.close()

    def _copy_locations_address(self, source, target, repair_missing_links=False):
        """Переносит адрес, не ломая target Location_ID и Document_ID."""
        if self._same_file(source, target):
            raise ValueError("Source и target указывают на один MDB")
        source_conn = target_conn = None
        link_table = self._quote_identifier("Местоположения_картаплан")
        locations = self._quote_identifier("Locations")
        try:
            source_conn = self._get_conn(source)
            target_conn = self._get_conn(target)
            source_cursor = source_conn.cursor()
            target_cursor = target_conn.cursor()
            source_cursor.execute(
                f"SELECT L.* FROM {locations} AS L "
                f"INNER JOIN {link_table} AS M ON L.[ID]=M.[Location_ID]"
            )
            columns = [description[0] for description in source_cursor.description]
            source_rows = source_cursor.fetchall()
            if not source_rows:
                raise ValueError(
                    "В source MDB нет адреса Locations, связанного через "
                    "Местоположения_картаплан.Location_ID"
                )

            preserved = {"id", "document_id", "insert_date"}
            copied_columns = [
                column for column in columns
                if str(column).strip().casefold() not in preserved
            ]
            indexes = {column: columns.index(column) for column in copied_columns}
            variants = {
                tuple(row[indexes[column]] for column in copied_columns)
                for row in source_rows
            }
            if len(variants) != 1:
                raise ValueError(
                    "В source MDB найдено несколько разных связанных адресов. "
                    "Оставьте один адрес-источник."
                )
            source_values = next(iter(variants))
            id_index = next(
                index for index, column in enumerate(columns)
                if str(column).strip().casefold() == "id"
            )
            source_by_id = {
                str(row[id_index]).casefold(): row for row in source_rows
            }

            target_cursor.execute(
                f"SELECT DISTINCT [Location_ID] FROM {link_table} "
                "WHERE [Location_ID] IS NOT NULL"
            )
            referenced_ids = {
                str(row[0]).casefold(): row[0]
                for row in target_cursor.fetchall()
            }
            if not referenced_ids:
                raise ValueError(
                    "В target MDB нет заполненных "
                    "Местоположения_картаплан.Location_ID"
                )

            target_cursor.execute(
                f"SELECT DISTINCT M.[Location_ID] FROM {link_table} AS M "
                f"INNER JOIN {locations} AS L ON L.[ID]=M.[Location_ID] "
                "WHERE M.[Location_ID] IS NOT NULL"
            )
            target_ids = {str(row[0]).casefold() for row in target_cursor.fetchall()}
            missing_ids = set(referenced_ids).difference(target_ids)
            repaired_links = 0
            if missing_ids:
                if not repair_missing_links:
                    raise ValueError(
                        "В target MDB повреждена связь "
                        "Местоположения_картаплан.Location_ID → Locations.ID. "
                        "Для восстановления включите «Крайний случай»"
                    )
                if len(source_by_id) != 1:
                    raise ValueError(
                        "Аварийное восстановление требует ровно одну связанную "
                        "строку Locations в source MDB"
                    )
                source_key, source_row = next(iter(source_by_id.items()))
                source_location_id = source_row[id_index]

                target_cursor.execute(f"SELECT [ID] FROM {locations}")
                existing_ids = {
                    str(row[0]).casefold() for row in target_cursor.fetchall()
                }
                if source_key not in existing_ids:
                    names = ", ".join(
                        self._quote_identifier(column) for column in columns
                    )
                    placeholders = ", ".join("?" for _ in columns)
                    target_cursor.execute(
                        f"INSERT INTO {locations} ({names}) VALUES ({placeholders})",
                        tuple(source_row),
                    )

                target_cursor.execute(
                    f"UPDATE {link_table} SET [Location_ID]=? "
                    "WHERE [Location_ID] IS NOT NULL",
                    (source_location_id,),
                )
                repaired_links = max(target_cursor.rowcount, len(missing_ids))
                target_cursor.execute(
                    f"SELECT DISTINCT M.[Location_ID] FROM {link_table} AS M "
                    f"INNER JOIN {locations} AS L ON L.[ID]=M.[Location_ID] "
                    "WHERE M.[Location_ID] IS NOT NULL"
                )
                target_ids = {
                    str(row[0]).casefold() for row in target_cursor.fetchall()
                }
                if target_ids != {source_key}:
                    raise RuntimeError(
                        "Не удалось восстановить связь Location_ID с Locations.ID"
                    )

            assignments = ", ".join(
                f"L.{self._quote_identifier(column)}=?" for column in copied_columns
            )
            update_sql = (
                f"UPDATE {locations} AS L INNER JOIN {link_table} AS M "
                f"ON L.[ID]=M.[Location_ID] SET {assignments}"
            )
            target_cursor.execute(update_sql, source_values)

            selected_columns = ", ".join(
                f"L.{self._quote_identifier(column)}" for column in copied_columns
            )
            target_cursor.execute(
                f"SELECT L.[ID], {selected_columns} FROM {locations} AS L "
                f"INNER JOIN {link_table} AS M ON L.[ID]=M.[Location_ID]"
            )
            verified_ids = set()
            for row in target_cursor.fetchall():
                verified_ids.add(str(row[0]).casefold())
                if tuple(row[1:]) != source_values:
                    raise RuntimeError(
                        "Access не сохранил адрес в связанной строке Locations"
                    )
            if verified_ids != target_ids:
                raise RuntimeError(
                    "Не все связанные строки Locations удалось проверить после обновления"
                )
            target_conn.commit()
            return len(verified_ids), repaired_links
        except Exception:
            if target_conn is not None:
                target_conn.rollback()
            raise
        finally:
            if source_conn is not None:
                source_conn.close()
            if target_conn is not None:
                target_conn.close()

    def _get_locality(self, path):
        connection = self._get_conn(path)
        try:
            cursor = connection.cursor()
            cursor.execute(
                "SELECT L.* FROM [Locations] AS L "
                "INNER JOIN [Местоположения_картаплан] AS M "
                "ON L.[ID]=M.[Location_ID]"
            )
            columns = [description[0] for description in cursor.description]
            for row in cursor.fetchall():
                values = {
                    str(column).strip().casefold(): value
                    for column, value in zip(columns, row)
                }
                name, kind, parent_name, parent_kind = _location_components(values)
                if name and kind:
                    return tuple(
                        str(value).strip() if value else None
                        for value in (name, kind, parent_name, parent_kind)
                    )
            return None, None, None, None
        finally:
            connection.close()

    @staticmethod
    def _classify_location_rows(columns, rows, references):
        """Сравнивает Locations только со ссылками Местоположения_картаплан."""
        by_key = {
            str(column).strip().casefold(): index
            for index, column in enumerate(columns)
        }
        if "id" not in by_key:
            raise ValueError("В таблице Locations нет поля ID")
        id_index = by_key["id"]
        document_index = by_key.get("document_id")
        ignored = {"id", "document_id", "insert_date"}
        address_indexes = [
            index for index, column in enumerate(columns)
            if str(column).strip().casefold() not in ignored
        ]

        def hashable(value):
            try:
                hash(value)
                return value
            except TypeError:
                return bytes(value) if isinstance(value, bytearray) else repr(value)

        locations = {}
        duplicate_candidates = {}
        for row in rows:
            location_id = row[id_index]
            key = str(location_id).casefold()
            document_id = row[document_index] if document_index is not None else None
            locations[key] = (location_id, document_id)
            signature = tuple(hashable(row[index]) for index in address_indexes)
            duplicate_candidates.setdefault(signature, []).append(location_id)

        location_keys = set(locations)
        reference_keys = set(references)
        orphaned = [
            locations[key] for key in sorted(location_keys - reference_keys)
        ]
        dangling = {
            key: references[key]
            for key in sorted(reference_keys - location_keys)
        }
        duplicates = [
            ids for ids in duplicate_candidates.values() if len(ids) > 1
        ]
        return {
            "total": len(rows),
            "used": len(location_keys & reference_keys),
            "orphaned": orphaned,
            "dangling": dangling,
            "duplicates": duplicates,
        }

    def _inspect_locations(self, path):
        """Только читает Locations и Местоположения_картаплан."""
        connection = self._get_conn(path)
        try:
            cursor = connection.cursor()
            references = {}
            cursor.execute(
                "SELECT DISTINCT [Location_ID] "
                "FROM [Местоположения_картаплан] "
                "WHERE [Location_ID] IS NOT NULL"
            )
            for row in cursor.fetchall():
                references[str(row[0]).casefold()] = {
                    "value": row[0],
                    "sources": {"Местоположения_картаплан.Location_ID"},
                }

            cursor.execute("SELECT * FROM [Locations]")
            columns = [item[0] for item in cursor.description]
            rows = cursor.fetchall()
            classified = self._classify_location_rows(columns, rows, references)
            classified["scan_errors"] = []
            return classified
        finally:
            connection.close()

    @staticmethod
    def _format_locations_report(path, report):
        lines = [
            f"{path}",
            "  Locations: "
            f"всего {report['total']}, используются {report['used']}, "
            f"лишних {len(report['orphaned'])}, "
            f"оборванных ссылок {len(report['dangling'])}, "
            f"групп совпадающих адресов {len(report['duplicates'])}",
        ]
        for location_id, document_id in report["orphaned"]:
            lines.append(
                f"  Лишняя запись: ID={location_id}; Document_ID={document_id}"
            )
        for entry in report["dangling"].values():
            sources = ", ".join(sorted(entry["sources"], key=str.casefold))
            lines.append(
                f"  Оборванная ссылка: ID={entry['value']}; из {sources}"
            )
        for ids in report["duplicates"]:
            lines.append(
                "  Совпадающий адрес (может быть штатным): "
                + ", ".join(map(str, ids))
            )
        for error in report["scan_errors"]:
            lines.append(f"  Не удалось проверить ссылку: {error}")
        return "\n".join(lines)

    def check_locations(self):
        root = self.fias_target.text().strip()
        if not os.path.isdir(root):
            QMessageBox.warning(self, "Внимание", "Укажите корректную target-папку с MDB")
            return
        self.clear_log()
        if self.progress_bar is not None:
            self.progress_bar.setValue(0)
        self._set_action_buttons_enabled(False)
        self.start_task(
            self._execute_locations_check,
            root,
            on_result=lambda result: QMessageBox.information(self, "Проверка Locations", result),
            on_error=self._show_error,
            on_finished=lambda: self._set_action_buttons_enabled(True),
        )

    def _execute_locations_check(self, signals, root):
        files = self._collect_all_mdb(root)
        if not files:
            raise ValueError("MDB-файлы не найдены")
        failed = 0
        totals = {"orphaned": 0, "dangling": 0, "duplicates": 0}
        for position, path in enumerate(files, 1):
            try:
                report = self._inspect_locations(path)
                totals["orphaned"] += len(report["orphaned"])
                totals["dangling"] += len(report["dangling"])
                totals["duplicates"] += len(report["duplicates"])
                signals.message.emit(self._format_locations_report(path, report))
            except Exception as error:
                failed += 1
                signals.message.emit(f"Не удалось проверить {path}: {error}")
            signals.progress.emit(position, len(files))
        return (
            f"Проверено MDB: {len(files)}; ошибок чтения: {failed}; "
            f"лишних записей: {totals['orphaned']}; "
            f"оборванных ссылок: {totals['dangling']}; "
            f"групп совпадающих адресов: {totals['duplicates']}. "
            "Никакие данные не изменялись."
        )

    def _delete_unused_locations(self, path):
        """Удаляет только Locations.ID, отсутствующие в таблице связей."""
        connection = self._get_conn(path)
        try:
            cursor = connection.cursor()
            dangling_sql = (
                "SELECT DISTINCT M.[Location_ID] "
                "FROM [Местоположения_картаплан] AS M "
                "LEFT JOIN [Locations] AS L ON L.[ID]=M.[Location_ID] "
                "WHERE M.[Location_ID] IS NOT NULL AND L.[ID] IS NULL"
            )
            cursor.execute(dangling_sql)
            dangling = [row[0] for row in cursor.fetchall()]
            if dangling:
                raise ValueError(
                    "Сначала восстановите оборванную связь Location_ID: "
                    + ", ".join(map(str, dangling))
                )

            unused_count_sql = (
                "SELECT COUNT(*) FROM [Locations] AS L "
                "LEFT JOIN [Местоположения_картаплан] AS M "
                "ON L.[ID]=M.[Location_ID] WHERE M.[Location_ID] IS NULL"
            )
            cursor.execute(unused_count_sql)
            unused_before = cursor.fetchone()[0]
            if unused_before:
                cursor.execute(
                    "DELETE L.* FROM [Locations] AS L "
                    "LEFT JOIN [Местоположения_картаплан] AS M "
                    "ON L.[ID]=M.[Location_ID] WHERE M.[Location_ID] IS NULL"
                )

            cursor.execute(dangling_sql)
            if cursor.fetchall():
                raise RuntimeError("После очистки обнаружена оборванная связь Location_ID")
            cursor.execute(unused_count_sql)
            if cursor.fetchone()[0] != 0:
                raise RuntimeError("Access удалил не все лишние строки Locations")
            cursor.execute("SELECT COUNT(*) FROM [Locations]")
            remaining = cursor.fetchone()[0]
            connection.commit()
            return unused_before, remaining
        except Exception:
            connection.rollback()
            raise
        finally:
            connection.close()

    def cleanup_locations(self):
        root = self.fias_target.text().strip()
        if not os.path.isdir(root):
            QMessageBox.warning(self, "Внимание", "Укажите корректную target-папку с MDB")
            return
        answer = QMessageBox.warning(
            self,
            "Удаление лишних адресов",
            "Будут удалены все строки Locations, ID которых не указан в "
            "Местоположения_картаплан.Location_ID.\n\n"
            "Перед продолжением сделайте резервную копию MDB. Продолжить?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )
        if answer != QMessageBox.StandardButton.Yes:
            return
        self.clear_log()
        if self.progress_bar is not None:
            self.progress_bar.setValue(0)
        self._set_action_buttons_enabled(False)
        self.start_task(
            self._execute_locations_cleanup,
            root,
            on_result=lambda result: QMessageBox.information(self, "Очистка Locations", result),
            on_error=self._show_error,
            on_finished=lambda: self._set_action_buttons_enabled(True),
        )

    def _execute_locations_cleanup(self, signals, root):
        files = self._collect_all_mdb(root)
        if not files:
            raise ValueError("MDB-файлы не найдены")
        deleted = 0
        changed = 0
        failed = 0
        for position, path in enumerate(files, 1):
            try:
                removed, remaining = self._delete_unused_locations(path)
                deleted += removed
                changed += int(removed > 0)
                signals.message.emit(
                    f"OK: {path} — удалено {removed}, осталось {remaining}"
                )
            except Exception as error:
                failed += 1
                signals.message.emit(f"Ошибка очистки: {path}: {error}")
            signals.progress.emit(position, len(files))
        if failed:
            raise RuntimeError(
                f"Очистка выполнена частично: успешно {len(files) - failed}, "
                f"ошибок {failed}, удалено строк {deleted}. Подробности — в журнале."
            )
        return (
            f"Проверено MDB: {len(files)}; изменено MDB: {changed}; "
            f"удалено лишних строк Locations: {deleted}."
        )

    def _set_action_buttons_enabled(self, enabled):
        self.run_btn.setEnabled(enabled)
        self.fias_check_btn.setEnabled(enabled)
        self.fias_cleanup_btn.setEnabled(enabled)

    def _update_title(
        self,
        path,
        name,
        kind,
        outside=False,
        parent_name=None,
        parent_kind=None,
    ):
        connection, updated = self._get_conn(path), 0
        try:
            cursor = connection.cursor()
            cursor.execute("SELECT * FROM [Титульный_картаплан]")
            columns = [description[0] for description in cursor.description]
            by_key = {str(column).strip().casefold(): column for column in columns}
            field = by_key.get("объект_зу")
            if field is None:
                raise ValueError(
                    "В таблице Титульный_картаплан нет поля Объект_ЗУ"
                )
            primary = by_key.get("id", columns[0])
            matched = False
            for row in cursor.fetchall():
                values = dict(zip(columns, row))
                old = values.get(field)
                new, found = _replace_locality_text(
                    old,
                    name,
                    kind,
                    outside,
                    parent_name,
                    parent_kind,
                )
                matched = matched or found
                if found and new != old:
                    cursor.execute(
                        f"UPDATE {self._quote_identifier('Титульный_картаплан')} "
                        f"SET {self._quote_identifier(field)}=? "
                        f"WHERE {self._quote_identifier(primary)}=?",
                        (new, values[primary]),
                    )
                    updated += 1
            if not matched:
                raise ValueError(
                    "В Объект_ЗУ не найден поддерживаемый фрагмент после «в границах»"
                )
            connection.commit()
            return updated
        except Exception:
            connection.rollback()
            raise
        finally:
            connection.close()

    def run_operation(self):
        if not PYODBC_AVAILABLE:
            QMessageBox.critical(self, "Ошибка", "Установите pyodbc.")
            return
        mode = self.tabs.currentIndex()
        if mode == 0:
            params = (
                self.replace_source.text().strip(),
                self.replace_target.text().strip(),
                set(self.replace_selector.get_selected()),
            )
        elif mode == 1:
            params = (
                self.vri_source.text().strip(),
                self.vri_target.text().strip(),
                set(self.vri_selector.get_selected()),
            )
        elif mode == 2:
            params = (
                self.table_source.text().strip(),
                self.table_target.text().strip(),
                self.table_combo.currentText().strip(),
            )
        elif mode == 3:
            repair_links = self.fias_repair_links.isChecked()
            if repair_links:
                answer = QMessageBox.warning(
                    self,
                    "Аварийное восстановление MDB",
                    "Этот режим предназначен только для повреждённых связей. "
                    "Он изменит Location_ID в таблице Местоположения_картаплан.\n\n"
                    "Перед продолжением рекомендуется сделать резервную копию MDB. "
                    "Продолжить?",
                    QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                    QMessageBox.StandardButton.No,
                )
                if answer != QMessageBox.StandardButton.Yes:
                    return
            params = (
                self.fias_source.text().strip(),
                self.fias_target.text().strip(),
                repair_links,
            )
        else:
            params = (self.text_folder.text().strip(), self.text_outside.isChecked())
        self.clear_log()
        if self.progress_bar is not None:
            self.progress_bar.setValue(0)
        self._set_action_buttons_enabled(False)
        self.start_task(
            self._execute,
            mode,
            params,
            on_result=lambda result: QMessageBox.information(self, "Готово", result),
            on_error=self._show_error,
            on_finished=lambda: self._set_action_buttons_enabled(True),
        )

    def _execute(self, signals, mode, params):
        if mode in (0, 1):
            source_root, target_root, selected = params
            if not os.path.isdir(source_root) or not os.path.isdir(target_root):
                raise ValueError("Укажите корректные папки SOURCE и TARGET")
            if not selected:
                raise ValueError("Выберите хотя бы один индекс")
            sources = self._collect_source_by_index(source_root, selected)
            missing = sorted(selected.difference(sources), key=str.casefold)
            if missing:
                raise ValueError(
                    "SOURCE MDB не найдены для индексов: " + ", ".join(missing)
                )
            jobs = []
            without_targets = []
            for index, source in sorted(sources.items(), key=lambda item: item[0].casefold()):
                targets = [
                    target for target in self._find_target_mdb_by_index(target_root, index)
                    if not self._same_file(source, target)
                ]
                if not targets:
                    without_targets.append(index)
                jobs.append((index, source, targets))
            if without_targets:
                raise ValueError(
                    "Target MDB не найдены для индексов: "
                    + ", ".join(without_targets)
                )
            changed = 0
            failed = 0
            for position, (index, source, targets) in enumerate(jobs, 1):
                signals.message.emit(f"Индекс {index}: target файлов — {len(targets)}")
                for target in targets:
                    try:
                        if mode == 0:
                            self._replace_mdb_file(source, target)
                        else:
                            self._copy_table(source, target, "Utilizations_KP")
                        changed += 1
                        signals.message.emit(f"  OK: {target}")
                    except Exception as error:
                        failed += 1
                        signals.message.emit(f"  Ошибка: {target}: {error}")
                signals.progress.emit(position, len(sources))
            if failed:
                raise RuntimeError(
                    f"Операция выполнена частично: успешно {changed}, ошибок {failed}. "
                    "Подробности — в журнале."
                )
            return f"Операция завершена. Обновлено: {changed}"

        if mode == 2:
            source, target_root, table = params
            if not os.path.isfile(source) or not os.path.isdir(target_root) or not table:
                raise ValueError("Проверьте source MDB, target папку и имя таблицы")
            targets = [
                item for item in self._collect_all_mdb(target_root)
                if not self._same_file(item, source)
            ]
            return self._copy_to_targets(signals, source, targets, table)

        if mode == 3:
            source, target_root, *options = params
            repair_links = bool(options[0]) if options else False
            if not os.path.isfile(source) or not os.path.isdir(target_root):
                raise ValueError("Проверьте source MDB и target папку")
            targets = [
                item for item in self._collect_all_mdb(target_root)
                if not self._same_file(item, source)
            ]
            return self._copy_to_targets(
                signals, source, targets, "адрес Locations",
                copier=lambda source_path, target_path: self._copy_locations_address(
                    source_path,
                    target_path,
                    repair_missing_links=repair_links,
                ),
            )

        folder, outside = params
        if not os.path.isdir(folder):
            raise ValueError("Укажите корректную папку с MDB")
        files = self._collect_all_mdb(folder)
        if not files:
            raise ValueError("MDB-файлы не найдены")
        total_updated = 0
        failed = 0
        for position, path in enumerate(files, 1):
            try:
                if outside:
                    updated = self._update_title(path, None, None, True)
                else:
                    name, kind, parent_name, parent_kind = self._get_locality(path)
                    if not name:
                        signals.message.emit(f"НП не найден: {path}")
                        signals.progress.emit(position, len(files))
                        continue
                    updated = self._update_title(
                        path,
                        name,
                        kind,
                        parent_name=parent_name,
                        parent_kind=parent_kind,
                    )
                total_updated += updated
                signals.message.emit(f"{os.path.basename(path)}: обновлено строк {updated}")
            except Exception as error:
                failed += 1
                signals.message.emit(f"{path}: {error}")
            signals.progress.emit(position, len(files))
        if failed:
            raise RuntimeError(
                f"Адрес обновлён частично: строк {total_updated}, ошибок MDB {failed}. "
                "Подробности — в журнале."
            )
        return f"Операция завершена. Обновлено строк: {total_updated}"

    def _copy_to_targets(self, signals, source, targets, table, copier=None):
        if not targets:
            raise ValueError("Target MDB-файлы не найдены")
        copier = copier or (lambda source_path, target_path: self._copy_table(
            source_path, target_path, table
        ))
        changed = 0
        failed = 0
        for position, target in enumerate(targets, 1):
            try:
                result = copier(source, target)
                changed += 1
                details = ""
                if isinstance(result, tuple) and len(result) == 2:
                    updated_rows, repaired_links = result
                    details = f" — строк: {updated_rows}"
                    if repaired_links:
                        details += f", аварийно восстановлено связей: {repaired_links}"
                signals.message.emit(f"OK: {target}{details}")
            except Exception as error:
                failed += 1
                signals.message.emit(f"Ошибка: {target}: {error}")
            signals.progress.emit(position, len(targets))
        if failed:
            raise RuntimeError(
                f"{table}: успешно {changed}, ошибок {failed}. Подробности — в журнале."
            )
        return f"Таблица {table} обновлена в {changed} MDB"

    def _show_error(self, text):
        self.log(text)
        QMessageBox.critical(self, "Ошибка", text.splitlines()[-1])
