"""Страница оформления графического описания границ в Microsoft Word."""

from __future__ import annotations

from pathlib import Path

from PySide6.QtCore import QSignalBlocker, Qt
from PySide6.QtWidgets import (
    QComboBox,
    QFileDialog,
    QFormLayout,
    QHBoxLayout,
    QLabel,
    QMessageBox,
    QPushButton,
    QVBoxLayout,
    QWidget,
)

from core import BasePage, PathEdit
from core.word_description_formatter import (
    DocumentInfo,
    format_document,
    inspect_document,
    load_lo_addresses,
    word_automation_available,
)


class AddressComboBox(QComboBox):
    """Редактируемый справочник с ограниченной высотой выпадающего списка."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setEditable(True)
        self.setInsertPolicy(QComboBox.InsertPolicy.NoInsert)
        self.setMaxVisibleItems(12)
        self.setSizeAdjustPolicy(
            QComboBox.SizeAdjustPolicy.AdjustToMinimumContentsLengthWithIcon
        )
        self.setMinimumContentsLength(10)

    def showPopup(self):
        super().showPopup()
        view = self.view()
        visible_rows = min(max(self.count(), 1), self.maxVisibleItems())
        row_height = max(view.sizeHintForRow(0), view.fontMetrics().height() + 8)
        popup_height = visible_rows * row_height + 8
        view.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        view.setMaximumHeight(popup_height)
        view.window().setMaximumHeight(popup_height)


class LoDescriptionPage(BasePage):
    def __init__(self, controller=None, parent=None):
        super().__init__(controller, parent)
        self.addresses = load_lo_addresses()
        self._inspection_generation = 0
        self._inspecting = False
        self._formatting = False

        root = self.page_layout(
            "Описание ЛО",
            "Оформляет выгруженное графическое описание границ в отдельный файл Word.",
        )
        root.setAlignment(Qt.AlignmentFlag.AlignTop)

        if not word_automation_available():
            warning = QLabel(
                "Компонент pywin32 не установлен. Для работы нужны Windows, Microsoft Word "
                "и полная сборка K Tools. Остальные инструменты приложения доступны."
            )
            warning.setObjectName("warningBanner")
            warning.setWordWrap(True)
            root.addWidget(warning)

        _, file_card = self.card_layout(root, "Файлы Word")
        file_form = QFormLayout()
        file_form.setFieldGrowthPolicy(QFormLayout.FieldGrowthPolicy.AllNonFixedFieldsGrow)
        self.source_edit, self.source_button = self._file_field(
            file_form, "Исходный файл", "Выбрать…", self._choose_source
        )
        self.target_edit, self.target_button = self._file_field(
            file_form, "Сохранить как", "Выбрать…", self._choose_target
        )
        self.source_edit.setPlaceholderText("Файл .doc или .docx")
        self.target_edit.setPlaceholderText("Отдельный файл результата")
        self.source_edit.path_dropped.connect(self._source_selected)
        self.source_edit.editingFinished.connect(self._inspect_typed_source)
        file_card.addLayout(file_form)

        _, address_card = self.card_layout(root, "Адрес объекта")
        address_form = QFormLayout()
        address_form.setFieldGrowthPolicy(QFormLayout.FieldGrowthPolicy.AllNonFixedFieldsGrow)
        self.region_edit = PathEdit()
        self.region_edit.setText("Ленинградская область")
        address_form.addRow("Регион", self.region_edit)

        self.district_combo = AddressComboBox()
        self.district_combo.addItems(self.addresses)
        self.district_combo.setCurrentIndex(-1)
        self.district_combo.textActivated.connect(self._district_selected)
        self.district_combo.lineEdit().editingFinished.connect(
            lambda: self._load_settlements(preserve=True)
        )
        address_form.addRow("Район / округ", self.district_combo)

        self.settlement_combo = AddressComboBox()
        address_form.addRow("Поселение", self.settlement_combo)

        self.locality_edit = PathEdit()
        self.locality_edit.setPlaceholderText("Будет извлечён из исходного Word")
        address_form.addRow("Населённый пункт", self.locality_edit)
        address_card.addLayout(address_form)

        preview_layout = QVBoxLayout()
        preview_title = QLabel("Итоговый адрес")
        preview_title.setObjectName("helpSectionTitle")
        self.preview_label = QLabel()
        self.preview_label.setObjectName("helpSectionText")
        self.preview_label.setWordWrap(True)
        self.preview_label.setTextInteractionFlags(
            Qt.TextInteractionFlag.TextSelectableByMouse
        )
        preview_layout.addWidget(preview_title)
        preview_layout.addWidget(self.preview_label)
        address_card.addLayout(preview_layout)

        hint = QLabel(
            "Район и поселение Ленинградской области можно выбрать из списка. "
            "Все поля редактируются вручную; проверка по ФИАС не выполняется."
        )
        hint.setObjectName("pageSubtitle")
        hint.setWordWrap(True)
        address_card.addWidget(hint)

        _, action_card = self.card_layout(root)
        self.status_label = QLabel("Выберите исходный файл Word.")
        self.status_label.setWordWrap(True)
        action_card.addWidget(self.status_label)
        action_card.addWidget(self.setup_progress_bar())
        self.progress_bar.setVisible(False)
        self.run_button = QPushButton("Оформить графическое описание")
        self.run_button.setProperty("primary", True)
        self.run_button.clicked.connect(self._run)
        action_card.addWidget(self.run_button)

        for widget in (
            self.region_edit,
            self.district_combo,
            self.settlement_combo,
            self.locality_edit,
        ):
            if isinstance(widget, QComboBox):
                widget.editTextChanged.connect(self._update_preview)
            else:
                widget.textChanged.connect(self._update_preview)
        self._update_preview()
        self._refresh_controls()
        root.addStretch()

    @staticmethod
    def _file_field(form, label, button_text, callback):
        row = QWidget()
        layout = QHBoxLayout(row)
        layout.setContentsMargins(0, 0, 0, 0)
        edit = PathEdit()
        button = QPushButton(button_text)
        button.clicked.connect(callback)
        layout.addWidget(edit, 1)
        layout.addWidget(button)
        form.addRow(label, row)
        return edit, button

    def _choose_source(self):
        selected, _ = QFileDialog.getOpenFileName(
            self,
            "Выберите графическое описание",
            self.source_edit.text().strip(),
            "Документы Word (*.doc *.docx)",
        )
        if selected:
            self.source_edit.setText(selected)
            self._source_selected(selected)

    def _source_selected(self, selected: str):
        source = Path(selected)
        if source.suffix.casefold() not in {".doc", ".docx"}:
            QMessageBox.warning(
                self, "Неверный формат", "Выберите документ Word с расширением .doc или .docx."
            )
            return
        self.target_edit.setText(
            str(source.with_name(f"{source.stem}_оформлено{source.suffix}"))
        )
        self._start_inspection(source)

    def _inspect_typed_source(self):
        value = self.source_edit.text().strip()
        if value and Path(value).is_file():
            self._source_selected(value)

    def _start_inspection(self, source: Path):
        self._inspection_generation += 1
        generation = self._inspection_generation
        self._inspecting = True
        self.status_label.setText("Word читает адрес и название населённого пункта…")
        self._show_activity()
        self._refresh_controls()
        self.start_task(
            self._inspect_task,
            source,
            on_result=lambda info: self._inspection_done(generation, source, info),
            on_error=lambda text: self._inspection_failed(generation, text),
            on_finished=lambda: self._inspection_finished(generation),
        )

    @staticmethod
    def _inspect_task(_signals, source):
        return inspect_document(source)

    def _inspection_done(self, generation: int, source: Path, info: DocumentInfo):
        if generation != self._inspection_generation:
            return
        if Path(self.source_edit.text().strip()) != source:
            return
        self.locality_edit.setText(info.locality)
        parts = [part.strip() for part in info.address.split(",") if part.strip()]
        if parts:
            self.region_edit.setText(parts[0])
        if len(parts) > 1:
            self.district_combo.setEditText(parts[1])
            self._load_settlements(preserve=False)
        known_settlements = self.addresses.get(self.district_combo.currentText(), [])
        if len(parts) > 2 and (parts[2] in known_settlements or len(parts) > 3):
            self.settlement_combo.setEditText(parts[2])
        self.status_label.setText(
            f"Найдено координатных таблиц: {info.coordinate_tables}. "
            "Проверьте адрес и при необходимости исправьте населённый пункт."
        )

    def _inspection_failed(self, generation: int, traceback_text: str):
        if generation != self._inspection_generation:
            return
        message = self._error_message(traceback_text)
        self.status_label.setText("Не удалось прочитать исходный Word. Поля можно заполнить вручную.")
        QMessageBox.critical(self, "Не удалось прочитать документ", message)

    def _inspection_finished(self, generation: int):
        if generation != self._inspection_generation:
            return
        self._inspecting = False
        self._hide_activity()
        self._refresh_controls()

    def _choose_target(self):
        source = Path(self.source_edit.text().strip() or "описание.docx")
        suffix = source.suffix.casefold() if source.suffix.casefold() in {".doc", ".docx"} else ".docx"
        initial = self.target_edit.text().strip() or str(
            source.with_name(f"{source.stem}_оформлено{suffix}")
        )
        selected, _ = QFileDialog.getSaveFileName(
            self,
            "Сохранить оформленное описание",
            initial,
            f"Документ Word (*{suffix})",
        )
        if selected:
            path = Path(selected)
            if not path.suffix:
                path = path.with_suffix(suffix)
            self.target_edit.setText(str(path))

    def _district_selected(self, district: str):
        if district in self.addresses:
            self.region_edit.setText("Ленинградская область")
        self._load_settlements(preserve=False)

    def _load_settlements(self, *, preserve: bool):
        current = self.settlement_combo.currentText() if preserve else ""
        options = self.addresses.get(self.district_combo.currentText().strip(), [])
        blocker = QSignalBlocker(self.settlement_combo)
        self.settlement_combo.clear()
        self.settlement_combo.addItems(options)
        self.settlement_combo.setEditText(current)
        del blocker
        self._update_preview()

    def _update_preview(self, *_args):
        parts = (
            self.region_edit.text(),
            self.district_combo.currentText(),
            self.settlement_combo.currentText(),
            self.locality_edit.text(),
        )
        self.preview_label.setText(
            ", ".join(part.strip().rstrip(",") for part in parts if part.strip())
            or "Адрес пока не заполнен"
        )

    def _run(self):
        if self._formatting or self._inspecting:
            return
        values = (
            self.source_edit.text().strip(),
            self.target_edit.text().strip(),
            self.region_edit.text().strip(),
            self.district_combo.currentText().strip(),
            self.settlement_combo.currentText().strip(),
            self.locality_edit.text().strip(),
        )
        if not values[0] or not values[1]:
            QMessageBox.warning(
                self, "Выберите файлы", "Укажите исходный Word и отдельный файл результата."
            )
            return
        target = Path(values[1])
        if target.exists():
            answer = QMessageBox.question(
                self,
                "Заменить файл?",
                f"Файл уже существует:\n{target}\n\nЗаменить его после успешной обработки?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.No,
            )
            if answer != QMessageBox.StandardButton.Yes:
                return

        self._formatting = True
        self.status_label.setText("Microsoft Word обрабатывает документ и таблицы…")
        self._show_activity()
        self._refresh_controls()
        self.start_task(
            self._format_task,
            values,
            on_result=self._format_done,
            on_error=self._format_failed,
            on_finished=self._format_finished,
        )

    @staticmethod
    def _format_task(_signals, values):
        return format_document(*values)

    def _format_done(self, info: DocumentInfo):
        output = self.target_edit.text().strip()
        self.status_label.setText(f"Готово: {output}")
        QMessageBox.information(
            self,
            "Графическое описание оформлено",
            f"Файл сохранён:\n{output}\n\n"
            f"Адрес:\n{info.address}\n\n"
            f"Обработано координатных таблиц: {info.coordinate_tables}",
        )

    def _format_failed(self, traceback_text: str):
        self.status_label.setText("Ошибка обработки. Исходный файл не изменён.")
        QMessageBox.critical(
            self,
            "Не удалось оформить документ",
            self._error_message(traceback_text),
        )

    def _format_finished(self):
        self._formatting = False
        self._hide_activity()
        self._refresh_controls()

    def _show_activity(self):
        self.progress_bar.setVisible(True)
        self.progress_bar.setRange(0, 0)

    def _hide_activity(self):
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(100)
        self.progress_bar.setVisible(False)

    def _refresh_controls(self):
        busy = self._formatting
        self.source_button.setEnabled(not busy)
        self.target_button.setEnabled(not busy)
        self.source_edit.setEnabled(not busy)
        self.target_edit.setEnabled(not busy)
        self.region_edit.setEnabled(not busy)
        self.district_combo.setEnabled(not busy)
        self.settlement_combo.setEnabled(not busy)
        self.locality_edit.setEnabled(not busy)
        self.run_button.setEnabled(not busy and not self._inspecting)

    @staticmethod
    def _error_message(traceback_text: str) -> str:
        lines = [line.strip() for line in str(traceback_text).splitlines() if line.strip()]
        if not lines:
            return "Неизвестная ошибка Microsoft Word."
        last = lines[-1]
        if ": " in last:
            return last.split(": ", 1)[1]
        return last
