"""Страница оформления графического описания границ в Microsoft Word."""

from __future__ import annotations

from pathlib import Path

from PySide6.QtCore import QSignalBlocker, Qt
from PySide6.QtWidgets import (
    QAbstractItemView,
    QComboBox,
    QFileDialog,
    QFormLayout,
    QHeaderView,
    QHBoxLayout,
    QLabel,
    QMessageBox,
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from core import BasePage, PathEdit
from core.word_description_formatter import (
    BatchFormatResult,
    DocumentInfo,
    build_full_address,
    find_batch_documents,
    format_description_documents,
    format_document,
    is_zone_folder_name,
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
        self._batch_documents = ()

        root = self.page_layout(
            "Описание ЛО",
            "Оформляет одно или сразу несколько графических описаний границ в Word.",
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
        (
            self.source_edit,
            self.source_button,
            self.source_folder_button,
        ) = self._source_field(file_form)
        self.target_edit, self.target_button = self._file_field(
            file_form, "Сохранить как", "Выбрать…", self._choose_target
        )
        self.source_edit.setPlaceholderText("Файл .doc/.docx или корневая папка")
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

        self.batch_preview_card, batch_preview = self.card_layout(
            root, "Предпросмотр пакетной обработки"
        )
        preview_hint = QLabel(
            "Снимите флажок у документа, который не нужно изменять. "
            "Адрес показан в том виде, в котором будет записан в Word."
        )
        preview_hint.setObjectName("pageSubtitle")
        preview_hint.setWordWrap(True)
        batch_preview.addWidget(preview_hint)

        self.batch_table = QTableWidget(0, 5)
        self.batch_table.setHorizontalHeaderLabels(
            ["✓", "Исходный Word", "Тип", "Итоговый адрес", "Результат"]
        )
        self.batch_table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.batch_table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.batch_table.setAlternatingRowColors(True)
        self.batch_table.verticalHeader().setVisible(False)
        self.batch_table.setMinimumHeight(190)
        self.batch_table.setMaximumHeight(280)
        header = self.batch_table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        header.setSectionResizeMode(2, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(3, QHeaderView.ResizeMode.Stretch)
        header.setSectionResizeMode(4, QHeaderView.ResizeMode.Stretch)
        self.batch_table.itemChanged.connect(self._batch_selection_changed)
        batch_preview.addWidget(self.batch_table)

        batch_actions = QHBoxLayout()
        self.batch_select_all_button = QPushButton("Выбрать все")
        self.batch_clear_button = QPushButton("Снять все")
        self.batch_refresh_button = QPushButton("Обновить список")
        self.batch_select_all_button.clicked.connect(lambda: self._set_all_batch_checked(True))
        self.batch_clear_button.clicked.connect(lambda: self._set_all_batch_checked(False))
        self.batch_refresh_button.clicked.connect(self._refresh_batch_documents)
        batch_actions.addWidget(self.batch_select_all_button)
        batch_actions.addWidget(self.batch_clear_button)
        batch_actions.addStretch()
        batch_actions.addWidget(self.batch_refresh_button)
        batch_preview.addLayout(batch_actions)
        self.batch_preview_card.setVisible(False)

        _, action_card = self.card_layout(root)
        self.status_label = QLabel("Выберите исходный Word или корневую папку.")
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

    def _source_field(self, form):
        row = QWidget()
        layout = QHBoxLayout(row)
        layout.setContentsMargins(0, 0, 0, 0)
        edit = PathEdit()
        file_button = QPushButton("Файл…")
        folder_button = QPushButton("Папка…")
        file_button.clicked.connect(self._choose_source)
        folder_button.clicked.connect(self._choose_source_folder)
        layout.addWidget(edit, 1)
        layout.addWidget(file_button)
        layout.addWidget(folder_button)
        form.addRow("Источник", row)
        return edit, file_button, folder_button

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

    def _choose_source_folder(self):
        selected = QFileDialog.getExistingDirectory(
            self,
            "Выберите корневую папку с папками населённых пунктов",
            self.source_edit.text().strip(),
        )
        if selected:
            self.source_edit.setText(selected)
            self._source_selected(selected)

    def _source_selected(self, selected: str):
        source = Path(selected)
        if source.is_dir():
            self._inspection_generation += 1
            self._inspecting = False
            self.target_edit.clear()
            self.target_edit.setPlaceholderText(
                "Результаты появятся рядом с исходниками: *_оформлено.docx"
            )
            self.locality_edit.clear()
            self.locality_edit.setPlaceholderText("Для пакета берётся из имени каждой папки НП")
            self.run_button.setText("Оформить все найденные описания")
            self.batch_preview_card.setVisible(True)
            self._refresh_batch_documents()
            self._hide_activity()
            self._update_preview()
            self._refresh_controls()
            return
        if source.suffix.casefold() not in {".doc", ".docx"}:
            QMessageBox.warning(
                self, "Неверный формат", "Выберите документ Word с расширением .doc или .docx."
            )
            return
        self._batch_documents = ()
        self.batch_table.setRowCount(0)
        self.batch_preview_card.setVisible(False)
        self.run_button.setText("Оформить графическое описание")
        self.target_edit.setPlaceholderText("Отдельный файл результата")
        self.locality_edit.setPlaceholderText("Будет извлечён из исходного Word")
        self.target_edit.setText(
            str(source.with_name(f"{source.stem}_оформлено{source.suffix}"))
        )
        self._start_inspection(source)

    def _inspect_typed_source(self):
        value = self.source_edit.text().strip()
        if value and Path(value).exists():
            self._source_selected(value)

    def _refresh_batch_documents(self):
        root = Path(self.source_edit.text().strip())
        if not root.is_dir():
            return
        self._batch_documents = find_batch_documents(root)
        self._populate_batch_table(root)
        self._batch_selection_changed()

    def _populate_batch_table(self, root: Path):
        blocker = QSignalBlocker(self.batch_table)
        self.batch_table.setRowCount(len(self._batch_documents))
        for row, document in enumerate(self._batch_documents):
            check_item = QTableWidgetItem()
            check_item.setFlags(
                Qt.ItemFlag.ItemIsEnabled
                | Qt.ItemFlag.ItemIsSelectable
                | Qt.ItemFlag.ItemIsUserCheckable
            )
            check_item.setCheckState(Qt.CheckState.Checked)
            check_item.setData(Qt.ItemDataRole.UserRole, row)

            try:
                source_text = str(document.source.relative_to(root))
                target_text = str(document.target.relative_to(root))
            except ValueError:
                source_text = str(document.source)
                target_text = str(document.target)

            zone_name = document.source.parent.name
            if document.locality:
                document_type = (
                    f"ТЗ {zone_name}"
                    if is_zone_folder_name(zone_name)
                    else "НП"
                )
            else:
                document_type = (
                    f"Вне НП · ТЗ {zone_name}"
                    if is_zone_folder_name(zone_name)
                    else "Вне НП"
                )

            address = build_full_address(
                self.region_edit.text(),
                self.district_combo.currentText(),
                self.settlement_combo.currentText(),
                document.locality,
            )
            values = (source_text, document_type, address, target_text)
            self.batch_table.setItem(row, 0, check_item)
            for column, value in enumerate(values, start=1):
                item = QTableWidgetItem(value)
                if column in {1, 4}:
                    item.setToolTip(str(document.source if column == 1 else document.target))
                self.batch_table.setItem(row, column, item)
        del blocker

    def _update_batch_addresses(self):
        if not self._is_batch_mode():
            return
        blocker = QSignalBlocker(self.batch_table)
        for row, document in enumerate(self._batch_documents):
            item = self.batch_table.item(row, 3)
            if item is not None:
                item.setText(
                    build_full_address(
                        self.region_edit.text(),
                        self.district_combo.currentText(),
                        self.settlement_combo.currentText(),
                        document.locality,
                    )
                )
        del blocker

    def _selected_batch_documents(self):
        selected = []
        for row, document in enumerate(self._batch_documents):
            item = self.batch_table.item(row, 0)
            if item is not None and item.checkState() == Qt.CheckState.Checked:
                selected.append(document)
        return tuple(selected)

    def _set_all_batch_checked(self, checked: bool):
        state = Qt.CheckState.Checked if checked else Qt.CheckState.Unchecked
        blocker = QSignalBlocker(self.batch_table)
        for row in range(self.batch_table.rowCount()):
            item = self.batch_table.item(row, 0)
            if item is not None:
                item.setCheckState(state)
        del blocker
        self._batch_selection_changed()

    def _batch_selection_changed(self, *_args):
        if not self._is_batch_mode():
            return
        selected = len(self._selected_batch_documents())
        total = len(self._batch_documents)
        self.status_label.setText(
            f"Найдено: {total}. Выбрано для обработки: {selected}. "
            "Служебные каталоги пропущены."
        )
        self.run_button.setText(f"Оформить выбранные описания ({selected})")
        self._refresh_controls()

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

    def _is_batch_mode(self) -> bool:
        value = self.source_edit.text().strip()
        return bool(value) and Path(value).is_dir()

    def _update_preview(self, *_args):
        parts = (
            self.region_edit.text(),
            self.district_combo.currentText(),
            self.settlement_combo.currentText(),
        )
        if not self._is_batch_mode():
            parts = (*parts, self.locality_edit.text())
        preview = ", ".join(
            part.strip().rstrip(",") for part in parts if part.strip()
        )
        if self._is_batch_mode():
            batch_hint = "НП — из структуры папок; для «Вне НП» не добавляется"
            preview = f"{preview}, {batch_hint}" if preview else batch_hint
        self.preview_label.setText(preview or "Адрес пока не заполнен")
        self._update_batch_addresses()

    def _run(self):
        if self._formatting or self._inspecting:
            return
        if self._is_batch_mode():
            self._run_batch()
            return
        values = (
            self.source_edit.text().strip(),
            self.target_edit.text().strip(),
            self.region_edit.text().strip(),
            self.district_combo.currentText().strip(),
            self.settlement_combo.currentText().strip(),
            self.locality_edit.text().strip() or None,
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

    def _run_batch(self):
        if not self._batch_documents:
            QMessageBox.warning(
                self,
                "Документы не найдены",
                "В выбранной структуре нет Word-файлов непосредственно в папках НП "
                "или в их дочерних папках с индексами ТЗ.",
            )
            return
        documents = self._selected_batch_documents()
        if not documents:
            QMessageBox.warning(
                self,
                "Ничего не выбрано",
                "Отметьте в предпросмотре хотя бы один Word-файл.",
            )
            return

        region = self.region_edit.text().strip()
        district = self.district_combo.currentText().strip()
        settlement = self.settlement_combo.currentText().strip()
        if not region or not district:
            QMessageBox.warning(
                self,
                "Заполните адрес",
                "Укажите регион и муниципальный район или округ.",
            )
            return
        if not settlement and "округ" not in district.casefold():
            QMessageBox.warning(
                self,
                "Заполните адрес",
                "Для муниципального района укажите поселение.",
            )
            return

        existing = sum(document.target.exists() for document in documents)
        if existing:
            answer = QMessageBox.question(
                self,
                "Заменить готовые файлы?",
                f"Уже существует результатов: {existing}.\n\n"
                "Заменить их после успешной обработки соответствующих исходников?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.No,
            )
            if answer != QMessageBox.StandardButton.Yes:
                return

        self._formatting = True
        self.status_label.setText(
            f"Пакетная обработка Word: 0 из {len(documents)}…"
        )
        self.progress_bar.setVisible(True)
        self.update_progress(0, len(documents))
        self._refresh_controls()
        self.start_task(
            self._format_documents_task,
            documents,
            region,
            district,
            settlement,
            on_result=self._batch_format_done,
            on_error=self._format_failed,
            on_finished=self._format_finished,
        )

    @staticmethod
    def _format_documents_task(signals, documents, region, district, settlement):
        return format_description_documents(
            documents,
            region,
            district,
            settlement,
            lambda current, total: signals.progress.emit(current, total),
        )

    def _batch_format_done(self, result: BatchFormatResult):
        success_count = len(result.completed)
        failure_count = len(result.failures)
        self.status_label.setText(
            f"Пакет завершён: готово {success_count}, ошибок {failure_count}."
        )
        details = ""
        if result.failures:
            lines = [
                f"• {failure.source}: {failure.message}"
                for failure in result.failures[:8]
            ]
            if len(result.failures) > len(lines):
                lines.append(f"• …ещё ошибок: {len(result.failures) - len(lines)}")
            details = "\n\nОшибки:\n" + "\n".join(lines)
        QMessageBox.information(
            self,
            "Пакетная обработка завершена",
            f"Найдено документов: {len(result.documents)}\n"
            f"Успешно оформлено: {success_count}\n"
            f"С ошибками: {failure_count}"
            f"{details}",
        )

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
        busy = self._formatting or self._inspecting
        batch_mode = self._is_batch_mode()
        self.source_button.setEnabled(not busy)
        self.source_folder_button.setEnabled(not busy)
        self.target_button.setEnabled(not busy and not batch_mode)
        self.source_edit.setEnabled(not busy)
        self.target_edit.setEnabled(not busy and not batch_mode)
        self.region_edit.setEnabled(not busy)
        self.district_combo.setEnabled(not busy)
        self.settlement_combo.setEnabled(not busy)
        self.locality_edit.setEnabled(not busy and not batch_mode)
        self.batch_table.setEnabled(not busy)
        self.batch_select_all_button.setEnabled(not busy)
        self.batch_clear_button.setEnabled(not busy)
        self.batch_refresh_button.setEnabled(not busy)
        has_batch_selection = not batch_mode or bool(self._selected_batch_documents())
        self.run_button.setEnabled(
            not busy and not self._inspecting and has_batch_selection
        )

    @staticmethod
    def _error_message(traceback_text: str) -> str:
        lines = [line.strip() for line in str(traceback_text).splitlines() if line.strip()]
        if not lines:
            return "Неизвестная ошибка Microsoft Word."
        last = lines[-1]
        if ": " in last:
            return last.split(": ", 1)[1]
        return last
