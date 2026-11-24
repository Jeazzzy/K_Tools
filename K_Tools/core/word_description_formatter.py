"""Оформление выгруженных графических описаний через Microsoft Word COM."""

from __future__ import annotations

import csv
import gc
import os
import re
import shutil
import tempfile
import time
from dataclasses import dataclass
from pathlib import Path
from typing import Callable

try:
    import pythoncom
    import win32com.client
except ImportError as error:  # Приложение должно запускаться и без Word/pywin32.
    pythoncom = None
    win32com = None
    _WORD_IMPORT_ERROR = error
else:
    _WORD_IMPORT_ERROR = None


WD_ACTIVE_END_PAGE_NUMBER = 3
WD_ALIGN_PARAGRAPH_CENTER = 1

LOCALITY_TYPES = {
    "д": "деревня",
    "с": "село",
    "п": "посёлок",
    "пос": "посёлок",
    "г": "город",
    "дп": "дачный посёлок",
    "гп": "городской посёлок",
}
LOCALITY_PATTERN = re.compile(
    r"\b(деревня|село|пос[её]лок|город|хутор|станция|дп|гп|пос|д|с|п|г)\.?\s+"
    r"(.+?)(?=\s+муниципального\s+образования|\s+[А-ЯЁ][а-яё]+\s+области|\s*\(|$)",
    re.IGNORECASE,
)
FOLDER_LOCALITY_PATTERN = re.compile(
    r"^(?P<kind>городской\s+пос[её]лок|дачный\s+пос[её]лок|рабочий\s+пос[её]лок|"
    r"пос[её]лок\s+городского\s+типа|деревня|село|пос[её]лок|город|хутор|станция|"
    r"пгт|рп|дп|гп|пос|д|с|п|г)\.?\s+(?P<name>.+)$",
    re.IGNORECASE,
)
ZONE_FOLDER_PATTERN = re.compile(
    r"^[А-ЯЁ]{1,8}[-_. ]*\d+(?:[.-]\d+)*(?:[_ -].*)?$",
    re.IGNORECASE,
)
FOLDER_LOCALITY_TYPES = {
    "деревня": "деревня",
    "село": "село",
    "поселок": "посёлок",
    "посёлок": "посёлок",
    "город": "город",
    "хутор": "хутор",
    "станция": "станция",
    "городской поселок": "городской посёлок",
    "городской посёлок": "городской посёлок",
    "дачный поселок": "дачный посёлок",
    "дачный посёлок": "дачный посёлок",
    "рабочий поселок": "рабочий посёлок",
    "рабочий посёлок": "рабочий посёлок",
    "поселок городского типа": "посёлок городского типа",
    "посёлок городского типа": "посёлок городского типа",
    "пгт": "посёлок городского типа",
    "рп": "рабочий посёлок",
    **LOCALITY_TYPES,
}
DECIMAL_PATTERN = re.compile(r"(?<=\d)\.(?=\d)")
PART_PATTERN = re.compile(r"\bчасть\s*(?:№|N|No\.?)\s*\d+", re.IGNORECASE)


@dataclass(frozen=True)
class DocumentInfo:
    address: str
    locality: str
    coordinate_tables: int


@dataclass(frozen=True)
class BatchDocument:
    source: Path
    target: Path
    locality: str


@dataclass(frozen=True)
class BatchFailure:
    source: Path
    message: str


@dataclass(frozen=True)
class BatchFormatResult:
    documents: tuple[BatchDocument, ...]
    completed: tuple[DocumentInfo, ...]
    failures: tuple[BatchFailure, ...]


def addresses_file() -> Path:
    """Возвращает путь к упакованному справочнику районов и поселений."""

    return Path(__file__).resolve().parents[1] / "resources" / "addresses_lo.csv"


def load_lo_addresses(path: str | Path | None = None) -> dict[str, list[str]]:
    """Загружает подсказки Ленинградской области из CSV."""

    result: dict[str, set[str]] = {}
    address_path = Path(path) if path is not None else addresses_file()
    with address_path.open(encoding="utf-8-sig", newline="") as stream:
        for row in csv.DictReader(stream):
            district = (row.get("район") or "").strip()
            settlement = (row.get("поселение") or "").strip()
            if not district:
                continue
            result.setdefault(district, set())
            if settlement:
                result[district].add(settlement)
    return {
        district: sorted(settlements, key=str.casefold)
        for district, settlements in sorted(result.items(), key=lambda item: item[0].casefold())
    }


def clean_cell(cell) -> str:
    return cell.Range.Text.rstrip("\r\x07").strip()


def replace_cell(cell, value: str) -> None:
    # Конечный маркер ячейки остаётся на месте, поэтому сетка Word сохраняется.
    cell.Range.Text = value


def normalize_locality(value: str) -> str:
    value = re.sub(r"\s+", " ", value.strip()).rstrip(".,")
    match = re.match(r"^(дп|гп|пос|д|с|п|г)\.?\s+(.+)$", value, re.IGNORECASE)
    if match:
        return f"{LOCALITY_TYPES[match.group(1).lower()]} {match.group(2)}"
    return value


def extract_locality(address: str, title: str) -> str:
    for text in ((address or "").split(",")[-1].strip(), title or ""):
        match = LOCALITY_PATTERN.search(text)
        if match:
            return normalize_locality(f"{match.group(1)} {match.group(2).strip()}")
    return ""


def locality_from_folder_name(name: str) -> str:
    """Возвращает НП только для каталога с явным типом населённого пункта."""

    value = re.sub(r"^[\s_!–—-]+", "", str(name))
    value = re.sub(r"\s+", " ", value.strip()).rstrip(".,")
    match = FOLDER_LOCALITY_PATTERN.fullmatch(value)
    if not match:
        return ""
    kind = FOLDER_LOCALITY_TYPES[match.group("kind").casefold()]
    locality_name = match.group("name").strip().rstrip(".,")
    return f"{kind} {locality_name}" if locality_name else ""


def is_outside_locality_folder(name: str) -> bool:
    """Распознаёт каталог «Вне НП» с техническими префиксами."""

    value = re.sub(r"^[\s_!–—-]+", "", str(name))
    value = re.sub(r"[._\s-]+", " ", value.strip()).casefold()
    return value == "вне нп"


def is_zone_folder_name(name: str) -> bool:
    """Распознаёт индекс ТЗ и не принимает служебные каталоги вроде РН/XML."""

    value = re.sub(r"^[\s_!–—-]+", "", str(name)).strip()
    return bool(ZONE_FOLDER_PATTERN.fullmatch(value))


def find_batch_documents(root: str | Path) -> tuple[BatchDocument, ...]:
    """Находит Word в каталогах НП и в их непосредственных каталогах ТЗ.

    Каталоги без названия НП или индекса зоны не обрабатываются, но рекурсивный
    обход через них продолжается. Поэтому «Архив» безопасно пропускается, а
    находящаяся внутри него самостоятельная папка «деревня Пример» будет найдена.
    """

    root = Path(root)
    if not root.is_dir():
        raise ValueError("Выберите существующую корневую папку.")

    documents: list[BatchDocument] = []
    for directory_name, subdirectories, file_names in os.walk(root):
        subdirectories.sort(key=str.casefold)
        file_names.sort(key=str.casefold)
        directory = Path(directory_name)
        locality = locality_from_folder_name(directory.name)
        eligible = bool(locality) or is_outside_locality_folder(directory.name)
        if not eligible and is_zone_folder_name(directory.name):
            parent = directory.parent.name
            locality = locality_from_folder_name(parent)
            eligible = bool(locality) or is_outside_locality_folder(parent)
        if not eligible:
            continue
        for file_name in file_names:
            source = directory / file_name
            if source.suffix.casefold() not in {".doc", ".docx"}:
                continue
            if source.name.startswith("~$") or source.stem.casefold().endswith("_оформлено"):
                continue
            target = source.with_name(f"{source.stem}_оформлено{source.suffix}")
            documents.append(BatchDocument(source, target, locality))
    return tuple(documents)


def build_full_address(region: str, district: str, settlement: str, locality: str) -> str:
    return ", ".join(
        part.strip().rstrip(",")
        for part in (region, district, settlement, locality)
        if part.strip()
    )


def resolve_locality(
    explicit_locality: str | None,
    current_address: str,
    title: str,
) -> str:
    """Отличает автоопределение НП от явно выбранного режима «Вне НП»."""

    if explicit_locality is None:
        return extract_locality(current_address, title)
    return normalize_locality(explicit_locality)


def _require_word_automation() -> None:
    if pythoncom is None or win32com is None:
        message = (
            "Недоступен компонент автоматизации Microsoft Word. Функция «Описание ЛО» "
            "работает только в Windows с установленным Microsoft Word. Переустановите "
            "K Tools с компонентом pywin32 или обратитесь к администратору."
        )
        if _WORD_IMPORT_ERROR is not None:
            message += f"\n\nТехническая причина: {_WORD_IMPORT_ERROR}"
        raise RuntimeError(message)


def word_automation_available() -> bool:
    """Показывает, установлен ли Python-компонент COM (наличие Word проверяется при запуске)."""

    return pythoncom is not None and win32com is not None


def _open_word():
    _require_word_automation()
    try:
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0
        word.AutomationSecurity = 3
        return word
    except Exception as error:
        raise RuntimeError(
            "Не удалось запустить Microsoft Word через COM. Убедитесь, что Word "
            "установлен и запускается вручную; при повреждённой регистрации Office "
            "выполните восстановление Microsoft 365/Office."
        ) from error


def _open_document(word, path: Path, *, read_only: bool):
    return word.Documents.Open(
        FileName=str(path.resolve()),
        ConfirmConversions=False,
        ReadOnly=read_only,
        AddToRecentFiles=False,
        NoEncodingDialog=True,
    )


def _close_document(document, *, save: bool = False) -> None:
    if document is not None:
        document.Close(-1 if save else 0)


def _finish_word(word) -> None:
    if word is not None:
        word.Quit()


def _is_coordinate_table(table) -> bool:
    if table.Columns.Count not in (6, 8):
        return False
    prefix = table.Range.Text[:1200]
    return "координат" in prefix.lower() or bool(re.search(r"\d{5,}[.,]\d{2}", prefix))


def _validate_source(path: Path) -> None:
    if not path.is_file():
        raise ValueError("Выберите существующий исходный файл Word.")
    if path.suffix.lower() not in {".doc", ".docx"}:
        raise ValueError("Выберите файл Word с расширением .doc или .docx.")


def inspect_document(path: str | Path) -> DocumentInfo:
    path = Path(path)
    _validate_source(path)
    _require_word_automation()
    pythoncom.CoInitialize()
    word = None
    document = None
    primary_error: BaseException | None = None
    try:
        word = _open_word()
        document = _open_document(word, path, read_only=True)
        if document.Tables.Count < 1 or document.Paragraphs.Count < 3:
            raise ValueError("Структура Word не похожа на графическое описание границ.")
        address = clean_cell(document.Tables(1).Cell(4, 3))
        title = document.Paragraphs(3).Range.Text.strip()
        count = sum(
            _is_coordinate_table(document.Tables(index))
            for index in range(1, document.Tables.Count + 1)
        )
        return DocumentInfo(address, extract_locality(address, title), count)
    except BaseException as error:
        primary_error = error
        raise
    finally:
        try:
            _close_document(document)
        except Exception:
            if primary_error is None:
                raise
        finally:
            document = None
            try:
                _finish_word(word)
            except Exception:
                if primary_error is None:
                    raise
            finally:
                word = None
                gc.collect()
                pythoncom.CoUninitialize()


def _format_coordinate_table(table) -> None:
    column_count = table.Columns.Count
    rows: dict[int, list[tuple]] = {}
    for cell in table.Range.Cells:
        rows.setdefault(cell.RowIndex, []).append((cell, clean_cell(cell)))

    numbering_rows = {
        number
        for number, row_cells in rows.items()
        if [value for _, value in row_cells]
        == [str(index) for index in range(1, column_count + 1)]
    }
    first_numbering_row = min(numbering_rows, default=column_count + 1)

    for row_number, row_cells in rows.items():
        values = [value for _, value in row_cells]
        if len(row_cells) == 1:
            if PART_PATTERN.search(values[0]) or "Внутренний контур" in values[0]:
                row_cells[0][0].Range.ParagraphFormat.Alignment = WD_ALIGN_PARAGRAPH_CENTER
            continue

        if row_number <= first_numbering_row or row_number in numbering_rows:
            continue
        if len(row_cells) != column_count:
            continue

        if PART_PATTERN.search(" ".join(values)):
            label = next(value for value in values if PART_PATTERN.search(value))
            row_cells[0][0].Merge(row_cells[-1][0])
            replace_cell(table.Cell(row_number, 1), label)
            table.Cell(row_number, 1).Range.ParagraphFormat.Alignment = WD_ALIGN_PARAGRAPH_CENTER
            continue

        if all(not value for value in values):
            row_cells[0][0].Merge(row_cells[-1][0])
            replace_cell(table.Cell(row_number, 1), "Внутренний контур")
            table.Cell(row_number, 1).Range.ParagraphFormat.Alignment = WD_ALIGN_PARAGRAPH_CENTER
            continue

        if values[0] in {"", "-"}:
            continue

        numeric_columns = (2, 3, 4, 5, 7) if column_count == 8 else (2, 3, 5)
        for column in numeric_columns:
            old = values[column - 1]
            new = DECIMAL_PATTERN.sub(",", old)
            if new != old:
                replace_cell(row_cells[column - 1][0], new)

        if values[column_count - 1] == "-":
            replace_cell(row_cells[-1][0], "Закрепление отсутствует")


def _is_numbering_row(table, row: int) -> bool:
    try:
        return [clean_cell(table.Cell(row, column)) for column in range(1, 9)] == [
            str(column) for column in range(1, 9)
        ]
    except Exception:
        return False


def _disable_all_heading_rows(word, document) -> None:
    for table in document.Tables:
        table.Rows.HeadingFormat = False
        if table.Rows.HeadingFormat == 0:
            continue
        # Word может оставить верхнюю строку отмеченной при объединённых ячейках.
        for row in range(1, min(table.Rows.Count, 8) + 1):
            if table.Rows.HeadingFormat == 0:
                break
            try:
                table.Cell(row, 1).Range.Select()
                word.Selection.SelectRow()
                word.Selection.Rows.HeadingFormat = False
            except Exception:
                continue
        if table.Rows.HeadingFormat != 0:
            raise RuntimeError("Не удалось отключить повтор заголовков таблицы.")


def _insert_number_rows_on_new_pages(word, document, table) -> None:
    """Добавляет физические строки 1–8 в начале страниц продолжения."""

    if table.Columns.Count != 8 or table.Rows.Count < 9:
        return
    number_row_text = "".join(f"{column}\r\x07" for column in range(1, 9))
    if table.Range.Text.count(number_row_text) > 1:
        return

    table.Rows.AllowBreakAcrossPages = False
    for _ in range(100):
        document.Repaginate()
        first_numbering_row = next(
            (
                row
                for row in range(1, min(table.Rows.Count, 8) + 1)
                if _is_numbering_row(table, row)
            ),
            None,
        )
        if first_numbering_row is None:
            return
        previous_page = table.Cell(first_numbering_row, 1).Range.Information(
            WD_ACTIVE_END_PAGE_NUMBER
        )
        added = False
        for row in range(first_numbering_row + 1, table.Rows.Count + 1):
            page = table.Cell(row, 1).Range.Information(WD_ACTIVE_END_PAGE_NUMBER)
            if page > previous_page and not _is_numbering_row(table, row):
                table.Cell(row, 1).Range.Select()
                word.Selection.InsertRowsAbove(1)
                for column in range(1, 9):
                    cell = table.Cell(row, column)
                    replace_cell(cell, str(column))
                    cell.Range.ParagraphFormat.Alignment = WD_ALIGN_PARAGRAPH_CENTER
                table.Cell(row, 1).Range.ParagraphFormat.PageBreakBefore = True
                added = True
                break
            previous_page = page
        if not added:
            return
    raise RuntimeError("Не удалось расставить номера столбцов на всех страницах.")


def _validate_format_request(
    source: Path,
    target: Path,
    region: str,
    district: str,
    settlement: str,
) -> None:
    _validate_source(source)
    if source.resolve() == target.resolve():
        raise ValueError("Исходный и выходной файлы должны отличаться.")
    if target.suffix.lower() != source.suffix.lower():
        raise ValueError("Расширение выходного файла должно совпадать с исходным.")
    if not region.strip() or not district.strip():
        raise ValueError("Укажите регион и муниципальный район или округ.")
    if not settlement.strip() and "округ" not in district.casefold():
        raise ValueError("Укажите поселение.")


def _format_document_in_word(
    word,
    source: str | Path,
    target: str | Path,
    region: str,
    district: str,
    settlement: str,
    locality: str | None = None,
    message_callback: Callable[[str], None] | None = None,
) -> DocumentInfo:
    """Оформляет один документ в уже запущенном экземпляре Microsoft Word."""

    source, target = Path(source), Path(target)
    _validate_format_request(source, target, region, district, settlement)
    target.parent.mkdir(parents=True, exist_ok=True)

    descriptor, temporary_name = tempfile.mkstemp(
        prefix=".ktools-description-",
        suffix=source.suffix,
        dir=target.parent,
    )
    os.close(descriptor)
    temporary = Path(temporary_name)
    try:
        shutil.copy2(source, temporary)
    except Exception:
        temporary.unlink(missing_ok=True)
        raise

    document = None
    info: DocumentInfo | None = None
    primary_error: BaseException | None = None
    try:
        if message_callback is not None:
            message_callback(f"Открытие Word: {source}")
        document = _open_document(word, temporary, read_only=False)
        if document.ReadOnly:
            raise RuntimeError("Word открыл временную копию только для чтения.")
        if document.Tables.Count < 1 or document.Paragraphs.Count < 3:
            raise ValueError("Структура Word не похожа на графическое описание границ.")

        current_address = clean_cell(document.Tables(1).Cell(4, 3))
        title = document.Paragraphs(3).Range.Text.strip()
        place = resolve_locality(locality, current_address, title)
        if locality is None and not place:
            raise ValueError("Не удалось найти населённый пункт. Укажите его вручную.")

        full_address = build_full_address(region, district, settlement, place)
        replace_cell(document.Tables(1).Cell(4, 3), full_address)
        if place:
            document.Paragraphs(3).Range.Text = place + "\r"

        coordinate_tables = []
        for index in range(1, document.Tables.Count + 1):
            table = document.Tables(index)
            table.Rows.HeadingFormat = False
            if _is_coordinate_table(table):
                _format_coordinate_table(table)
                coordinate_tables.append(table)
        if message_callback is not None:
            message_callback(
                f"Координатных таблиц: {len(coordinate_tables)}; подготовка пагинации…"
            )
        for table_number, table in enumerate(coordinate_tables, start=1):
            if table.Columns.Count == 8:
                if message_callback is not None:
                    message_callback(
                        f"Перепагинация таблицы {table_number} из {len(coordinate_tables)}…"
                    )
                _insert_number_rows_on_new_pages(word, document, table)
        _disable_all_heading_rows(word, document)

        if message_callback is not None:
            message_callback("Сохранение результата…")
        document.Save()
        info = DocumentInfo(full_address, place, len(coordinate_tables))
    except BaseException as error:
        primary_error = error
    finally:
        try:
            _close_document(document)
        except Exception:
            if primary_error is None:
                primary_error = RuntimeError("Word не смог закрыть обработанный документ.")
        finally:
            document = None
            gc.collect()

        if primary_error is not None and info is not None:
            info = None

    if info is None:
        temporary.unlink(missing_ok=True)
        if primary_error is not None:
            raise primary_error
        raise RuntimeError("Microsoft Word не создал результат.")

    try:
        os.replace(temporary, target)
    except Exception:
        temporary.unlink(missing_ok=True)
        raise
    return info


def format_document(
    source: str | Path,
    target: str | Path,
    region: str,
    district: str,
    settlement: str,
    locality: str | None = None,
) -> DocumentInfo:
    """Создаёт оформленную копию, не изменяя исходник и старый результат при сбое."""

    _validate_format_request(Path(source), Path(target), region, district, settlement)
    _require_word_automation()
    pythoncom.CoInitialize()
    word = None
    primary_error: BaseException | None = None
    try:
        word = _open_word()
        return _format_document_in_word(
            word,
            source,
            target,
            region,
            district,
            settlement,
            locality,
        )
    except BaseException as error:
        primary_error = error
        raise
    finally:
        try:
            _finish_word(word)
        except Exception:
            if primary_error is None:
                raise RuntimeError("Word не смог завершить COM-сеанс.")
        finally:
            word = None
            gc.collect()
            pythoncom.CoUninitialize()


def format_description_documents(
    documents: tuple[BatchDocument, ...] | list[BatchDocument],
    region: str,
    district: str,
    settlement: str,
    progress_callback: Callable[[int, int], None] | None = None,
    message_callback: Callable[[str], None] | None = None,
) -> BatchFormatResult:
    """Оформляет выбранные документы в одном общем сеансе Microsoft Word."""

    documents = tuple(documents)
    completed: list[DocumentInfo] = []
    failures: list[BatchFailure] = []
    total = len(documents)
    if progress_callback is not None:
        progress_callback(0, max(total, 1))

    if not documents:
        return BatchFormatResult(documents, (), ())

    _require_word_automation()
    pythoncom.CoInitialize()
    word = None
    batch_started = time.monotonic()
    try:
        if message_callback is not None:
            message_callback(f"Запуск Microsoft Word. Документов в очереди: {total}.")
        word = _open_word()
        for position, document in enumerate(documents, start=1):
            document_started = time.monotonic()
            if message_callback is not None:
                message_callback(f"[{position}/{total}] Начало: {document.source}")
            failed = False
            stage_callback = None
            if message_callback is not None:
                stage_callback = lambda text, current=position: message_callback(
                    f"[{current}/{total}] {text}"
                )
            try:
                completed.append(
                    _format_document_in_word(
                        word,
                        document.source,
                        document.target,
                        region,
                        district,
                        settlement,
                        document.locality,
                        stage_callback,
                    )
                )
            except Exception as error:
                failed = True
                failures.append(BatchFailure(document.source, str(error)))
                if message_callback is not None:
                    message_callback(
                        f"[{position}/{total}] ОШИБКА: {document.source} — {error}"
                    )

            elapsed = time.monotonic() - document_started
            batch_elapsed = time.monotonic() - batch_started
            remaining = (batch_elapsed / position) * (total - position)
            if message_callback is not None:
                outcome = "с ошибкой" if failed else "готово"
                message_callback(
                    f"[{position}/{total}] {outcome} за {_format_duration(elapsed)}. "
                    f"Примерно осталось: {_format_duration(remaining)}."
                )
            if progress_callback is not None:
                progress_callback(position, max(total, 1))
    finally:
        try:
            _finish_word(word)
        finally:
            word = None
            gc.collect()
            pythoncom.CoUninitialize()

    if message_callback is not None:
        message_callback(
            f"Пакет завершён за {_format_duration(time.monotonic() - batch_started)}: "
            f"готово {len(completed)}, ошибок {len(failures)}."
        )

    return BatchFormatResult(documents, tuple(completed), tuple(failures))


def _format_duration(seconds: float) -> str:
    seconds = max(0, int(round(seconds)))
    hours, remainder = divmod(seconds, 3600)
    minutes, seconds = divmod(remainder, 60)
    if hours:
        return f"{hours} ч {minutes} мин"
    if minutes:
        return f"{minutes} мин {seconds} с"
    return f"{seconds} с"


def format_description_folder(
    root: str | Path,
    region: str,
    district: str,
    settlement: str,
    progress_callback: Callable[[int, int], None] | None = None,
    message_callback: Callable[[str], None] | None = None,
) -> BatchFormatResult:
    """Последовательно оформляет все документы, найденные в корневой папке."""

    return format_description_documents(
        find_batch_documents(root),
        region,
        district,
        settlement,
        progress_callback,
        message_callback,
    )
