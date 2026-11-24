from PySide6.QtCore import Qt
from PySide6.QtWidgets import QHBoxLayout, QLabel, QPushButton

from core import BasePage


TOOLS = [
    {
        "page": 1,
        "name": "XML → CSV",
        "summary": "Собирает URL-адреса из протокольных XML-файлов в один CSV.",
        "input": "Папка, в которой находятся файлы с именами proto_*.xml. Поиск выполняется и во вложенных папках.",
        "workflow": "Выберите исходную папку, нажмите «Обработать в CSV» и укажите имя итогового файла.",
        "result": "CSV в кодировке UTF-8 со всеми найденными значениями Stage/URL и статистика обработки.",
    },
    {
        "page": 2,
        "name": "Распаковка ZIP",
        "summary": "Пакетно разбирает архивы ЕГРН и именует файлы по кадастровому номеру.",
        "input": "Исходная папка с ZIP-архивами и отдельная папка для результата.",
        "workflow": "Программа читает первый XML архива, находит кадастровый номер и извлекает содержимое.",
        "result": "Папки ZIP, XML и PDF с переименованными файлами. Отдельная зона позволяет быстро переименовывать ZIP/XML на месте.",
    },
    {
        "page": 3,
        "name": "Исправление MIF",
        "summary": "Заменяет строку CoordSys в выбранных MIF-файлах MapInfo.",
        "input": "Один или несколько MIF-файлов, перетащенных из Проводника.",
        "workflow": "Добавьте файлы, проверьте их количество и нажмите «Исправить пределы».",
        "result": "Исходные MIF обновляются на месте. Перед массовой обработкой рекомендуется иметь резервную копию.",
    },
    {
        "page": 4,
        "name": "Описание ЛО",
        "summary": "Оформляет выгруженное графическое описание границ в отдельном документе Word.",
        "input": "Исходный .doc/.docx, отдельный файл результата, регион, район или округ, поселение и населённый пункт. Для Ленинградской области доступны встроенные подсказки.",
        "workflow": "Программа извлекает населённый пункт, формирует полный адрес, приводит координаты и погрешности к десятичной запятой, оформляет внутренние контуры и строки продолжения таблиц.",
        "result": "Новый Word-файл того же формата. Исходный документ не изменяется; при ошибке временный результат удаляется. Требуются Windows и установленный Microsoft Word.",
    },
    {
        "page": 5,
        "name": "Работа с MDB",
        "summary": "Выполняет пять массовых операций с базами Microsoft Access.",
        "input": "В зависимости от режима: source MDB, папки индексов, target-папка и имя таблицы.",
        "workflow": "Доступны полная замена MDB, перенос Utilizations_KP, перенос выбранной таблицы, безопасный перенос адреса Locations и обновление текста Объект_ЗУ.",
        "result": "Изменённые MDB и подробный журнал по каждому обработанному файлу. Требуется Microsoft Access Database Engine.",
    },
    {
        "page": 6,
        "name": "Разделение ТЗ",
        "summary": "Распределяет территориальные зоны по населённым пунктам на основе геометрического пересечения.",
        "input": "TAB с границами НП, TAB с территориальными зонами, поле названия НП и папка результата.",
        "workflow": "Настройте пороги «в НП» и «вне НП». Значения между порогами будут отмечены для ручной проверки.",
        "result": "Отдельные TAB по населённым пунктам, ВНЕ_НП.tab и СОМНИТЕЛЬНЫЕ.tab при наличии соответствующих объектов.",
    },
    {
        "page": 7,
        "name": "Проверка выпуска",
        "summary": "Проверяет площади в PDF и сведения, геометрию и точность координат в выпускных XML.",
        "input": "Папка всего выпуска с вложенными папками pdf и xml либо одна из этих папок. Выбирается режим территориальных зон или населённых пунктов и допустимая точность XML. Если объект пересекает планшеты 1:2000 и 1:10000, включается отдельный флажок.",
        "workflow": "Во вкладке PDF извлекаются площадь и количество страниц. Во вкладке XML программа читает XML внутри ZIP, считает все контуры и точки без замыкающих повторов, а по координатам отдельно определяет внешние контуры и внутренние кольца (дырки). Если координат недостаточно, классификация показывается прочерком без ложного вычитания. Также проверяются delta_geopoint/geopoint_opred. Для ТЗ индекс сверяется с папкой зоны. Для НП проверяются папка населённого пункта, type_boundary и reg_numb_border. Режим смешанных планшетов разрешает совместную картометрическую погрешность 1 и 5 м. В меню кнопки «Создать оглавление» доступны варианты с XML, без XML и по общим PDF ТЗ для населённых пунктов. K Tools сам создаёт стандартный титульник: достаточно проверить муниципальное образование, название документа и необязательный номер тома. Подменю «По своему DOCX» оставлено для нестандартного оформления.",
        "result": "Отдельные таблицы результатов PDF и XML. Ошибки выделяются цветом и расшифровываются во всплывающей подсказке. Результаты можно экспортировать в XLSX. Кнопка «Создать оглавление» формирует готовый титульный лист и DOCX с вычисленными начальными страницами PDF и названиями объектов из XML либо самих PDF. Если Word COM недоступен, файл сохраняется с обязательным предупреждением о ручной проверке страниц.",
    },
]


class HelpPage(BasePage):
    def __init__(self, controller=None, parent=None):
        super().__init__(controller, parent)
        root = self.page_layout(
            "Справка по K Tools",
            "Назначение инструментов, необходимые данные и результат обработки.",
        )
        root.setAlignment(Qt.AlignmentFlag.AlignTop)

        for tool in TOOLS:
            self._add_tool_card(root, tool)

        _, notes = self.card_layout(root, "Перед началом работы")
        notes_text = QLabel(
            "Для операций с MDB должен быть установлен Microsoft Access Database Engine "
            "подходящей разрядности. Для операций с TAB используются GeoPandas, GDAL и "
            "драйвер MapInfo. Массовые операции изменяют или создают файлы, поэтому для "
            "важных исходных данных рекомендуется заранее сделать резервную копию. "
            "Для инструмента «Описание ЛО» необходим установленный Microsoft Word."
        )
        notes_text.setWordWrap(True)
        notes.addWidget(notes_text)
        root.addStretch()

    def _add_tool_card(self, root, tool):
        _, card = self.card_layout(root)
        header = QHBoxLayout()
        title = QLabel(tool["name"])
        title.setObjectName("toolTitle")
        header.addWidget(title, 1)
        open_button = QPushButton("Открыть инструмент")
        open_button.setObjectName("helpOpenButton")
        open_button.setProperty("primary", True)
        open_button.setEnabled(
            self.controller is not None and hasattr(self.controller, "navigate_to")
        )
        open_button.clicked.connect(
            lambda checked=False, page=tool["page"]: self.controller.navigate_to(page)
        )
        header.addWidget(open_button)
        card.addLayout(header)

        summary = QLabel(tool["summary"])
        summary.setObjectName("toolSummary")
        summary.setWordWrap(True)
        card.addWidget(summary)

        self._add_section(card, "Исходные данные", tool["input"])
        self._add_section(card, "Как работает", tool["workflow"])
        self._add_section(card, "Результат", tool["result"])

    @staticmethod
    def _add_section(layout, heading, text):
        title = QLabel(heading)
        title.setObjectName("helpSectionTitle")
        layout.addWidget(title)
        body = QLabel(text)
        body.setObjectName("helpSectionText")
        body.setWordWrap(True)
        body.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)
        layout.addWidget(body)
