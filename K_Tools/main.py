"""Точка входа K Tools."""

from __future__ import annotations

import os
import sys


# The build script launches the finished windowed EXE with this private flag.
# Keep the branch before Qt and application imports so packaging failures are
# reported by the focused GIS test instead of being hidden by the GUI.
if __name__ == "__main__" and "--build-self-test" in sys.argv:
    from build_self_test import run as run_build_self_test

    raise SystemExit(run_build_self_test(os.environ.get("K_TOOLS_SELF_TEST_REPORT")))

from PySide6.QtCore import QEvent, QTimer, Qt
from PySide6.QtGui import QIcon
from PySide6.QtWidgets import (
    QApplication,
    QButtonGroup,
    QFrame,
    QHBoxLayout,
    QLabel,
    QMainWindow,
    QPushButton,
    QScrollArea,
    QStackedWidget,
    QVBoxLayout,
    QWidget,
)

from pages import (
    HelpPage,
    LoDescriptionPage,
    MdbCopyPage,
    MifProjectionPage,
    ReleaseCheckerPage,
    TzSplitterPage,
    XmlExtractorPage,
    ZipProcessorPage,
)
from theme import Theme, ThemeManager
from version import APP_VERSION, DISPLAY_VERSION


def resource_path(name: str) -> str:
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, name)


class ResponsiveScrollArea(QScrollArea):
    """Подгоняет страницу под окно и прокручивает только реальный избыток."""

    def setWidget(self, widget):
        previous = self.widget()
        if previous is not None:
            previous.removeEventFilter(self)
        super().setWidget(widget)
        if widget is not None:
            widget.installEventFilter(self)
        self._schedule_fit()

    def resizeEvent(self, event):
        super().resizeEvent(event)
        self._fit_widget()

    def eventFilter(self, watched, event):
        if watched is self.widget() and event.type() == QEvent.Type.LayoutRequest:
            self._schedule_fit()
        return super().eventFilter(watched, event)

    def _schedule_fit(self):
        QTimer.singleShot(0, self._fit_widget)

    def _fit_widget(self):
        page = self.widget()
        if page is None:
            return
        viewport = self.viewport().size()
        minimum = page.minimumSizeHint()
        page.resize(
            max(viewport.width(), minimum.width()),
            max(viewport.height(), minimum.height()),
        )


class Application(QMainWindow):
    def __init__(self, theme_manager: ThemeManager):
        super().__init__()
        self.theme_manager = theme_manager
        self.setWindowTitle(
            f"K Tools {DISPLAY_VERSION} — Кадастровые инструменты"
        )
        self.setMinimumSize(760, 560)
        self.resize(1240, 860)
        icon_path = resource_path("icon.ico")
        if os.path.exists(icon_path):
            self.setWindowIcon(QIcon(icon_path))

        root = QWidget()
        root.setObjectName("appRoot")
        self.setCentralWidget(root)
        layout = QHBoxLayout(root)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)

        self.sidebar = QFrame()
        self.sidebar.setObjectName("sidebar")
        self.sidebar.setFixedWidth(235)
        side_layout = QVBoxLayout(self.sidebar)
        side_layout.setContentsMargins(16, 22, 16, 18)
        side_layout.setSpacing(4)
        brand = QLabel("K Tools")
        brand.setObjectName("brand")
        side_layout.addWidget(brand)
        self.brand_subtitle = QLabel("Кадастровые инструменты")
        self.brand_subtitle.setObjectName("brandSub")
        side_layout.addWidget(self.brand_subtitle)
        side_layout.addSpacing(14)

        self.stack = QStackedWidget()
        self.nav_buttons = []
        pages = [
            ("Справка", HelpPage),
            ("XML → CSV", XmlExtractorPage),
            ("Распаковка ZIP", ZipProcessorPage),
            ("Исправление MIF", MifProjectionPage),
            ("Описание ЛО", LoDescriptionPage),
            ("Работа с MDB", MdbCopyPage),
            ("Разделение ТЗ", TzSplitterPage),
            ("Проверка выпуска", ReleaseCheckerPage),
        ]
        buttons = QButtonGroup(self)
        buttons.setExclusive(True)
        for index, (title, page_class) in enumerate(pages):
            button = QPushButton(title)
            button.setObjectName("navButton")
            button.setCheckable(True)
            button.setMinimumHeight(40)
            button.clicked.connect(lambda checked=False, i=index: self.navigate_to(i))
            buttons.addButton(button)
            self.nav_buttons.append(button)
            side_layout.addWidget(button)

            page = page_class(controller=self)
            scroll = ResponsiveScrollArea()
            scroll.setWidgetResizable(False)
            scroll.setFrameShape(QFrame.Shape.NoFrame)
            scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
            scroll.setWidget(page)
            self.stack.addWidget(scroll)

        side_layout.addStretch()
        self.theme_button = QPushButton()
        self.theme_button.setObjectName("themeButton")
        self.theme_button.setMinimumHeight(40)
        self.theme_button.clicked.connect(self.theme_manager.toggle)
        side_layout.addWidget(self.theme_button)
        self.version_label = QLabel(f"v {DISPLAY_VERSION}")
        self.version_label.setObjectName("brandSub")
        side_layout.addWidget(self.version_label)
        self.nav_buttons[0].setChecked(True)
        layout.addWidget(self.sidebar)
        layout.addWidget(self.stack, 1)
        self.theme_manager.theme_changed.connect(self._on_theme_changed)
        self._on_theme_changed(self.theme_manager.theme.value)

    def resizeEvent(self, event):
        """Освобождает больше места инструментам в компактном окне."""
        compact = event.size().width() < 1000
        if event.size().width() < 850:
            sidebar_width = 185
        elif compact:
            sidebar_width = 195
        else:
            sidebar_width = 235
        self.sidebar.setFixedWidth(sidebar_width)
        self.brand_subtitle.setVisible(not compact)
        super().resizeEvent(event)

    def navigate_to(self, index: int):
        """Открывает инструмент и синхронизирует активный пункт навигации."""
        if 0 <= index < self.stack.count():
            self.stack.setCurrentIndex(index)
            self.nav_buttons[index].setChecked(True)

    def _on_theme_changed(self, theme_name: str):
        if Theme(theme_name) == Theme.LIGHT:
            self.theme_button.setText("☾  Тёмная тема")
            current_name = "светлая"
            next_name = "тёмную"
        else:
            self.theme_button.setText("☀  Светлая тема")
            current_name = "тёмная"
            next_name = "светлую"
        self.theme_button.setToolTip(
            f"Сейчас включена {current_name} тема. Переключить на {next_name}.\n"
            "При следующем запуске тема снова будет выбрана по настройке Windows."
        )


def main():
    app = QApplication(sys.argv)
    app.setApplicationName("K Tools")
    app.setApplicationVersion(APP_VERSION)
    app.setStyle("Fusion")
    theme_manager = ThemeManager(app)
    window = Application(theme_manager)
    window.show()
    theme_manager.refresh_widgets()
    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
