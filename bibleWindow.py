from PyQt5.QtWidgets import (
    QMainWindow, QLabel, QPushButton, QVBoxLayout, QHBoxLayout,
    QGridLayout, QFrame, QScrollArea, QWidget, QMessageBox, QLineEdit,
)
from PyQt5.QtGui import QPixmap, QIntValidator
from PyQt5.QtCore import Qt

import bibleNavigationData as bible_nav
from commonFunctions import relative_path, load_background_image, open_presentation_on_slide_safe
from NotificationBar import NotificationBar  # Assuming NotificationBar is in the same directory
from WorkerThread import WorkerThread

CHAPTER_GRID_COLUMNS = 3
VERSE_GRID_COLUMNS = 3
CHOOSE_BOOK_PLACEHOLDER = "اختر سفرًا"
CHOOSE_CHAPTER_PLACEHOLDER = "اختر إصحاحًا"

# Deuterocanonical addition with no Phase 3 deck yet; kept visible but disabled.
UNAVAILABLE_BOOK_LABEL = "صلاة منسى"


def _build_manifest_task(testament, number, path, **_worker_kwargs):
    # WorkerThread always injects progress_callback; this task doesn't report progress.
    return bible_nav.get_book_manifest(testament, number, path)


def _open_slide_task(path, slide_number, **_worker_kwargs):
    return open_presentation_on_slide_safe(path, slide_number)


class bibleWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        
        self.setWindowTitle("St. Mary Maadi Liturgies")
        self.setGeometry(100, 100, 625, 600)
        self.setFixedSize(625, 600)

        self.current_book = None       # {testament, number, name, shortcut, path, chapters}
        self.current_chapter = None    # {number, start_slide, verses}
        self._background_workers = []  # keep refs so background QThreads aren't garbage-collected
        self._busy_operation = False   # only one WorkerThread may run at a time (see _run_worker)
        self._closing = False

        # Create a central widget
        self.central_widget = QLabel(self)
        self.central_widget.setAlignment(Qt.AlignCenter)
        self.central_widget.setGeometry(0, 0, self.width(), self.height())
        self.setCentralWidget(self.central_widget)

        # Create a vertical layout for the central widget
        layout = QVBoxLayout(self.central_widget)

        button_width = 100
        button_height = 30
        button_x = self.width() - button_width - 10
        button_y = self.height() - button_height - 10
        self.back_button = QPushButton("Back", self)
        self.back_button.setGeometry(button_x, button_y, button_width, button_height)
        self.back_button.clicked.connect(self.go_back)
        self.back_button.setText("⬅ العودة")
        self.back_button.setStyleSheet("""
            QPushButton {
                background-color: #e67e22;
                color: white;
                font-weight: bold;
                border-radius: 12px;
                padding: 6px 14px;
                font-size: 11pt;
            }
            QPushButton:hover {
                background-color: #d35400;
            }
        """)
        layout.addWidget(self.back_button, alignment=Qt.AlignBottom | Qt.AlignRight)
        
        # Add NotificationBar
        self.notification_bar = NotificationBar(self)
        self.notification_bar.setGeometry(0, 70, self.width(), 50)

        # Load background image
        try:
            load_background_image(self.central_widget)
        except Exception as e:
            self.notification_bar.show_message(f"خطأ في تحميل الخلفية: {str(e)}")

        frame0 = QFrame(self)
        frame0.setGeometry(0, 0, 625, 80)
        image_label = QLabel(frame0)
        image_label.setGeometry(0, 0, 625, 80)
        image_path = relative_path(r"Data\الصور\Untitled-4.png")
        pixmap = QPixmap(image_path)
        image_label.setPixmap(pixmap)

        # Verses, chapters and the book list stay visible side by side, so picking a
        # different book or chapter never hides the other levels.
        self.frame = QFrame(self)
        self.frame.setGeometry(20, 90, 585, 450)
        self.frame.setStyleSheet("QFrame { background-color: rgba(204, 178, 119, 200); border: 2px solid black; }")
        frame_layout = QHBoxLayout(self.frame)
        frame_layout.setContentsMargins(8, 8, 8, 8)
        frame_layout.setSpacing(8)

        self.verses_breadcrumb, self.verses_grid_layout, self.verses_jump_field = \
            self._build_number_pane(frame_layout, "رقم الآية", self._on_verse_jump, stretch=3)
        frame_layout.addWidget(self._make_divider())
        self.chapters_breadcrumb, self.chapters_grid_layout, self.chapters_jump_field = \
            self._build_number_pane(frame_layout, "رقم الإصحاح", self._on_chapter_jump, stretch=3)
        frame_layout.addWidget(self._make_divider())
        self._build_books_pane(frame_layout, stretch=3)

        self.chapters_breadcrumb.setText(CHOOSE_BOOK_PLACEHOLDER)
        self.verses_breadcrumb.setText(CHOOSE_CHAPTER_PLACEHOLDER)

    def _make_divider(self):
        line = QFrame(self.frame)
        line.setFrameShape(QFrame.VLine)
        line.setFrameShadow(QFrame.Sunken)
        line.setStyleSheet("background-color: black;")
        return line

    # ------------------------------------------------------------------
    # Books pane
    # ------------------------------------------------------------------
    def _build_books_pane(self, parent_layout, stretch):
        pane_layout = QVBoxLayout()
        pane_layout.setContentsMargins(0, 0, 0, 0)

        self.buttons_layout = QVBoxLayout()
        self._populate_book_buttons()

        scroll_area = QScrollArea()
        scroll_area.setStyleSheet("background-color: transparent; border: none; color: white;")
        scroll_area.setWidgetResizable(True)
        scroll_content = QWidget()
        scroll_content.setLayout(self.buttons_layout)
        scroll_area.setWidget(scroll_content)
        scroll_area.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        scroll_area.verticalScrollBar().setStyleSheet(
            "QScrollBar:vertical {border: none; background: transparent; width: 10px;}"
            "QScrollBar::handle:vertical {background: rgba(255, 255, 255, 100); border-radius: 5px;}"
            "QScrollBar::add-line:vertical {background: none;}"
            "QScrollBar::sub-line:vertical {background: none;}"
            "QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical {background: none;}"
        )
        pane_layout.addWidget(scroll_area)

        container = QWidget()
        container.setLayout(pane_layout)
        parent_layout.addWidget(container, stretch)

    def _populate_book_buttons(self):
        books = bible_nav.list_books()
        old_testament_numbers = [number for testament, number, *_ in books if testament == bible_nav.OLD_TESTAMENT]
        last_old_testament_number = max(old_testament_numbers) if old_testament_numbers else None

        current_testament = None
        for testament, number, name, shortcut, path in books:
            if testament != current_testament:
                current_testament = testament
                self.buttons_layout.addWidget(self._make_section_label(testament))

            button = QPushButton(name)
            button.setToolTip(name)
            self._set_default_button_style(button)
            if path is None:
                self._disable_button(button)
            else:
                button.clicked.connect(
                    lambda _, t=testament, n=number, bn=name, s=shortcut, p=path:
                        self._open_book(t, n, bn, s, p)
                )
            self.buttons_layout.addWidget(button)

            # Prayer of Manasseh has no Phase 3 deck; keep a disabled placeholder in place.
            if testament == bible_nav.OLD_TESTAMENT and number == last_old_testament_number:
                placeholder = QPushButton(UNAVAILABLE_BOOK_LABEL)
                self._set_default_button_style(placeholder)
                self._disable_button(placeholder)
                self.buttons_layout.addWidget(placeholder)

    def _disable_button(self, button):
        button.setEnabled(False)
        button.setToolTip("غير متاح حاليًا")
        button.setStyleSheet(
            "QPushButton {"
            "   background-color: rgba(200, 200, 200, 80);"
            "   border: 1px solid #bbbbbb;"
            "   border-radius: 5px;"
            "   color: #999999;"
            "   padding: 6px 8px;"
            "   font-size: 18px;"
            "   font-family: 'Arial';"
            "   font-weight: bold;"
            "}"
        )

    def _make_section_label(self, testament):
        label = QLabel("العهد القديم" if testament == bible_nav.OLD_TESTAMENT else "العهد الجديد")
        label.setStyleSheet(
            "color: #4a3418; font-size: 15px; font-weight: bold;"
            " padding: 8px 4px 2px 4px; background: transparent;"
        )
        return label

    def _set_default_button_style(self, button):
        button.setStyleSheet(
            "QPushButton {"
            "   background-color: rgba(240, 240, 240, 100);"
            "   border: 1px solid #c4c4c4;"
            "   border-radius: 5px;"
            "   color: #333333;"
            "   padding: 6px 8px;"
            "   font-size: 18px;"
            "   font-family: 'Arial';" 
            "   font-weight: bold;"
            "}"
            "QPushButton:hover {"
            "   background-color: #e0e0e0;"
            "}"
            "QPushButton:pressed {"
            "   background-color: #d9d9d9;"
            "}"
        )

    # ------------------------------------------------------------------
    # Chapters / verses panes (breadcrumb + quick-jump + numeric grid), always visible
    # ------------------------------------------------------------------
    def _build_number_pane(self, parent_layout, jump_placeholder, on_jump, stretch):
        pane_layout = QVBoxLayout()
        pane_layout.setContentsMargins(0, 0, 0, 0)

        breadcrumb = QLabel("")
        breadcrumb.setWordWrap(True)
        breadcrumb.setStyleSheet("color: #4a3418; font-size: 14px; font-weight: bold; background: transparent;")
        pane_layout.addWidget(breadcrumb)

        jump_field = QLineEdit()
        jump_field.setPlaceholderText(jump_placeholder)
        jump_field.setAlignment(Qt.AlignCenter)
        jump_field.setStyleSheet(
            "QLineEdit {"
            "   background-color: rgba(255, 255, 255, 160);"
            "   border: 1px solid #c4c4c4;"
            "   border-radius: 5px;"
            "   padding: 4px;"
            "   font-size: 14px;"
            "}"
        )
        jump_field.returnPressed.connect(lambda: on_jump(jump_field.text()))
        pane_layout.addWidget(jump_field)

        scroll_area = QScrollArea()
        scroll_area.setStyleSheet("background-color: transparent; border: none;")
        scroll_area.setWidgetResizable(True)
        scroll_area.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        scroll_area.verticalScrollBar().setStyleSheet(
            "QScrollBar:vertical {border: none; background: transparent; width: 10px;}"
            "QScrollBar::handle:vertical {background: rgba(255, 255, 255, 100); border-radius: 5px;}"
            "QScrollBar::add-line:vertical {background: none;}"
            "QScrollBar::sub-line:vertical {background: none;}"
            "QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical {background: none;}"
        )
        scroll_content = QWidget()
        grid_layout = QGridLayout(scroll_content)
        grid_layout.setSpacing(4)
        scroll_content.setLayout(grid_layout)
        scroll_area.setWidget(scroll_content)
        pane_layout.addWidget(scroll_area)

        container = QWidget()
        container.setLayout(pane_layout)
        parent_layout.addWidget(container, stretch)

        return breadcrumb, grid_layout, jump_field

    def _make_grid_button(self, number, on_click):
        button = QPushButton(str(number))
        button.setFixedSize(48, 34)
        button.setStyleSheet(
            "QPushButton {"
            "   background-color: rgba(240, 240, 240, 140);"
            "   border: 1px solid #c4c4c4;"
            "   border-radius: 6px;"
            "   color: #333333;"
            "   font-size: 13px;"
            "   font-weight: bold;"
            "}"
            "QPushButton:hover { background-color: #e0e0e0; }"
            "QPushButton:pressed { background-color: #d9d9d9; }"
        )
        button.clicked.connect(lambda: on_click(number))
        return button

    def _fill_grid(self, grid_layout, numbers, on_click, columns):
        self._clear_layout(grid_layout)
        for index, number in enumerate(numbers):
            row, col = divmod(index, columns)
            grid_layout.addWidget(self._make_grid_button(number, on_click), row, col)

    def _clear_layout(self, layout):
        while layout.count():
            item = layout.takeAt(0)
            widget = item.widget()
            if widget is not None:
                widget.deleteLater()

    # ------------------------------------------------------------------
    # Navigation: books -> chapters -> verses -> PowerPoint (all panes stay visible)
    # ------------------------------------------------------------------
    def _open_book(self, testament, number, name, shortcut, path):
        worker = WorkerThread(_build_manifest_task, testament, number, path)
        self._run_worker(
            worker,
            on_result=lambda manifest: self._on_book_ready(testament, number, name, shortcut, path, manifest),
            on_error=lambda msg: self._on_background_error("تعذر تحضير السفر", msg),
        )

    def _on_book_ready(self, testament, number, name, shortcut, path, manifest):
        try:
            self.current_book = {
                "testament": testament,
                "number": number,
                "name": name,
                "shortcut": shortcut,
                "path": path,
                "chapters": manifest["chapters"],
            }
            self.current_chapter = None

            self.chapters_breadcrumb.setText(name)
            chapter_numbers = sorted((int(k) for k in manifest["chapters"]))
            self.chapters_jump_field.setValidator(QIntValidator(1, chapter_numbers[-1], self))
            self.chapters_jump_field.clear()
            self._fill_grid(self.chapters_grid_layout, chapter_numbers, self._open_chapter, CHAPTER_GRID_COLUMNS)

            # A new book invalidates whichever chapter's verses were shown before.
            self.verses_breadcrumb.setText(CHOOSE_CHAPTER_PLACEHOLDER)
            self.verses_jump_field.clear()
            self._clear_layout(self.verses_grid_layout)
        except RuntimeError:
            pass  # window/widgets were torn down while this result was in flight

    def _on_chapter_jump(self, text):
        if text.strip().isdigit():
            self._open_chapter(int(text.strip()))

    def _open_chapter(self, chapter_number):
        if self.current_book is None:
            return
        chapter_data = self.current_book["chapters"].get(str(chapter_number))
        if chapter_data is None:
            self.notification_bar.show_message("رقم الإصحاح غير موجود")
            return

        self.current_chapter = {"number": chapter_number, **chapter_data}
        self.verses_breadcrumb.setText(f"{self.current_book['name']} — الإصحاح {chapter_number}")
        verse_numbers = sorted((int(k) for k in chapter_data["verses"]))
        if verse_numbers:
            self.verses_jump_field.setValidator(QIntValidator(1, verse_numbers[-1], self))
        self.verses_jump_field.clear()
        self._fill_grid(self.verses_grid_layout, verse_numbers, self._open_verse, VERSE_GRID_COLUMNS)

        # Clicking a chapter both opens it and shows the verse grid, per the confirmed design.
        self._open_slide_async(self.current_book["path"], chapter_data["start_slide"])

    def _on_verse_jump(self, text):
        if text.strip().isdigit():
            self._open_verse(int(text.strip()))

    def _open_verse(self, verse_number):
        if self.current_book is None or self.current_chapter is None:
            return
        slide_number = self.current_chapter["verses"].get(str(verse_number))
        if slide_number is None:
            self.notification_bar.show_message("رقم الآية غير موجود")
            return
        self._open_slide_async(self.current_book["path"], slide_number)

    def _open_slide_async(self, path, slide_number):
        worker = WorkerThread(_open_slide_task, path, slide_number)
        self._run_worker(
            worker,
            on_error=lambda msg: self._on_background_error("تعذر فتح العرض التقديمي", msg),
        )

    def _run_worker(self, worker, on_result=None, on_error=None):
        """Run at most one background task at a time; ignores the call if one is already in flight.

        Rapid repeated clicks used to spawn several concurrent WorkerThreads touching
        python-pptx/COM at once, which could crash the whole process; the whole frame is
        disabled meanwhile so clicks physically can't queue up more overlapping work.
        """
        if self._busy_operation:
            return
        self._busy_operation = True
        self.frame.setEnabled(False)
        if on_result is not None:
            worker.result.connect(on_result)
        if on_error is not None:
            worker.error.connect(on_error)
        worker.finished.connect(self._on_worker_finished)
        self._track_worker(worker)
        worker.start()

    def _on_worker_finished(self):
        self._busy_operation = False
        try:
            if not self._closing:
                self.frame.setEnabled(True)
        except RuntimeError:
            pass  # window/widgets were torn down while this worker was running

    def _track_worker(self, worker):
        self._background_workers.append(worker)
        worker.finished.connect(lambda: self._forget_worker(worker))

    def _forget_worker(self, worker):
        if worker in self._background_workers:
            self._background_workers.remove(worker)

    def _on_background_error(self, prefix, message):
        try:
            first_line = message.splitlines()[0] if message else message
            self.notification_bar.show_message(f"{prefix}: {first_line}")
        except RuntimeError:
            pass  # window/widgets were torn down while this error was in flight

    def go_back(self):
        self.close()

    def closeEvent(self, event):
        # Let any in-flight worker finish before the window (and its widgets) go away,
        # so a late result/error signal can't fire into already-deleted widgets.
        self._closing = True
        for worker in list(self._background_workers):
            if worker.isRunning():
                worker.wait(3000)
        super().closeEvent(event)

    def show_error_message(self, error_message):
        QMessageBox.critical(self, "Error", error_message)
