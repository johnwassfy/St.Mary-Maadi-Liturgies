from PyQt5.QtWidgets import (QDialog, QPushButton, QVBoxLayout, QLabel, QFrame, QHBoxLayout,
                              QGridLayout, QWidget, QCheckBox, QLineEdit, QListWidget,
                              QListWidgetItem, QAbstractItemView, QToolButton,
                              QFileDialog, QMessageBox, QGraphicsDropShadowEffect)
from PyQt5.QtGui import QFont, QPixmap, QColor
from PyQt5.QtCore import Qt, QSize, QMimeData, QEvent
from commonFunctions import (relative_path, list_section_rows, classify_sections_by_keyword,
                              list_hymns_by_part, get_open_presentations)
from agbyaConfig import AGBYA_PRAYERS, AGBYA_PRAYER_ORDER
from agbya import VIDEO_EXTENSIONS
import qtawesome as qta
import os

_HYMN_MIME_TYPE = "application/x-agbya-hymn-index"
_ROW_MIN_HEIGHT = 34

class HoverButton(QPushButton):
    def __init__(self, text, parent=None):
        super(HoverButton, self).__init__(text, parent)
        self.setMouseTracking(True)

    def enterEvent(self, event):
        if self.isEnabled():
            current_icon, current_icon_size = self.icon(), self.iconSize()
            self.setStyleSheet("""
                QPushButton {
                    background-color: rgba(35, 107, 142, 220);
                    border: 1px solid rgba(35, 107, 142, 250);
                    border-radius: 12px;
                    color: white;
                    padding: 8px;
                    font-size: 16px;
                    font-weight: bold;
                    text-align: center;
                }
            """)
            if not current_icon.isNull():
                self.setIcon(current_icon)
                self.setIconSize(current_icon_size)
        super().enterEvent(event)

    def leaveEvent(self, event):
        if self.isEnabled():
            current_icon, current_icon_size = self.icon(), self.iconSize()
            self.setStyleSheet(self._base_style())
            if not current_icon.isNull():
                self.setIcon(current_icon)
                self.setIconSize(current_icon_size)
        super().leaveEvent(event)

    def _base_style(self):
        return """
            QPushButton {
                background-color: rgba(240, 240, 240, 200);
                border: 1px solid #c4c4c4;
                border-radius: 12px;
                color: #333333;
                padding: 8px;
                font-size: 16px;
                font-weight: bold;
                text-align: center;
            }
        """

_DIALOG_GRADIENT_STYLE = """
    QDialog {
        background: qlineargradient(
            x1: 0, y1: 0, x2: 1, y2: 1,
            stop: 0 rgba(15, 46, 71, 245),
            stop: 0.6 rgba(30, 91, 138, 245),
            stop: 1 rgba(140, 217, 255, 245)
        );
        border-radius: 10px;
        border: 1px solid rgba(200, 200, 200, 150);
    }
    QPushButton, QToolButton {
        outline: none;
    }
"""

_HEADER_STYLE = """
    QFrame {
        background: qlineargradient(
            x1: 0, y1: 0, x2: 1, y2: 0,
            stop: 0 #1e5b8a,
            stop: 1 #3498db
        );
        border-top-left-radius: 10px;
        border-top-right-radius: 10px;
    }
"""

_SEARCH_BAR_STYLE = """
    QLineEdit {
        text-align: center;
        border: 2px solid #c4c4c4;
        border-radius: 15px;
        padding: 5px 10px;
        background-color: rgba(255, 255, 255, 220);
        font-size: 16px;
        color: #333333;
    }
    QLineEdit:focus {
        border-color: #a0a0ff;
        background-color: #ffffff;
    }
"""

_CHECKBOX_STYLE = """
    QCheckBox {
        color: white;
        font-size: 14px;
        font-weight: bold;
        background: transparent;
        spacing: 8px;
        padding: 4px;
    }
    QCheckBox::indicator {
        width: 18px;
        height: 18px;
    }
"""

_LIST_STYLE = """
    QListWidget {
        background-color: rgba(255, 255, 255, 25);
        border: 1px solid rgba(255, 255, 255, 60);
        border-radius: 8px;
        color: white;
        font-size: 13px;
        padding: 4px;
        outline: none;
    }
    QListWidget::item {
        padding: 6px;
        border-radius: 4px;
    }
    QListWidget::item:hover {
        background-color: rgba(255, 255, 255, 40);
    }
    QListWidget::item:focus {
        outline: none;
        border: none;
    }
    QScrollBar:vertical {
        background: transparent;
        width: 10px;
        margin: 2px 0px 2px 0px;
    }
    QScrollBar::handle:vertical {
        background: rgba(255, 255, 255, 110);
        min-height: 24px;
        border-radius: 5px;
    }
    QScrollBar::handle:vertical:hover {
        background: rgba(255, 255, 255, 170);
    }
    QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
        height: 0px;
        border: none;
        background: none;
    }
    QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical {
        background: none;
    }
    QScrollBar:horizontal {
        background: transparent;
        height: 10px;
        margin: 0px 2px 0px 2px;
    }
    QScrollBar::handle:horizontal {
        background: rgba(255, 255, 255, 110);
        min-width: 24px;
        border-radius: 5px;
    }
    QScrollBar::handle:horizontal:hover {
        background: rgba(255, 255, 255, 170);
    }
    QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal {
        width: 0px;
        border: none;
        background: none;
    }
    QScrollBar::add-page:horizontal, QScrollBar::sub-page:horizontal {
        background: none;
    }
"""


class _HymnsListWidget(QListWidget):
    """Drag source listing كتاب المدائح's الترانيم hymns; each row carries its hymn dict via Qt.UserRole."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setDragEnabled(True)
        self.setSelectionMode(QAbstractItemView.SingleSelection)
        self.setDragDropMode(QAbstractItemView.DragOnly)
        self.setCursor(Qt.OpenHandCursor)

    def mimeData(self, items):
        mime = QMimeData()
        if items:
            mime.setData(_HYMN_MIME_TYPE, str(self.row(items[0])).encode("utf-8"))
        return mime


class _SectionNameDialog(QDialog):
    """Small styled prompt for a new section's name — replaces QInputDialog, which rendered with
    broken/invisible native (black background, unreadable) styling under this app's global QSS."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.section_name = ""

        self.setWindowFlags(Qt.Dialog | Qt.FramelessWindowHint | Qt.WindowSystemMenuHint | Qt.WindowTitleHint)
        self.setModal(True)
        self.setFixedSize(360, 190)
        self.setStyleSheet(_DIALOG_GRADIENT_STYLE)
        self.setLayoutDirection(Qt.RightToLeft)

        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(0)
        main_layout.addWidget(_build_header("اسم القسم الجديد", "fa5s.pen", self.reject))

        content = QFrame()
        content.setStyleSheet("background: transparent; border: none;")
        content_layout = QVBoxLayout(content)
        content_layout.setContentsMargins(20, 15, 20, 15)
        content_layout.setSpacing(12)

        label = QLabel("ادخل اسم القسم الجديد:")
        label.setStyleSheet("color: white; font-size: 13px; background: transparent;")
        content_layout.addWidget(label)

        self.name_edit = QLineEdit()
        self.name_edit.setFixedHeight(36)
        self.name_edit.setLayoutDirection(Qt.RightToLeft)
        self.name_edit.setStyleSheet(_SEARCH_BAR_STYLE)
        self.name_edit.returnPressed.connect(self._confirm)
        content_layout.addWidget(self.name_edit)
        content_layout.addStretch(1)

        buttons_row = QHBoxLayout()
        ok_button = QPushButton("موافق")
        ok_button.setCursor(Qt.PointingHandCursor)
        ok_button.clicked.connect(self._confirm)
        cancel_button = QPushButton("إلغاء")
        cancel_button.setCursor(Qt.PointingHandCursor)
        cancel_button.clicked.connect(self.reject)
        for button in (ok_button, cancel_button):
            button.setMinimumHeight(34)
            button.setStyleSheet("""
                QPushButton {
                    background-color: #1e5b8a;
                    color: white;
                    border-radius: 10px;
                    font-weight: bold;
                    padding: 6px;
                    border: none;
                }
                QPushButton:hover { background-color: #3498db; }
            """)
            buttons_row.addWidget(button)
        content_layout.addLayout(buttons_row)

        main_layout.addWidget(content, 1)
        self.name_edit.setFocus()

    def _confirm(self):
        text = self.name_edit.text().strip()
        if not text:
            return
        self.section_name = text
        self.accept()


class _SectionsListWidget(QListWidget):
    """Drop target listing the prayer's sections; accepts a hymn dropped in between two sections,
    and also supports reordering any row (section or hymn) within itself by drag-and-drop."""

    _ADD_BUTTON_SIZE = 22
    _HOVER_LINE_HEIGHT = 3

    def __init__(self, hymns_list, parent=None):
        super().__init__(parent)
        self._hymns_list = hymns_list
        self._hover_row = -1
        self.setAcceptDrops(True)
        self.setDragEnabled(True)
        self.setDragDropMode(QAbstractItemView.DragDrop)
        self.setDefaultDropAction(Qt.MoveAction)
        self.setSelectionMode(QAbstractItemView.SingleSelection)
        self.setDropIndicatorShown(True)
        self.setMouseTracking(True)
        self.viewport().setMouseTracking(True)
        self.viewport().setAttribute(Qt.WA_Hover, True)

        # Floating "insert here" affordance: a bold line + a "+" button straddling it, positioned
        # by raw coordinates over whichever row is hovered. Deliberately NOT part of any row's own
        # layout/widget (earlier attempts that did that resized/rebuilt rows on every hover change).
        self._hover_line = QFrame(self.viewport())
        self._hover_line.setStyleSheet("background-color: white; border: none;")
        self._hover_line.hide()

        self._add_button = QToolButton(self.viewport())
        self._add_button.setFocusPolicy(Qt.NoFocus)
        self._add_button.setCursor(Qt.PointingHandCursor)
        # A plain "+" glyph on an explicit solid circle (rather than qtawesome's "plus-circle",
        # whose plus is a transparent cutout that barely showed up against the dark background).
        self._add_button.setStyleSheet("""
            QToolButton {
                background-color: #2ecc71;
                border: 2px solid white;
                border-radius: 11px;
            }
            QToolButton:hover {
                background-color: #27ae60;
            }
        """)
        self._add_button.setIconSize(QSize(12, 12))
        self._add_button.setIcon(qta.icon("fa5s.plus", color="white"))
        self._add_button.setToolTip("إضافة فيديو أو عرض تقديمي بعد هذا القسم")
        self._add_button.clicked.connect(self._handle_add_button_clicked)
        self._add_button.hide()

        self.verticalScrollBar().valueChanged.connect(self._update_hover_overlay)

    def viewportEvent(self, event):
        # HoverMove/Enter/Leave (not mouseMoveEvent) is what drives the QSS ":hover" rule reliably
        # across the whole item even though a child row widget covers it — mouseMoveEvent on the
        # list itself only fired for the sliver of viewport NOT covered by a child widget, which is
        # why hovering only "worked" near the icons before.
        if event.type() in (QEvent.HoverMove, QEvent.HoverEnter):
            self._set_hover_row(self.indexAt(event.pos()).row())
        elif event.type() == QEvent.HoverLeave:
            self._set_hover_row(-1)
        return super().viewportEvent(event)

    def resizeEvent(self, event):
        super().resizeEvent(event)
        self._update_hover_overlay()

    def _set_hover_row(self, row):
        if row == self._hover_row:
            return
        self._hover_row = row
        self._update_hover_overlay()

    def _update_hover_overlay(self):
        if not (0 <= self._hover_row < self.count()):
            self._hover_line.hide()
            self._add_button.hide()
            return

        rect = self.visualItemRect(self.item(self._hover_row))
        self._hover_line.setGeometry(
            rect.left(), rect.bottom() - self._HOVER_LINE_HEIGHT // 2, rect.width(), self._HOVER_LINE_HEIGHT
        )
        self._hover_line.show()
        self._hover_line.raise_()

        self._add_button.setGeometry(
            rect.right() - self._ADD_BUTTON_SIZE + 6, rect.bottom() - self._ADD_BUTTON_SIZE // 2,
            self._ADD_BUTTON_SIZE, self._ADD_BUTTON_SIZE
        )
        self._add_button.show()
        self._add_button.raise_()

    def _handle_add_button_clicked(self):
        if 0 <= self._hover_row < self.count():
            self._prompt_insert_after(self.item(self._hover_row))

    def _reset_hover(self):
        # Called after any structural change (insert/remove/reorder) since the hovered index may
        # now point at a different row or be out of range; the next mouse move re-establishes it.
        self._hover_row = -1
        self._update_hover_overlay()

    def dragEnterEvent(self, event):
        if event.mimeData().hasFormat(_HYMN_MIME_TYPE) or event.source() is self:
            event.acceptProposedAction()
        else:
            event.ignore()

    def dragMoveEvent(self, event):
        if event.mimeData().hasFormat(_HYMN_MIME_TYPE) or event.source() is self:
            # Let the base class compute/paint the insertion-point indicator line first.
            super().dragMoveEvent(event)
            event.acceptProposedAction()
        else:
            event.ignore()

    def dropEvent(self, event):
        mime = event.mimeData()

        if event.source() is self:
            # Reordering an existing row (section or hymn) within this same list.
            super().dropEvent(event)
            self.refresh_row_widgets()
            self._reset_hover()
            return

        if not mime.hasFormat(_HYMN_MIME_TYPE):
            event.ignore()
            return

        hymn_index = int(bytes(mime.data(_HYMN_MIME_TYPE)).decode("utf-8"))
        hymn = self._hymns_list.item(hymn_index).data(Qt.UserRole)["hymn"]

        # Hymns can only land in between existing sections, never before the very first one.
        drop_row = self.indexAt(event.pos()).row()
        if drop_row <= 0:
            drop_row = 1 if self.count() > 0 else 0

        item = QListWidgetItem()
        item.setData(Qt.UserRole, {"kind": "hymn", "hymn": hymn, "visible": True})
        self.insertItem(drop_row, item)
        self.refresh_row_widgets()
        self._reset_hover()
        event.acceptProposedAction()

    def _make_row_widget(self, item):
        data = item.data(Qt.UserRole)
        kind = data["kind"]
        visible = data.get("visible", True)
        name = data["hymn"]["name"] if kind == "hymn" else data["name"]

        row = QWidget()
        # Hardcoded (not self.layoutDirection()): before this list is parented into the dialog's
        # RTL hierarchy, layoutDirection() still reports LTR, which used to flip the eye icon to the
        # wrong side for rows built during __init__ vs. rows rebuilt later (e.g. after a toggle).
        row.setLayoutDirection(Qt.RightToLeft)
        row.setStyleSheet("background: transparent;")
        row.setMinimumHeight(_ROW_MIN_HEIGHT)
        layout = QHBoxLayout(row)
        layout.setContentsMargins(4, 4, 4, 4)
        layout.setSpacing(8)

        eye_button = QToolButton()
        eye_button.setAutoRaise(True)
        eye_button.setFocusPolicy(Qt.NoFocus)
        eye_button.setCursor(Qt.PointingHandCursor)
        eye_button.setIconSize(QSize(16, 16))
        eye_button.setIcon(qta.icon("fa5s.eye" if visible else "fa5s.eye-slash",
                                     color="#8bd3ff" if visible else "rgba(255, 255, 255, 110)"))
        eye_button.setToolTip("إخفاء" if visible else "إظهار")
        eye_button.clicked.connect(lambda: self._toggle_visibility(item))
        layout.addWidget(eye_button)

        drag_handle = QLabel()
        drag_handle.setPixmap(qta.icon("fa5s.grip-lines", color="rgba(255, 255, 255, 140)").pixmap(14, 14))
        drag_handle.setStyleSheet("background: transparent;")
        drag_handle.setCursor(Qt.OpenHandCursor)
        drag_handle.setToolTip("اسحب لإعادة الترتيب")
        layout.addWidget(drag_handle)

        if kind in ("hymn", "media"):
            remove_button = QToolButton()
            remove_button.setAutoRaise(True)
            remove_button.setFocusPolicy(Qt.NoFocus)
            remove_button.setCursor(Qt.PointingHandCursor)
            remove_button.setIconSize(QSize(14, 14))
            remove_button.setIcon(qta.icon("fa5s.trash-alt", color="rgba(255, 120, 120, 220)"))
            remove_button.setToolTip("إزالة هذا العنصر")
            remove_button.clicked.connect(lambda: self._remove_item(item))
            layout.addWidget(remove_button)

            if kind == "hymn":
                kind_icon_name, kind_icon_color = "fa5s.music", "#ffd54a"
            else:
                kind_icon_name = "fa5s.film" if data["media_kind"] == "video" else "fa5s.file-powerpoint"
                kind_icon_color = "#7ed6a5"
            kind_icon_label = QLabel()
            kind_icon_label.setPixmap(qta.icon(kind_icon_name, color=kind_icon_color).pixmap(14, 14))
            kind_icon_label.setStyleSheet("background: transparent;")
            layout.addWidget(kind_icon_label)

        text_color = "#ffd54a" if kind == "hymn" else "#7ed6a5" if kind == "media" else "white"
        if not visible:
            text_color = "rgba(255, 255, 255, 100)"
        text_label = QLabel(name)
        text_label.setStyleSheet(f"color: {text_color}; background: transparent; font-size: 13px;")
        layout.addWidget(text_label, 1)

        return row

    def _prompt_insert_after(self, item):
        name_dialog = _SectionNameDialog(self)
        if name_dialog.exec_() != QDialog.Accepted:
            return
        name = name_dialog.section_name

        path, _ = QFileDialog.getOpenFileName(
            self, "اختر ملف فيديو أو عرض تقديمي", "",
            "فيديو أو عرض تقديمي (*.mp4 *.wmv *.avi *.mov *.m4v *.mpg *.mpeg *.pptx *.ppt)"
        )
        if not path:
            return
        # QFileDialog always returns forward-slash paths on Windows; PowerPoint's COM automation
        # (Presentations.Open/InsertFromFile) can fail with a generic "file not found" COM error on
        # those, so normalize to the native backslash form before it's stored/used anywhere.
        path = os.path.normpath(path)

        ext = os.path.splitext(path)[1].lower()
        if ext in VIDEO_EXTENSIONS:
            media_kind = "video"
        elif ext in (".pptx", ".ppt"):
            media_kind = "pptx"
        else:
            QMessageBox.warning(self, "امتداد غير مدعوم", "الملف المختار ليس فيديو أو عرض تقديمي مدعومًا.")
            return

        new_item = QListWidgetItem()
        new_item.setData(Qt.UserRole, {
            "kind": "media", "media_kind": media_kind, "name": name.strip(),
            "path": path, "visible": True,
        })
        self.insertItem(self.row(item) + 1, new_item)
        self.refresh_row_widgets()
        self._reset_hover()

    def _item_size_hint(self, row):
        # row.sizeHint() can under-report height for a widget that's never been shown/polished
        # yet (style padding not applied), which is what made rows shrink after a refresh.
        hint = row.sizeHint()
        return QSize(hint.width(), max(hint.height(), _ROW_MIN_HEIGHT))

    def _toggle_visibility(self, item):
        data = item.data(Qt.UserRole)
        data["visible"] = not data.get("visible", True)
        item.setData(Qt.UserRole, data)
        row = self._make_row_widget(item)
        self.setItemWidget(item, row)
        item.setSizeHint(self._item_size_hint(row))

    def _remove_item(self, item):
        self.takeItem(self.row(item))
        self._reset_hover()

    def set_all_visible(self, visible):
        for i in range(self.count()):
            item = self.item(i)
            data = item.data(Qt.UserRole)
            data["visible"] = visible
            item.setData(Qt.UserRole, data)
        self.refresh_row_widgets()

    def refresh_row_widgets(self):
        for i in range(self.count()):
            item = self.item(i)
            row = self._make_row_widget(item)
            self.setItemWidget(item, row)
            # Without an explicit size hint the item reports its (empty) text's width, so
            # a section name longer than the viewport just gets clipped instead of being
            # reachable via the horizontal scrollbar.
            item.setSizeHint(self._item_size_hint(row))


def _build_header(title, icon_name, close_slot):
    header = QFrame()
    header.setFixedHeight(50)
    header.setStyleSheet(_HEADER_STYLE)

    header_layout = QHBoxLayout(header)
    header_layout.setContentsMargins(15, 0, 15, 0)

    title_layout = QHBoxLayout()
    try:
        icon_label = QLabel()
        icon_label.setPixmap(qta.icon(icon_name, color="white").pixmap(24, 24))
        icon_label.setStyleSheet("background: transparent;")
        title_layout.addWidget(icon_label)
        title_layout.addSpacing(10)
    except Exception:
        pass

    title_label = QLabel(title)
    title_font = QFont()
    title_font.setPointSize(16)
    title_font.setBold(True)
    title_label.setFont(title_font)
    title_label.setStyleSheet("color: white; background: transparent;")
    title_layout.addWidget(title_label)

    close_button = QPushButton()
    close_button.setFixedSize(30, 30)
    close_button.setCursor(Qt.PointingHandCursor)
    close_button.setStyleSheet("""
        QPushButton { background-color: transparent; border: none; }
        QPushButton:hover { background-color: rgba(255, 0, 0, 150); border-radius: 15px; }
    """)
    try:
        close_button.setIcon(qta.icon("fa5s.times", color="white"))
        close_button.setIconSize(QSize(16, 16))
    except Exception:
        close_button.setText("×")
    close_button.clicked.connect(close_slot)

    header_layout.addLayout(title_layout)
    header_layout.addStretch()
    header_layout.addWidget(close_button)
    return header


class AgbyaDialog(QDialog):
    """Lets the user pick one of the 7 Agbya prayer hours; only 'implemented' ones respond to clicks."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.selected_prayer_key = None

        self.setWindowTitle("الأجبية")
        self.setWindowFlags(Qt.Dialog | Qt.FramelessWindowHint | Qt.WindowSystemMenuHint | Qt.WindowTitleHint)
        self.setModal(True)
        self.setFixedSize(420, 380)
        self.setStyleSheet(_DIALOG_GRADIENT_STYLE)
        self.setLayoutDirection(Qt.RightToLeft)

        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(0)
        main_layout.addWidget(_build_header("الأجبية", "fa5s.book", self.reject))

        content = QFrame()
        content.setStyleSheet("background: transparent; border: none;")
        content_layout = QVBoxLayout(content)
        content_layout.setContentsMargins(15, 15, 15, 15)

        grid = QGridLayout()
        grid.setSpacing(10)
        columns = 2
        open_presentations = [p.lower() for p in get_open_presentations()]
        for index, prayer_key in enumerate(AGBYA_PRAYER_ORDER):
            config = AGBYA_PRAYERS[prayer_key]
            button = HoverButton(config["label"])
            button.setMinimumHeight(70)
            button.setFont(QFont("Calibri", 13, QFont.Bold))
            button.setLayoutDirection(Qt.RightToLeft)
            if config["implemented"]:
                button.setCursor(Qt.PointingHandCursor)
                button.setStyleSheet(button._base_style())
                button.clicked.connect(lambda _, key=prayer_key: self._choose_prayer(key))
                working_path = os.path.abspath(relative_path(config["working_file"])).lower()
                if working_path in open_presentations:
                    glow = QGraphicsDropShadowEffect(button)
                    glow.setOffset(0)
                    glow.setBlurRadius(30)
                    glow.setColor(QColor(0, 255, 0))
                    button.setGraphicsEffect(glow)
            else:
                button.setEnabled(False)
                button.setToolTip("قريباً")
                button.setStyleSheet("""
                    QPushButton {
                        background-color: rgba(120, 120, 120, 120);
                        border: 1px solid rgba(160, 160, 160, 150);
                        border-radius: 12px;
                        color: rgba(255, 255, 255, 140);
                        padding: 8px;
                        font-size: 14px;
                        font-weight: bold;
                        text-align: center;
                    }
                """)
            grid.addWidget(button, index // columns, index % columns)

        content_layout.addLayout(grid)
        content_layout.addStretch(1)
        main_layout.addWidget(content, 1)

    def _choose_prayer(self, prayer_key):
        self.selected_prayer_key = prayer_key
        self.accept()


class AgbyaSetupDialog(QDialog):
    """Combines the sub-service checkbox picker (when the prayer has any) with the open/construct
    action choice in a single dialog, instead of two separate dialogs chained via a "متابعة" button."""

    def __init__(self, parent, prayer_key):
        super().__init__(parent)
        config = AGBYA_PRAYERS[prayer_key]
        self.sub_services = config["sub_services"] or []
        self.selected_services = []
        self.action = None

        self.setWindowTitle(config["label"])
        self.setWindowFlags(Qt.Dialog | Qt.FramelessWindowHint | Qt.WindowSystemMenuHint | Qt.WindowTitleHint)
        self.setModal(True)
        self.setFixedSize(420, 420 if self.sub_services else 260)
        self.setStyleSheet(_DIALOG_GRADIENT_STYLE)
        self.setLayoutDirection(Qt.RightToLeft)

        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(0)
        main_layout.addWidget(_build_header(config["label"], "fa5s.pray", self.reject))

        content = QFrame()
        content.setStyleSheet("background: transparent; border: none;")
        content_layout = QVBoxLayout(content)
        content_layout.setContentsMargins(20, 18, 20, 18)
        content_layout.setSpacing(14)

        self._checkboxes = []
        if self.sub_services:
            services_frame = QFrame()
            services_frame.setStyleSheet("""
                QFrame {
                    background-color: rgba(255, 255, 255, 22);
                    border: 1px solid rgba(255, 255, 255, 55);
                    border-radius: 10px;
                }
            """)
            services_layout = QVBoxLayout(services_frame)
            services_layout.setContentsMargins(15, 12, 15, 12)
            services_layout.setSpacing(8)

            services_label = QLabel("اختر الخدمات المطلوبة")
            services_label.setStyleSheet("color: white; font-size: 13px; font-weight: bold; background: transparent;")
            services_layout.addWidget(services_label)
            for service_key, arabic_label, _keyword in self.sub_services:
                checkbox = QCheckBox(arabic_label)
                checkbox.setChecked(True)
                checkbox.setStyleSheet(_CHECKBOX_STYLE)
                checkbox.setLayoutDirection(Qt.RightToLeft)
                checkbox.stateChanged.connect(self._validate)
                services_layout.addWidget(checkbox)
                self._checkboxes.append((service_key, checkbox))

            content_layout.addWidget(services_frame)

        disabled_style = "QPushButton:disabled { background-color: rgba(120, 120, 120, 120); color: rgba(255, 255, 255, 140); }"

        self.open_button = HoverButton("فتح مباشر")
        self.open_button.setMinimumHeight(60)
        self.open_button.setFont(QFont("Calibri", 14, QFont.Bold))
        self.open_button.setCursor(Qt.PointingHandCursor)
        self.open_button.setStyleSheet(self.open_button._base_style() + disabled_style)
        self.open_button.setIcon(qta.icon("fa5s.play", color="#1e5b8a"))
        self.open_button.setIconSize(QSize(22, 22))
        self.open_button.clicked.connect(lambda: self._choose("open"))

        self.construct_button = HoverButton("إنشاء اجتماع الصلاة")
        self.construct_button.setMinimumHeight(60)
        self.construct_button.setFont(QFont("Calibri", 14, QFont.Bold))
        self.construct_button.setCursor(Qt.PointingHandCursor)
        self.construct_button.setStyleSheet(self.construct_button._base_style() + disabled_style)
        self.construct_button.setIcon(qta.icon("fa5s.tools", color="#1e5b8a"))
        self.construct_button.setIconSize(QSize(22, 22))
        self.construct_button.clicked.connect(lambda: self._choose("construct"))

        content_layout.addWidget(self.open_button)
        content_layout.addWidget(self.construct_button)
        content_layout.addStretch(1)

        main_layout.addWidget(content, 1)
        self._validate()

    def _validate(self):
        any_checked = (not self.sub_services) or any(checkbox.isChecked() for _, checkbox in self._checkboxes)
        self.open_button.setEnabled(any_checked)
        self.construct_button.setEnabled(any_checked)

    def _choose(self, action):
        self.selected_services = [key for key, checkbox in self._checkboxes if checkbox.isChecked()]
        self.action = action
        self.accept()


class AgbyaGatheringBuilderDialog(QDialog):
    """Lets the user hide/show individual sections (checked = shown) before opening a constructed gathering."""

    def __init__(self, parent, prayer_key, selected_sub_services=None):
        super().__init__(parent)
        config = AGBYA_PRAYERS[prayer_key]
        self.hidden_section_ids = []
        self.hymn_insertions = []
        self.media_insertions = []

        excel_path = relative_path(r"Files Data.xlsx")
        if config["sub_services"]:
            keyword_map = {service_key: keyword for service_key, _, keyword in config["sub_services"]}
            buckets = classify_sections_by_keyword(excel_path, config["sheet"], keyword_map)
            allowed_guids = set(buckets["fixed"])
            for service_key in (selected_sub_services or []):
                allowed_guids.update(buckets.get(service_key, []))
            rows = [(name, guid) for name, guid in list_section_rows(excel_path, config["sheet"]) if guid in allowed_guids]
        else:
            rows = list_section_rows(excel_path, config["sheet"])

        hymns = list_hymns_by_part(excel_path, "المدائح", "الباب الثالث")

        self.setWindowTitle(f"إنشاء اجتماع صلاة - {config['label']}")
        self.setWindowFlags(Qt.Dialog | Qt.FramelessWindowHint | Qt.WindowSystemMenuHint | Qt.WindowTitleHint)
        self.setModal(True)
        self.setFixedSize(600, 520)
        self.setStyleSheet(_DIALOG_GRADIENT_STYLE)
        self.setLayoutDirection(Qt.RightToLeft)

        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(0)
        main_layout.addWidget(_build_header(f"إنشاء اجتماع صلاة - {config['label']}", "fa5s.tools", self.reject))

        content = QFrame()
        content.setStyleSheet("background: transparent; border: none;")
        content_layout = QVBoxLayout(content)
        content_layout.setContentsMargins(15, 10, 15, 10)

        self.search_bar = QLineEdit()
        self.search_bar.setPlaceholderText("بحث في أقسام الصلاة")
        self.search_bar.setFixedHeight(32)
        self.search_bar.setLayoutDirection(Qt.RightToLeft)
        self.search_bar.setStyleSheet(_SEARCH_BAR_STYLE)
        self.search_bar.textChanged.connect(self._filter_rows)

        lists_row = QHBoxLayout()
        lists_row.setSpacing(15)

        sections_panel = QVBoxLayout()
        sections_label = QLabel("أقسام الصلاة (اسحب للترتيب، واسحب الترنيمة هنا للإدراج بين قسمين)")
        sections_label.setStyleSheet("color: white; font-size: 13px; font-weight: bold; background: transparent;")
        sections_panel.addWidget(sections_label)

        visibility_row = QHBoxLayout()
        show_all_button = QPushButton("إظهار الكل")
        hide_all_button = QPushButton("إخفاء الكل")
        for button in (show_all_button, hide_all_button):
            button.setCursor(Qt.PointingHandCursor)
            button.setMinimumHeight(28)
            button.setStyleSheet("""
                QPushButton {
                    background-color: rgba(255, 255, 255, 30);
                    border: 1px solid rgba(255, 255, 255, 70);
                    border-radius: 8px;
                    color: white;
                    font-size: 12px;
                    padding: 4px 10px;
                }
                QPushButton:hover { background-color: rgba(255, 255, 255, 55); }
            """)
        visibility_row.addWidget(self.search_bar, 1)
        visibility_row.addWidget(show_all_button)
        visibility_row.addWidget(hide_all_button)
        sections_panel.addLayout(visibility_row)

        self.sections_list = _SectionsListWidget(None)  # hymns_list wired below
        self.sections_list.setStyleSheet(_LIST_STYLE)
        for name, guid in rows:
            item = QListWidgetItem()
            item.setData(Qt.UserRole, {"kind": "section", "guid": guid, "name": str(name), "visible": True})
            self.sections_list.addItem(item)
        self.sections_list.refresh_row_widgets()
        show_all_button.clicked.connect(lambda: self.sections_list.set_all_visible(True))
        hide_all_button.clicked.connect(lambda: self.sections_list.set_all_visible(False))
        sections_panel.addWidget(self.sections_list)

        hymns_panel = QVBoxLayout()
        hymns_label = QLabel(f"الترانيم ({len(hymns)}) — اسحب لإدراجها")
        hymns_label.setStyleSheet("color: white; font-size: 13px; font-weight: bold; background: transparent;")
        self.hymn_search_bar = QLineEdit()
        self.hymn_search_bar.setPlaceholderText("بحث في الترانيم")
        self.hymn_search_bar.setFixedHeight(32)
        self.hymn_search_bar.setLayoutDirection(Qt.RightToLeft)
        self.hymn_search_bar.setStyleSheet(_SEARCH_BAR_STYLE)
        self.hymn_search_bar.textChanged.connect(self._filter_hymns)
        self.hymns_list = _HymnsListWidget()
        self.hymns_list.setStyleSheet(_LIST_STYLE)
        for hymn in hymns:
            item = QListWidgetItem(str(hymn["name"]))
            item.setIcon(qta.icon("fa5s.music", color="white"))
            item.setToolTip("اسحب لإدراجها بين قسمين")
            item.setData(Qt.UserRole, {"hymn": hymn})
            self.hymns_list.addItem(item)
        hymns_panel.addWidget(hymns_label)
        hymns_panel.addWidget(self.hymn_search_bar)
        hymns_panel.addWidget(self.hymns_list)

        self.sections_list._hymns_list = self.hymns_list

        lists_row.addLayout(sections_panel, 2)
        lists_row.addLayout(hymns_panel, 1)
        content_layout.addLayout(lists_row, 1)

        hint_label = QLabel("💡 اسحب لترتيب الأقسام أو لإدراج ترنيمة، مرر الفأرة فوق قسم وانقر ➕ لإضافة فيديو أو عرض تقديمي، 👁 للإظهار/الإخفاء و🗑 لإزالة")
        hint_label.setStyleSheet("color: rgba(255, 255, 255, 180); font-size: 11px; background: transparent;")
        hint_label.setAlignment(Qt.AlignCenter)
        content_layout.addWidget(hint_label)
        content_layout.addSpacing(6)

        buttons_row = QHBoxLayout()
        open_button = QPushButton("فتح")
        open_button.setCursor(Qt.PointingHandCursor)
        open_button.clicked.connect(self._confirm)
        cancel_button = QPushButton("إلغاء")
        cancel_button.setCursor(Qt.PointingHandCursor)
        cancel_button.clicked.connect(self.reject)
        for button in (open_button, cancel_button):
            button.setMinimumHeight(38)
            button.setStyleSheet("""
                QPushButton {
                    background-color: #1e5b8a;
                    color: white;
                    border-radius: 12px;
                    font-weight: bold;
                    padding: 6px;
                    border: none;
                }
                QPushButton:hover { background-color: #3498db; }
            """)
        buttons_row.addWidget(open_button)
        buttons_row.addWidget(cancel_button)
        content_layout.addLayout(buttons_row)

        main_layout.addWidget(content, 1)

    def _filter_rows(self, text):
        normalized = text.strip().lower()
        for i in range(self.sections_list.count()):
            item = self.sections_list.item(i)
            data = item.data(Qt.UserRole)
            name = data["hymn"]["name"] if data["kind"] == "hymn" else data["name"]
            item.setHidden(normalized not in name.lower())

    def _filter_hymns(self, text):
        normalized = text.strip().lower()
        for i in range(self.hymns_list.count()):
            item = self.hymns_list.item(i)
            item.setHidden(normalized not in item.text().lower())

    def _confirm(self):
        hidden_section_ids = []
        hymn_insertions = []
        media_insertions = []
        last_section_guid = None

        for i in range(self.sections_list.count()):
            item = self.sections_list.item(i)
            data = item.data(Qt.UserRole)
            if data["kind"] == "section":
                last_section_guid = data["guid"]
                if not data.get("visible", True):
                    hidden_section_ids.append(data["guid"])
            elif data["kind"] == "hymn" and data.get("visible", True) and last_section_guid is not None:
                hymn = data["hymn"]
                hymn_insertions.append({
                    "after_guid": last_section_guid,
                    "name": hymn["name"],
                    "first_slide": hymn["first_slide"],
                    "last_slide": hymn["last_slide"],
                    "num_slides": hymn["num_slides"],
                })
            elif data["kind"] == "media" and data.get("visible", True) and last_section_guid is not None:
                media_insertions.append({
                    "after_guid": last_section_guid,
                    "name": data["name"],
                    "media_kind": data["media_kind"],
                    "path": data["path"],
                })

        self.hidden_section_ids = hidden_section_ids
        self.hymn_insertions = hymn_insertions
        self.media_insertions = media_insertions
        self.accept()

