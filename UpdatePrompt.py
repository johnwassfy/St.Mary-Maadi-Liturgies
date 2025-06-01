from PyQt5.QtWidgets import QDialog, QLabel, QPushButton, QVBoxLayout, QHBoxLayout, QScrollArea, QWidget
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QFont

class UpdatePrompt(QDialog):
    def __init__(self, version, notes, parent=None):
        super().__init__(parent)
        self.setWindowTitle("تحديث جديد متوفر")
        self.setFixedSize(420, 220)
        self.setStyleSheet("""
            QDialog {
                background-color: qlineargradient(
                    x1: 0, y1: 0, x2: 1, y2: 1,
                    stop: 0 rgba(15, 46, 71, 220),
                    stop: 0.6 rgba(30, 91, 138, 220),
                    stop: 1 rgba(140, 217, 255, 180)
                );
                color: white;
                border-radius: 12px;
            }
            QLabel {
                font-size: 12pt;
                color: white;
                background: transparent;
            }
            QPushButton {
                background-color: #2ecc71;
                color: white;
                font-weight: bold;
                padding: 6px 14px;
                border: none;
                border-radius: 8px;
                font-size: 10pt;
                min-width: 100px;
            }
            QPushButton:hover {
                background-color: #27ae60;
            }
            QPushButton#cancel {
                background-color: #e74c3c;
            }
            QPushButton#cancel:hover {
                background-color: #c0392b;
            }
            QScrollArea {
                background: transparent;
                border: none;
            }
            QScrollBar:vertical {
                border: none;
                background: rgba(30, 91, 138, 100);
                width: 10px;
                margin: 0px;
                border-radius: 5px;
            }
            QScrollBar::handle:vertical {
                background: rgba(140, 217, 255, 150);
                border-radius: 5px;
            }
            QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
                background: none;
            }
        """)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(15, 15, 15, 10)

        # Title
        title = QLabel(f"تحديث جديد متوفر (الإصدار {version})")
        title.setFont(QFont("Segoe UI", 13, QFont.Bold))
        title.setAlignment(Qt.AlignCenter)
        title.setStyleSheet("color: white; background: transparent;")
        layout.addWidget(title)

        # Notes
        notes_widget = QWidget()
        notes_widget.setStyleSheet("background: transparent;")
        notes_layout = QVBoxLayout(notes_widget)
        notes_layout.setContentsMargins(0, 0, 0, 0)

        if isinstance(notes, list):
            notes_text = (
                "<div style='direction: rtl; text-align: right;'>"
                + "".join(f"<p style='margin: 0 0 6px;'>✦ {note}</p>" for note in notes)
                + "</div>"
            )
        else:
            notes_text = f"<div style='direction: rtl; text-align: right;'>{notes if notes else 'لا توجد تفاصيل.'}</div>"

        notes_label = QLabel(notes_text)
        notes_label.setWordWrap(True)
        notes_label.setAlignment(Qt.AlignRight | Qt.AlignTop)
        notes_label.setTextFormat(Qt.RichText)
        notes_label.setFont(QFont("Segoe UI", 11))
        notes_label.setStyleSheet("color: white; background: transparent; padding-right: 10px;")

        notes_layout.addWidget(notes_label)
        notes_layout.addStretch()

        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_area.setWidget(notes_widget)
        scroll_area.setStyleSheet("background: transparent; border: none;")
        layout.addWidget(scroll_area)

        # Buttons centered
        btn_layout = QHBoxLayout()
        btn_layout.addStretch()

        self.update_button = QPushButton("تحميل التحديث")
        btn_layout.addWidget(self.update_button)

        self.cancel_button = QPushButton("إلغاء")
        self.cancel_button.setObjectName("cancel")
        btn_layout.addWidget(self.cancel_button)

        btn_layout.addStretch()
        layout.addLayout(btn_layout)

        self.setLayout(layout)
