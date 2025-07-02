from PyQt5.QtWidgets import (QDialog, QPushButton, QVBoxLayout, QLabel, QFrame, QHBoxLayout,
                           QScrollArea, QWidget, QGraphicsDropShadowEffect)
from PyQt5.QtGui import QFont, QPixmap, QColor
from PyQt5.QtCore import Qt, QSize
from commonFunctions import relative_path, open_presentation_relative_path
import qtawesome as qta

class Elbas5aDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.selected_option = None
        
        self.setWindowTitle("أسبوع الآلام")
        self.setFixedSize(550, 480)
        
        # Make dialog modal - will be attached to parent window
        self.setWindowFlags(Qt.Dialog | Qt.FramelessWindowHint | Qt.WindowSystemMenuHint | Qt.WindowTitleHint)
        self.setModal(True)
        
        # Main layout
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(0)
        
        # Set gradient background for the entire dialog
        self.setStyleSheet("""
            QDialog {
                background: qlineargradient(
                    x1: 0, y1: 0, x2: 1, y2: 1,
                    stop: 0 rgba(107, 6, 6, 245),
                    stop: 0.6 rgba(140, 30, 30, 245),
                    stop: 1 rgba(180, 80, 80, 245)
                );
                border-radius: 10px;
                border: 1px solid rgba(200, 200, 200, 150);
            }
        """)
        
        # Header
        header = self.create_header()
        main_layout.addWidget(header)
        
        # Main content
        content_container = QFrame()
        content_container.setStyleSheet("background: transparent; border: none;")
        content_layout = QHBoxLayout(content_container)
        content_layout.setContentsMargins(15, 10, 15, 10)
                
        # Buttons panel on the left with scroll area
        buttons_panel = self.create_buttons_panel()
        content_layout.addWidget(buttons_panel)

        # Add a little spacing between photo and buttons
        content_layout.addSpacing(15)
        
        # Photo panel on the right
        photo_panel = self.create_photo_panel()
        content_layout.addWidget(photo_panel)

        main_layout.addWidget(content_container, 1)  # 1 = stretch factor
        
        # Set Arabic RTL layout
        self.setLayoutDirection(Qt.RightToLeft)
        
    def create_header(self):
        header = QFrame()
        header.setFixedHeight(50)
        header.setStyleSheet("""
            QFrame {
                background: qlineargradient(
                    x1: 0, y1: 0, x2: 1, y2: 0,
                    stop: 0 #8b0000,
                    stop: 1 #c03232
                );
                border-top-left-radius: 10px;
                border-top-right-radius: 10px;
            }
        """)
        
        header_layout = QHBoxLayout(header)
        header_layout.setContentsMargins(15, 0, 15, 0)
        
        # Title with icon
        title_layout = QHBoxLayout()
        
        # Add icon (optional)
        try:
            icon_label = QLabel()
            icon = qta.icon("fa5s.cross", color="white").pixmap(24, 24)
            icon_label.setPixmap(icon)
            icon_label.setStyleSheet("background: transparent;")
            title_layout.addWidget(icon_label)
            title_layout.addSpacing(10)
        except:
            pass
        
        # Title
        title_label = QLabel("أسبوع الآلام")
        title_font = QFont()
        title_font.setPointSize(16)
        title_font.setBold(True)
        title_label.setFont(title_font)
        title_label.setStyleSheet("color: white; background: transparent;")
        title_layout.addWidget(title_label)
        
        # Close button in header
        close_button = QPushButton()
        close_button.setFixedSize(30, 30)
        close_button.setStyleSheet("""
            QPushButton {
                background-color: transparent;
                border: none;
            }
            QPushButton:hover {
                background-color: rgba(255, 0, 0, 150);
                border-radius: 15px;
            }
        """)
        close_button.setCursor(Qt.PointingHandCursor)
        
        # Set X icon
        try:
            close_button.setIcon(qta.icon("fa5s.times", color="white"))
            close_button.setIconSize(QSize(16, 16))
        except:
            close_button.setText("×")
            close_button.setStyleSheet("""
                QPushButton {
                    color: white;
                    font-size: 16pt;
                    font-weight: bold;
                    background-color: transparent;
                    border: none;
                }
                QPushButton:hover {
                    background-color: rgba(255, 0, 0, 150);
                    border-radius: 15px;
                }
            """)
        
        close_button.clicked.connect(self.reject)
        
        header_layout.addLayout(title_layout)
        header_layout.addStretch()
        header_layout.addWidget(close_button)
        
        return header

    def create_photo_panel(self):
        # Photo panel with border
        photo_frame = QFrame()
        
        photo_layout = QVBoxLayout(photo_frame)
        photo_layout.setContentsMargins(0, 0, 0, 0)
        photo_layout.setSpacing(0)
        
        # Add the photo with larger dimensions
        try:
            photo_label = QLabel()
            pixmap = QPixmap(relative_path(r"Data\الصور\esbo3elalam.png"))
            
            # Make photo larger while maintaining aspect ratio
            pixmap = pixmap.scaled(220, 320, Qt.KeepAspectRatio, Qt.SmoothTransformation)
            
            photo_label.setPixmap(pixmap)
            photo_label.setAlignment(Qt.AlignCenter)
            photo_label.setStyleSheet("background: transparent; border: none;")
            photo_layout.addWidget(photo_label, 1, alignment=Qt.AlignCenter)
        except Exception as e:
            # Fallback text if image fails to load
            fallback = QLabel("صورة أسبوع الآلام")
            fallback.setAlignment(Qt.AlignCenter)
            fallback.setStyleSheet("color: white; font-size: 14px; background: transparent;")
            photo_layout.addWidget(fallback)
    
        return photo_frame

    def create_buttons_panel(self):
        buttons_frame = QFrame()
        buttons_frame.setMinimumWidth(250)
        
        # Main layout for the buttons frame
        buttons_frame_layout = QVBoxLayout(buttons_frame)
        buttons_frame_layout.setContentsMargins(0, 0, 0, 0)
        
        # Create scroll area
        scroll_area = QScrollArea()
        scroll_area.setStyleSheet("""
            QScrollArea {
                background-color: transparent; 
                border: none;
            }
        """)
        scroll_area.setWidgetResizable(True)
        scroll_area.setMinimumHeight(350)
        
        # Set stylesheet for scrollbar
        scroll_area.verticalScrollBar().setStyleSheet("""
            QScrollBar:vertical {
                border: none;
                background: transparent;
                width: 10px;
            }
            QScrollBar::handle:vertical {
                background: rgba(255, 255, 255, 100);
                border-radius: 5px;
            }
            QScrollBar::add-line:vertical {
                background: none;
            }
            QScrollBar::sub-line:vertical {
                background: none;
            }
            QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical {
                background: none;
            }
        """)
        
        # Create content widget for the scroll area
        scroll_content = QWidget()
        scroll_content.setStyleSheet("background: transparent;")
        
        # Create layout for the scroll content
        buttons_layout = QVBoxLayout(scroll_content)
        buttons_layout.setSpacing(15)  # Increased spacing for better separation
        
        # Add button groups with enhanced UI
        self.add_button_group(buttons_layout, "الأحد", [
            ("الجناز العام", "Data\\اسبوع الالام\\تجنيز احد الشعانين 2022.pptx"),
            ("ليلة الإثنين", "Data\\اسبوع الالام\\ليلة الاثنين.pptx"),
        ])

        self.add_button_group(buttons_layout, "الإثنين", [
            ("يوم الإثنين", "Data\\اسبوع الالام\\يوم الاثنين.pptx"),
            ("ليلة الثلاثاء", "Data\\اسبوع الالام\\ليله الثلاثاء.pptx"),
        ])

        self.add_button_group(buttons_layout, "الثلاثاء", [
            ("يوم الثلاثاء", "Data\\اسبوع الالام\\يوم الثلاثاء.pptx"),
            ("ليلة الاربعاء", "Data\\اسبوع الالام\\ليلة الاربع.pptx"),
        ])

        self.add_button_group(buttons_layout, "الأربعاء", [
            ("يوم الاربعاء", "Data\\اسبوع الالام\\يوم الاربع.pptx"),
            ("ليلة الخميس", "Data\\اسبوع الالام\\ليلة الخميس.pptx"),
        ])

        self.add_button_group(buttons_layout, "الخميس", [
            ("خميس العهد", "Data\\اسبوع الالام\\خميس العهد.pptx"),
            ("ليلة الجمعة العطيمة", "Data\\اسبوع الالام\\ليلة الجمعة.pptx"),
        ])

        self.add_button_group(buttons_layout, "الجمعة", [
            ("الجمعة العطيمة", "Data\\اسبوع الالام\\الجمعة العظيمة كاملة 2022.pptx"),
        ])

        self.add_button_group(buttons_layout, "السبت", [
            ("ليلة أبو غلامسيس", "Data\\اسبوع الالام\\سبت النور.pptx"),
        ])
        
        # Set the scroll content to the scroll area
        scroll_area.setWidget(scroll_content)
        buttons_frame_layout.addWidget(scroll_area)
        
        return buttons_frame

    def add_button_group(self, layout, day, buttons):
        # Create a container for the day section
        day_container = QFrame()
        day_container.setStyleSheet("""
            QFrame {
                background-color: rgba(80, 10, 10, 100);
                border-radius: 10px;
                padding: 5px;
            }
        """)
        
        day_layout = QVBoxLayout(day_container)
        day_layout.setContentsMargins(10, 10, 10, 15)
        day_layout.setSpacing(8)
        
        # Create and add label for the day with icon
        day_label_container = QHBoxLayout()
        
        # Try to add icon before day label
        try:
            day_icon = QLabel()
            # Choose different icons based on day of the week
            icon_name = "fa5s.calendar-day"
            if day == "الأحد":
                icon_name = "fa5s.church"
            elif day == "الخميس":
                icon_name = "fa5s.wine-glass"
            elif day == "الجمعة":
                icon_name = "fa5s.cross"
            elif day == "السبت":
                icon_name = "fa5s.menorah"
            
            icon_pixmap = qta.icon(icon_name, color="white").pixmap(16, 16)
            day_icon.setPixmap(icon_pixmap)
            day_icon.setStyleSheet("background: transparent;")
            day_label_container.addWidget(day_icon)
            day_label_container.addSpacing(8)
        except:
            pass
        
        day_label = QLabel(day)
        day_label.setStyleSheet("""
            QLabel {
                color: white;
                font-size: 16px;
                font-weight: bold;
                background-color: transparent;
            }
        """)
        day_label_container.addWidget(day_label)
        day_label_container.addStretch()
        
        day_layout.addLayout(day_label_container)
        
        # Add a subtle separator line
        separator = QFrame()
        separator.setFrameShape(QFrame.HLine)
        separator.setFrameShadow(QFrame.Sunken)
        separator.setStyleSheet("background-color: rgba(255, 255, 255, 70); margin: 2px 0;")
        separator.setMaximumHeight(1)
        day_layout.addWidget(separator)
        
        # Add buttons for this day
        for button_text, path in buttons:
            button = QPushButton(button_text)
            button.setCursor(Qt.PointingHandCursor)
            button.setStyleSheet("""
                QPushButton {
                    background-color: rgba(255, 255, 255, 200);
                    border: none;
                    border-radius: 12px;
                    color: #3a0000;
                    padding: 10px;
                    font-size: 13px;
                    font-weight: bold;
                    text-align: center;
                    min-height: 20px;
                }
                QPushButton:hover {
                    background-color: rgba(255, 240, 240, 230);
                    color: #690000;
                    border: 1px solid rgba(255, 255, 255, 50);
                }
                QPushButton:pressed {
                    background-color: rgba(200, 180, 180, 250);
                    padding-top: 11px;
                    padding-bottom: 9px;
                }
            """)
            
            # Add shadow effect to button
            shadow = QGraphicsDropShadowEffect()
            shadow.setBlurRadius(10)
            shadow.setColor(QColor(0, 0, 0, 80))
            shadow.setOffset(2, 2)
            button.setGraphicsEffect(shadow)
            
            button.clicked.connect(lambda checked=False, p=path: open_presentation_relative_path(p))
            day_layout.addWidget(button)
        
        layout.addWidget(day_container)
    
    def mousePressEvent(self, event):
        # Allow dragging the frameless window from the header area
        if event.button() == Qt.LeftButton and event.y() < 50:  # 50 is header height
            self._drag_pos = event.globalPos() - self.frameGeometry().topLeft()
            event.accept()

    def mouseMoveEvent(self, event):
        # Move the window with mouse
        if event.buttons() == Qt.LeftButton and hasattr(self, '_drag_pos'):
            self.move(event.globalPos() - self._drag_pos)
            event.accept()