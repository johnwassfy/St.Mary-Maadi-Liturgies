from PyQt5.QtWidgets import (QDialog, QPushButton, QVBoxLayout, QLabel, QFrame, QHBoxLayout)
from PyQt5.QtGui import QFont, QPixmap
from PyQt5.QtCore import Qt, QSize, QRect
from commonFunctions import relative_path
import qtawesome as qta

class HoverButton(QPushButton):
    def __init__(self, text, parent=None):
        super(HoverButton, self).__init__(text, parent)
        self.setMouseTracking(True)
        self._animation_value = 0
        self.setMinimumHeight(90)  # Increased minimum height

    def enterEvent(self, event):
        # Preserve the icon when setting new style
        current_icon = self.icon()
        current_icon_size = self.iconSize()
        
        self.setStyleSheet("""
            QPushButton {
                background-color: rgba(35, 107, 142, 220);
                border: 1px solid rgba(35, 107, 142, 250);
                border-radius: 15px;
                color: white;
                padding: 10px;
                font-size: 18px;
                text-align: center;
            }
        """)
        
        # Reapply icon
        if not current_icon.isNull():
            self.setIcon(current_icon)
            self.setIconSize(current_icon_size)
        
        super().enterEvent(event)

    def leaveEvent(self, event):
        # Preserve the icon when setting new style
        current_icon = self.icon()
        current_icon_size = self.iconSize()
        
        self.setStyleSheet("""
            QPushButton {
                background-color: rgba(240, 240, 240, 200);
                border: 1px solid #c4c4c4;
                border-radius: 15px;
                color: #333333;
                padding: 10px;
                font-size: 18px;
                text-align: center;
            }
        """)
        
        # Reapply icon
        if not current_icon.isNull():
            self.setIcon(current_icon)
            self.setIconSize(current_icon_size)
            
        super().leaveEvent(event)
        
class LakanSelectionDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.selected_option = None
        
        self.setWindowTitle("صلاة اللقان")
        self.setFixedSize(480, 400)
        
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
                    stop: 0 rgba(15, 46, 71, 245),
                    stop: 0.6 rgba(30, 91, 138, 245),
                    stop: 1 rgba(140, 217, 255, 245)
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
                
        # Buttons panel on the left
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
                    stop: 0 #1e5b8a,
                    stop: 1 #3498db
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
            icon = qta.icon("fa5s.book-open", color="white").pixmap(24, 24)
            icon_label.setPixmap(icon)
            icon_label.setStyleSheet("background: transparent;")
            title_layout.addWidget(icon_label)
            title_layout.addSpacing(10)
        except:
            pass
        
        # Title
        title_label = QLabel("الإبصلمودية المقدسة")
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
        photo_layout.setContentsMargins(0,0,0,0)  # Reduced margins to maximize photo size
        photo_layout.setSpacing(0)  # No spacing to keep photo centered
        # Add the photo with larger dimensions
        try:
            photo_label = QLabel()
            pixmap = QPixmap(relative_path(r"Data\الصور\اللقان 2.png"))
            
            # Make photo larger while maintaining aspect ratio
            pixmap = pixmap.scaled(220, 320, Qt.KeepAspectRatio, Qt.SmoothTransformation)
            
            photo_label.setPixmap(pixmap)
            photo_label.setAlignment(Qt.AlignCenter)
            photo_label.setStyleSheet("background: transparent; border: none;")
            photo_layout.addWidget(photo_label, 1, alignment=Qt.AlignCenter)
        except Exception as e:
            # Fallback text if image fails to load
            fallback = QLabel("أيقونة لقان خميس العهد")
            fallback.setAlignment(Qt.AlignCenter)
            fallback.setStyleSheet("color: white; font-size: 14px; background: transparent;")
            photo_layout.addWidget(fallback)
    
        return photo_frame

    def create_buttons_panel(self):
        buttons_frame = QFrame()        
        buttons_layout = QVBoxLayout(buttons_frame)
        buttons_layout.setContentsMargins(10, 10, 10, 10)  # Add margins for spacing
        buttons_layout.setSpacing(15)  # Space between buttons

        # Midnight button with centered content
        button = HoverButton("عيد الغطاس")
        button.setMinimumHeight(90)
        button.setCursor(Qt.PointingHandCursor)
        button.setStyleSheet("""
            QPushButton {
                background-color: rgba(240, 240, 240, 200);
                border: 1px solid #c4c4c4;
                border-radius: 15px;
                color: #333333;
                padding: 10px;
                font-size: 18px;
            }
        """)
        
        # Style and set icon
        self.style_button(button)
        
        # Add custom CSS for center alignment with icon
        button.setStyleSheet(button.styleSheet() + """
            QPushButton {
                text-align: center;
            }
        """)
        
        button.clicked.connect(lambda: self.button_clicked("Baptism"))
        
        buttons_layout.addWidget(button)

        # Evening button with centered content
        button = HoverButton("خميس العهد")
        button.setMinimumHeight(90)
        button.setCursor(Qt.PointingHandCursor)
        button.setStyleSheet("""
            QPushButton {
                background-color: rgba(240, 240, 240, 200);
                border: 1px solid #c4c4c4;
                border-radius: 15px;
                color: #333333;
                padding: 10px;
                font-size: 18px;
            }
        """)
        
        # Style and set icon
        self.style_button(button)
        
        # Add custom CSS for center alignment with icon
        button.setStyleSheet(button.styleSheet() + """
            QPushButton {
                text-align: center;
            }
        """)
        
        button.clicked.connect(lambda: self.button_clicked("Holy Thursday"))
        
        buttons_layout.addWidget(button)

        # Evening button with centered content
        button = HoverButton("عيد الرسل")
        button.setMinimumHeight(90)
        button.setCursor(Qt.PointingHandCursor)
        button.setStyleSheet("""
            QPushButton {
                background-color: rgba(240, 240, 240, 200);
                border: 1px solid #c4c4c4;
                border-radius: 15px;
                color: #333333;
                padding: 10px;
                font-size: 18px;
            }
        """)
        
        # Style and set icon
        self.style_button(button)
        
        # Add custom CSS for center alignment with icon
        button.setStyleSheet(button.styleSheet() + """
            QPushButton {
                text-align: center;
            }
        """)
        
        button.clicked.connect(lambda: self.button_clicked("Apostles"))
        
        buttons_layout.addWidget(button)
        
        return buttons_frame    
    
    def style_button(self, button, icon_name=None):
        button_font = QFont()
        button_font.setPointSize(18)
        button_font.setBold(True)
        button.setFont(button_font)
        
        # Add icon if specified
        if icon_name:
            try:
                icon = qta.icon(icon_name, color='#1e5b8a')
                button.setIcon(icon)
                button.setIconSize(QSize(32, 32))
                
                # Center icon and text without using setToolButtonStyle
                button.setStyleSheet(button.styleSheet() + """
                    QPushButton {
                        text-align: center;
                        padding: 10px 10px 10px 10px;
                    }
                """)
                
                # Set proper alignment
                button.setLayoutDirection(Qt.RightToLeft)  # RTL for Arabic text
            except Exception as e:
                print(f"Icon error: {e}")
            
    def button_clicked(self, option):
        self.selected_option = option
        self.accept()

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