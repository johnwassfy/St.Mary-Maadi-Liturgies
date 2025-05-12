import os
from PyQt5.QtWidgets import QApplication, QMainWindow, QLabel, QPushButton, QVBoxLayout, QHBoxLayout, QFrame, QScrollArea, QWidget, QMessageBox
from PyQt5.QtGui import QPixmap
from PyQt5.QtCore import Qt
import win32com
from tasbha import *
from commonFunctions import relative_path, load_background_image

class tasbhawindow(QMainWindow):
    def __init__(self, copticdate,season):
        super().__init__()

        # self.main_window = main_window  # Reference to the main window

        self.setWindowTitle("Coptic Shasha")
        self.setGeometry(100, 100, 625, 600)
        self.setFixedSize(625, 600)

        # Create a central widget
        self.central_widget = QLabel(self)
        self.central_widget.setAlignment(Qt.AlignCenter)
        self.setCentralWidget(self.central_widget)

        # Create a vertical layout for the central widget
        layout = QVBoxLayout(self.central_widget)

        # Add back button
        self.back_button = QPushButton("Back")
        layout.addWidget(self.back_button, alignment=Qt.AlignBottom | Qt.AlignRight)  # Align the button to the top left corner
        self.back_button.clicked.connect(self.go_back)

        # Load background image
        try:
            load_background_image(self.central_widget)
        except Exception as e:
            self.show_error(f"خطأ في تحميل الخلفية: {str(e)}")

        frame0 = QFrame(self)
        frame0.setGeometry(0, 0, 625, 70)
        frame0.setStyleSheet("background-color: #ffffff;")
        # Add the picture to frame0
        image_label = QLabel(frame0)
        image_label.setGeometry(0, 0, 625, 70)
        image_path = relative_path(r"Data\الصور\Untitled-2.png")
        pixmap = QPixmap(image_path)
        image_label.setPixmap(pixmap)
        image_label.setScaledContents(True)

        frame = QFrame(self)
        frame.setGeometry(20, 90, 585, 450)
        frame.setStyleSheet("QFrame { background-color: rgba(229, 182, 102, 200); border: 2px solid black; }")

        layout = QHBoxLayout(frame)

        # Add a stretch to position the line where you want it
        layout.addStretch(7)

        # Add a line to divide the frame
        line = QFrame(frame)
        line.setFrameShape(QFrame.VLine)  # Set vertical line shape
        line.setFrameShadow(QFrame.Sunken)  # Set shadow style
        line.setStyleSheet("background-color: black;")  # Set line color
        layout.addWidget(line)

        # Add another stretch to fill remaining space
        layout.addStretch(2)

        # Add photo inside the first frame
        image_label = QLabel(frame)
        pixmap = QPixmap(relative_path(r"Data\الصور\داود.png"))  # Replace with your image path
        image_label.setPixmap(pixmap)
        image_label.setGeometry(25, 35, 247, 368)  # Adjust dimensions as needed
        image_label.setScaledContents(True)
        image_label.setStyleSheet("background-color: transparent;border: none;")  # Set transparent background

        # Create a nested layout for buttons
        self.buttons_layout = QVBoxLayout()
        
        button = QPushButton("تسبحة نصف الليل")
        button.setGeometry(0, 0, 30, 20)
        # if season == 5 :
        #     button.clicked.connect(lambda: kiahk(copticdate))
        # else:
        #     button.clicked.connect(lambda: tasbha(copticdate, False, season))
        self.set_default_button_style(button)
        self.buttons_layout.addWidget(button)

        button = QPushButton("تسبحة عشية")
        button.setGeometry(0, 0, 30, 20)
        # button.clicked.connect(lambda: tasbha(copticdate, True, season))
        self.set_default_button_style(button)
        self.buttons_layout.addWidget(button)

        # Add a scroll area for buttons
        scroll_area = QScrollArea()
        scroll_area.setStyleSheet("background-color: transparent; border: none; color: white;")
        scroll_area.setWidgetResizable(False)
        scroll_area.setMinimumWidth(50)
        scroll_content = QWidget()
        scroll_content.setLayout(self.buttons_layout)
        scroll_area.setWidget(scroll_content)

        # Set stylesheet for scrollbar to make it transparent
        scroll_area.verticalScrollBar().setStyleSheet(
            "QScrollBar:vertical {border: none; background: transparent; width: 10px;}"
            "QScrollBar::handle:vertical {background: rgba(255, 255, 255, 100); border-radius: 5px;}"
            "QScrollBar::add-line:vertical {background: none;}"
            "QScrollBar::sub-line:vertical {background: none;}"
            "QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical {background: none;}"
        )

        # Add the scroll area to the main layout
        layout.addWidget(scroll_area)

    def set_default_button_style(self, button):
        button.setStyleSheet(
            "QPushButton {"
            "   background-color: #f0f0f0;"
            "   border: 1px solid #c4c4c4;"
            "   border-radius: 5px;"
            "   color: #333333;"
            "   padding: 5px 10px;"
            "   font-size: 20px;"
            "}"
            "QPushButton:hover {"
            "   background-color: #e0e0e0;"
            "}"
            "QPushButton:pressed {"
            "   background-color: #d9d9d9;"
            "}"
        )

    def go_back(self):
        self.close()

    def show_error_message(self, error_message):
        QMessageBox.critical(self, "Error", error_message)

