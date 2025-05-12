import os
from PyQt5.QtWidgets import QApplication, QMainWindow, QLabel, QPushButton, QVBoxLayout, QHBoxLayout, QFrame, QScrollArea, QWidget, QMessageBox
from PyQt5.QtGui import QPixmap
from PyQt5.QtCore import Qt
import win32com
from commonFunctions import relative_path, load_background_image
from NotificationBar import NotificationBar 

class ellakanwindow(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("St. Mary Maadi Liturgies")
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

        # Add NotificationBar
        self.notification_bar = NotificationBar(self)
        self.notification_bar.setGeometry(0, 70, self.width(), 50)
        
        # Load background image
        try:
            load_background_image(self.central_widget)
        except Exception as e:
            self.notification_bar.show_message(f"خطأ في تحميل الخلفية: {str(e)}")

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
        frame.setStyleSheet("QFrame { background-color: rgba(204, 178, 119, 200); border: 2px solid black; }")

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
        layout.addStretch(4)

        # Add photo inside the first frame
        image_label = QLabel(frame)
        pixmap = QPixmap(relative_path(r"Data\الصور\اللقان 2.png"))  # Replace with your image path
        image_label.setPixmap(pixmap)
        image_label.setGeometry(30, 20, 210, 411)  # Adjust dimensions as needed
        image_label.setScaledContents(True)
        image_label.setStyleSheet("background-color: transparent;border: none;")  # Set transparent background

        # Create a nested layout for buttons
        self.buttons_layout = QVBoxLayout()
        
        self.add_button_group ([("خميس العهد", r"Data\اسبوع الالام\خميس العهد.pptx"), 
                                ("عيد الرسل", r"Data\لقان عيد الرسل.pptx")])

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

    def add_button_group(self, buttons):
        for button_text, pptx_path in buttons:
            button = QPushButton(button_text)
            button.setGeometry(0, 0, 30, 20)
            button.clicked.connect(lambda _, p=pptx_path: self.open_presentation(p))
            self.set_default_button_style(button)
            self.buttons_layout.addWidget(button)

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

    def open_presentation(self, file_name, slide_number=None):
        file_path = relative_path(file_name)
        if slide_number is None:
            os.startfile(file_path)
        else:
            self.open_presentation_on_slide(file_path, slide_number)
    
    def open_presentation_on_slide(self, presentation_path, slide_number):
        # Create a PowerPoint application object
        powerpoint = win32com.client.Dispatch("PowerPoint.Application")
        
        # Open the presentation
        presentation = powerpoint.Presentations.Open(presentation_path, WithWindow=True)
        
        # Make PowerPoint visible
        powerpoint.Visible = True
        
        # Navigate to the specified slide
        slide = presentation.Slides(slide_number)
        slide.Select()

    def go_back(self):
        self.close()

    def show_error_message(self, error_message):
        QMessageBox.critical(self, "Error", error_message)

# from sys import argv, exit
# app = QApplication(argv)
# window = ellakanwindow()
# window.show()
# exit(app.exec_())
