import asyncio
from PyQt5.QtWidgets import QMainWindow, QLabel, QPushButton, QVBoxLayout, QHBoxLayout, QFrame, QScrollArea, QWidget, QMessageBox
from PyQt5.QtGui import QPixmap, QIcon
from PyQt5.QtCore import Qt
from commonFunctions import relative_path, load_background_image, open_presentation_relative_path
from NotificationBar import NotificationBar

class elbas5aWindow(QMainWindow):
    def __init__(self, main_window):
        super().__init__()

        self.main_window = main_window  # Reference to the main window

        self.setWindowTitle("St. Mary Maadi Liturgies")
        self.setWindowIcon(QIcon(relative_path(r"Data\الصور\Logo.ico")))
        self.setGeometry(400, 100, 625, 600)
        self.setFixedSize(625, 600)

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

        # Add back button
        back_button = QPushButton("Back", self)
        back_button.setGeometry(button_x, button_y, button_width, button_height)
        back_button.clicked.connect(self.go_back)
        back_button.setText("⬅ العودة")
        back_button.setStyleSheet("""
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
        layout.addWidget(back_button, alignment=Qt.AlignBottom | Qt.AlignRight)

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

        frame = QFrame(self)
        frame.setGeometry(20, 90, 585, 450)
        frame.setStyleSheet("QFrame { background-color: rgba(107, 6, 6, 200); border: 2px solid black; }")

        layout = QHBoxLayout(frame)

        # Add a stretch to position the line where you want it
        layout.addStretch(8)

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
        pixmap = QPixmap(relative_path(r"Data\الصور\esbo3elalam.png"))  # Replace with your image path
        image_label.setPixmap(pixmap)
        image_label.setGeometry(20, 60, 230, 335)  # Adjust dimensions as needed
        image_label.setScaledContents(True)
        image_label.setStyleSheet("background-color: transparent;border: none;")  # Set transparent background

        # Create a nested layout for buttons
        self.buttons_layout = QVBoxLayout()

        # Run the async function to add buttons
        asyncio.run(self.add_buttons_async())

        # Add a scroll area for buttons
        scroll_area = QScrollArea()
        scroll_area.setStyleSheet("background-color: transparent; border: none; color: white;")
        scroll_area.setWidgetResizable(True)
        scroll_area.setMinimumWidth(170)
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

    async def add_buttons_async(self):
        await self.add_button_group("الأحد", [
            ("الجناز العام", "Data\\اسبوع الالام\\تجنيز احد الشعانين 2022.pptx"),
            ("ليلة الإثنين", "Data\\اسبوع الالام\\ليلة الاثنين.pptx"),
        ])

        await self.add_button_group("الإثنين", [
            ("يوم الإثنين", "Data\\اسبوع الالام\\يوم الاثنين.pptx"),
            ("ليلة الثلاثاء", "Data\\اسبوع الالام\\ليله الثلاثاء.pptx"),
        ])

        await self.add_button_group("الثلاثاء", [
            ("يوم الثلاثاء", "Data\\اسبوع الالام\\يوم الثلاثاء.pptx"),
            ("ليلة الاربعاء", "Data\\اسبوع الالام\\ليلة الاربع.pptx"),
        ])

        await self.add_button_group("الأربعاء", [
            ("يوم الاربعاء", "Data\\اسبوع الالام\\يوم الاربع.pptx"),
            ("ليلة الخميس", "Data\\اسبوع الالام\\ليلة الخميس.pptx"),
        ])

        await self.add_button_group("الخميس", [
            ("خميس العهد", "Data\\اسبوع الالام\\خميس العهد.pptx"),
            ("ليلة الجمعة العطيمة", "Data\\اسبوع الالام\\ليلة الجمعة.pptx"),
        ])

        await self.add_button_group("الجمعة", [
            ("الجمعة العطيمة", "Data\\اسبوع الالام\\الجمعة العظيمة كاملة 2022.pptx"),
        ])

        await self.add_button_group("السبت", [
            ("ليلة أبو غلامسيس", "Data\\اسبوع الالام\\سبت النور.pptx"),
        ])

    async def add_button_group(self, day, buttons):
        label = QLabel(day)
        label.setStyleSheet("background-color: transparent; border: none; color: white; font-size: 16px;")
        self.buttons_layout.addWidget(label)

        for button_text, pptx_path in buttons:
            button = QPushButton(button_text)
            button.clicked.connect(lambda _, p=pptx_path: open_presentation_relative_path(p))
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
            "   font-size: 12px;"
            "   font-weight: bold;"
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
