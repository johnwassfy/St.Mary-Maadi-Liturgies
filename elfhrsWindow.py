import os
from PyQt5.QtWidgets import QApplication, QMainWindow, QLabel, QPushButton, QVBoxLayout, QHBoxLayout, QFrame, QScrollArea, QWidget, QMessageBox, QComboBox
from PyQt5.QtGui import QPixmap
from PyQt5.QtCore import Qt
from UpdateTable import *
from commonFunctions import relative_path

class elfhrswindow(QMainWindow):
    def __init__(self):
        super().__init__()

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
        layout.addWidget(self.back_button, alignment=Qt.AlignBottom | Qt.AlignRight)
        self.back_button.clicked.connect(self.go_back)

        button = QPushButton("تحديث الملفات", self)
        button.setGeometry(self.width()- 200, 566, 115, 25)
        button.clicked.connect(lambda _: self.update_section_names())

        # Load background image
        self.load_background_image()

        frame0 = QFrame(self)
        frame0.setGeometry(0, 0, 625, 70)
        frame0.setStyleSheet("background-color: #ffffff;")
        # Add the picture to frame0
        image_label = QLabel(frame0)
        image_label.setGeometry(0, 0, 625, 70)
        image_path = relative_path(r"Data\الصور\Untitled-4.png")
        pixmap = QPixmap(image_path)
        image_label.setPixmap(pixmap)
        image_label.setScaledContents(True)

        frame = QFrame(self)
        frame.setGeometry(20, 90, 585, 450)
        frame.setStyleSheet("QFrame { background-color: rgba(229, 182, 102, 200); border: 2px solid black; }")

        layout = QHBoxLayout(frame)

        # Add a stretch to position the line where you want it
        layout.addStretch(8)

        # Add a line to divide the frame
        line = QFrame(frame)
        line.setFrameShape(QFrame.VLine)
        line.setFrameShadow(QFrame.Sunken)
        line.setStyleSheet("background-color: black;")
        layout.addWidget(line)

        # Add another stretch to fill remaining space
        layout.addStretch(1)

        # Add photo inside the first frame
        image_label = QLabel(frame)
        pixmap = QPixmap(relative_path(r"Data\الصور\الفهرس.png"))
        image_label.setPixmap(pixmap)
        image_label.setGeometry(18, 35, 247, 368)
        image_label.setScaledContents(True)
        image_label.setStyleSheet("background-color: transparent;border: none;")

        # Create a nested layout for buttons and dropdowns
        self.buttons_layout = QVBoxLayout()

        # Add the buttons
        self.add_buttons(["ملف باكر و عشية", "ملف القداس", "ملف قداس الطفل", 
                          "الإبصلمودية السنوية", "الإبصلمودية الكيهكية", "الذكصولوجيات"],
                         [r"Data\CopyData\رفع بخور عشية و باكر.pptx", r"Data\CopyData\قداس.pptx", 
                          r"Data\CopyData\قداس الطفل.pptx", r"Data\CopyData\الإبصلمودية.pptx", 
                          r"Data\CopyData\الإبصلمودية الكيهكية.pptx", r"Data\CopyData\الذكصولوجيات.pptx"])

        # Create a label with the text "القطمارس"
        self.label = QLabel("القطمارس", self)
        self.label.setStyleSheet("font-size: 20px; color: black; font-weight: bold;")  # Set the font size and color
        self.label.setAlignment(Qt.AlignCenter)  # Center align the text

        # Add the label to the vertical layout before the drop-downs
        self.buttons_layout.addWidget(self.label, alignment=Qt.AlignRight)

        # Create a horizontal layout for the drop-downs
        dropdowns_layout = QHBoxLayout()

        # Create the first drop-down menu
        self.first_dropdown = QComboBox(self)
        self.first_dropdown.addItems(["القداس", "العشية", "باكر"])
        self.first_dropdown.setFixedWidth(80)  # Set a fixed width to fit both dropdowns side by side
        self.first_dropdown.setStyleSheet("""
            QComboBox {
                background-color: white;
                color: black;
                font-size: 18px;
                padding: 2px;
                border: 1px solid gray;
            }
            QComboBox::drop-down {
                width: 15px;
            }
            QComboBox QAbstractItemView {
                background-color: white;
                color: black;
                selection-background-color: lightblue;
            }
        """)
        self.first_dropdown.setLayoutDirection(Qt.RightToLeft)

        # Create the second drop-down menu
        self.second_dropdown = QComboBox(self)
        self.second_dropdown.addItems(["سنوي ايام", "سنوي آحاد", "الصوم الكبير", "الخماسين"])
        self.second_dropdown.setFixedWidth(110)  # Set a fixed width to fit both dropdowns side by side
        self.second_dropdown.setStyleSheet("""
            QComboBox {
                background-color: white;
                color: black;
                font-size: 18px;
                padding: 2px;
                border: 1px solid gray;
            }
            QComboBox::drop-down {
                width: 15px;
            }
            QComboBox QAbstractItemView {
                background-color: white;
                color: black;
                selection-background-color: lightblue;
            }
        """)
        self.second_dropdown.setLayoutDirection(Qt.RightToLeft)

        # Add both drop-downs to the horizontal layout
        dropdowns_layout.addWidget(self.first_dropdown)
        dropdowns_layout.addWidget(self.second_dropdown)

        # Add the horizontal layout to the vertical layout of buttons and dropdowns
        self.buttons_layout.addLayout(dropdowns_layout)

        self.combination_to_file = {
            ("القداس", "سنوي ايام"): r"Data\القطمارس\الايام\القطمارس السنوي ايام (القداس).pptx",
            ("القداس", "الخماسين"): r"Data\القطمارس\قطمارس الخماسين (القداس).pptx",
            ("القداس", "الصوم الكبير"): r"Data\القطمارس\قطمارس الصوم الكبير (القداس).pptx",
            ("القداس", "سنوي آحاد"): r"Data\القطمارس\الاحاد\القطمارس السنوي احاد (القداس).pptx",
            ("العشية", "سنوي ايام"): r"Data\القطمارس\الايام\القطمارس السنوي ايام (عشية).pptx",
            ("العشية", "الخماسين"): None,
            ("العشية", "الصوم الكبير"): None,
            ("العشية", "سنوي آحاد"): r"Data\القطمارس\الاحاد\القطمارس السنوي احاد (عشية).pptx",
            ("باكر", "سنوي ايام"): r"Data\القطمارس\الايام\القطمارس السنوي ايام (باكر).pptx",
            ("باكر", "الخماسين"): None,
            ("باكر", "الصوم الكبير"): None,
            ("باكر", "سنوي آحاد"): None,
        }

        # Create and add the button below the drop-downs
        self.open_button = QPushButton("فتح القطمارس المختار", self)
        self.open_button.clicked.connect(self.handle_open_button)
        self.set_default_button_style(self.open_button)
        self.buttons_layout.addWidget(self.open_button)

        # Create and add the button below the drop-downs
        self.update_button = QPushButton("تحديث القطمارس المختار", self)
        self.update_button.clicked.connect(self.handle_update_button)
        self.set_default_button_style(self.update_button)
        self.buttons_layout.addWidget(self.update_button)

        # Add a scroll area for buttons and dropdowns
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

    def handle_open_button(self):
        # Get selected items from the drop-downs
        first_selection = self.first_dropdown.currentText()
        second_selection = self.second_dropdown.currentText()

        # Find the corresponding file path
        file_path = self.combination_to_file.get((first_selection, second_selection))

        if file_path:
            # If a file path exists, open the presentation
            self.open_presentation(file_path)
        else:
            # If no file path exists, show an error message
            self.show_error_message(f"هذا الملف غير متوفر الان.")

    def handle_update_button(self):
        first_selection = self.first_dropdown.currentText()
        second_selection = self.second_dropdown.currentText()

        try:
            if first_selection == "القداس" and second_selection == "سنوي ايام":
                katamarsOdasElsanawyAyam()
                self.show_message("تم التحديث بنجاح")
            elif first_selection == "القداس" and second_selection == "الخماسين":
                katamarsEl5amasyn()
                self.show_message("تم التحديث بنجاح")
            elif first_selection == "القداس" and second_selection == "الصوم الكبير":
                ElsomElkbyr()
                self.show_message("تم التحديث بنجاح")
            elif first_selection == "القداس" and second_selection == "الصوم الكبير":
                ElsomElkbyr()
                self.show_message("تم التحديث بنجاح")
            elif first_selection == "باكر" and second_selection == "الصوم الكبير":
                ElsomElkbyr()
                self.show_message("تم التحديث بنجاح")
            elif first_selection == "عشية" and second_selection == "سنوي آحاد":
                katamarsOdasElsanawyA7ad()
                self.show_message("تم التحديث بنجاح")
            elif first_selection == "العشية" and second_selection == "سنوي ايام":
                katamars3ashyaElsanawyAyam()
                self.show_message("تم التحديث بنجاح")
            elif first_selection == "العشية" and second_selection == "سنوي آحاد":
                katamars3ashyaElsanawyA7ad()
                self.show_message("تم التحديث بنجاح")
            elif first_selection == "باكر" and second_selection == "سنوي ايام":
                katamarsBakerElsanawyAyam()
                self.show_message("تم التحديث بنجاح")
            else:
                self.show_error_message("No update method is available for the selected combination.")
        except Exception as e:
            self.show_error_message(f"An error occurred: {str(e)}")

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

    def add_buttons(self, button_names, paths):
        for index, button in enumerate(button_names):
            button_name = QPushButton(button)
            button_name.clicked.connect(lambda _, p=paths[index]: self.open_presentation(p))
            self.set_default_button_style(button_name)
            self.buttons_layout.addWidget(button_name)

    def show_error_message(self, error_message):
        QMessageBox.critical(self, "Error", error_message)

    def load_background_image(self):
        pixmap = QPixmap(relative_path(r"Data\الصور\background.png"))
        self.central_widget.setPixmap(pixmap)
        self.central_widget.setScaledContents(True)

    def open_presentation(self, file_name):
        file_path = relative_path(file_name)
        os.startfile(file_path)

    def update_section_names(self):
        from sectionNames import extract_section_info2
        try:
            file_sheet_pairs = [
                (relative_path(r"Data\CopyData\قداس.pptx"), "القداس"),
                (relative_path(r"Data\CopyData\قداس الطفل.pptx"), "قداس الطفل"),
                (relative_path(r"Data\CopyData\رفع بخور عشية و باكر.pptx"), "رفع بخور"),
                (relative_path(r"Data\CopyData\الذكصولوجيات.pptx"), "الذكصولوجيات"),
                (relative_path(r"Data\CopyData\في حضور الاسقف و اساقفة ضيوف.pptx"), "في حضور الأسقف"),
                (relative_path(r"Data\CopyData\الإبصلمودية.pptx"), "التسبحة"),
                (relative_path(r"Data\CopyData\الإبصلمودية الكيهكية.pptx"), "تسبحة كيهك"),
            ]

            excel_file = relative_path(r'Files Data.xlsx')
            
            extract_section_info2(file_sheet_pairs, excel_file)

            self.replace_presentation()

            # Show success message
            self.show_message("تم التحديث بنجاح!")

        except Exception as e:
            self.show_error_message(str(e))

    def show_message(self, message):
        QMessageBox.information(self, "Message", message)

    def replace_presentation(self):
        from shutil import copy2
        from os import path, remove
        
        # Default to all False if no flags are passed
        replace_flags = {
                'odasEltfl': True, 
                'bakerWaashya': True, 
                'tasbha': True, 
                'tasbhaKiahk': True,
                'zoksologyat': True,
                'default': True
        }
        
        # Define the presentations and their paths
        presentations = {
            'odasEltfl': (r"قداس الطفل.pptx", r"Data\CopyData\قداس الطفل.pptx"),
            'bakerWaashya': (r"رفع بخور عشية و باكر.pptx", r"Data\CopyData\رفع بخور عشية و باكر.pptx"),
            'tasbha': (r"الإبصلمودية.pptx", r"Data\CopyData\الإبصلمودية.pptx"),
            'tasbhaKiahk': (r"الإبصلمودية الكيهكية.pptx", r"Data\CopyData\الإبصلمودية الكيهكية.pptx"),
            'zoksologyat': (r"الذكصولوجيات.pptx", r"Data\CopyData\الذكصولوجيات.pptx"),
            'default': (r"قداس.pptx", r"Data\CopyData\قداس.pptx")
        }
        
        # Loop through the flags and process each one
        for key, flag in replace_flags.items():
            if flag:
                old_presentation_path = relative_path(presentations[key][0])
                new_presentation_path = relative_path(presentations[key][1])

            try:
                # Check if the old presentation file exists
                if path.exists(old_presentation_path):
                    # If it exists, delete the old presentation
                    remove(old_presentation_path)
                    
                    # Copy the new presentation to the location of the old presentation
                    copy2(new_presentation_path, old_presentation_path)
            except Exception as e:
                # Print any errors that occur during the deletion and copying process
                print(f"Error: {str(e)}")


# from sys import argv, exit
# app = QApplication(argv)
# window = elfhrswindow()
# window.show()
# exit(app.exec_())
