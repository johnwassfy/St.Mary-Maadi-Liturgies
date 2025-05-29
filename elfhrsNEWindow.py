import os
import pandas as pd
from PyQt5.QtWidgets import QApplication, QMainWindow, QProgressBar, QSizePolicy, QLineEdit, QLabel, QPushButton, QVBoxLayout, QHBoxLayout, QFrame, QScrollArea, QWidget, QMessageBox, QComboBox
from PyQt5.QtGui import QPixmap
from PyQt5.QtCore import Qt
from commonFunctions import relative_path, load_background_image
from NotificationBar import NotificationBar
import win32com.client
from UpdateTable import All
from WorkerThread import WorkerThread

class elfhrswindow(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("St. Mary Maadi Liturgies")
        self.setGeometry(100, 100, 625, 600)
        self.setFixedSize(625, 600)

        # Create a central widget
        self.central_widget = QLabel(self)
        self.central_widget.setAlignment(Qt.AlignCenter)
        self.setCentralWidget(self.central_widget)

        # Add NotificationBar
        self.notification_bar = NotificationBar(self)
        self.notification_bar.setGeometry(0, 70, self.width(), 50)

        # Add a semi-transparent overlay (initially hidden)
        self.overlay = QLabel(self)
        self.overlay.setGeometry(0, 0, self.width(), self.height())  # Cover the entire window
        self.overlay.setStyleSheet("background-color: rgba(0, 0, 0, 150);")  # Semi-transparent black
        self.overlay.setVisible(False)  # Hide by default
        self.overlay.raise_()  # Ensure the overlay is on top of all widgets

        # Add a progress bar (centered in the window)
        self.progress_bar = QProgressBar(self)
        self.progress_bar.setFixedSize(400, 20)  # Set a fixed size for the progress bar
        self.progress_bar.move((self.width() - self.progress_bar.width()) // 2, 
                              (self.height() - self.progress_bar.height()) // 2)  # Center the progress bar
        self.progress_bar.setVisible(False)  # Hide by default
        self.progress_bar.setRange(0, 100)  # Set range to 0-100
        self.progress_bar.setStyleSheet("""
            QProgressBar {
                border: 2px solid #333;
                border-radius: 5px;
                background-color: #f0f0f0;
                text-align: center;
                color: #333;
                font-size: 14px;
            }
            QProgressBar::chunk {
                background-color: #4CAF50;  /* Green color for progress */
                border-radius: 5px;
            }
        """)
        self.progress_bar.raise_()  # Ensure the progress bar is above the overlay

        # Create a vertical layout for the central widget
        self.main_layout = QVBoxLayout(self.central_widget)

        # Add back button
        self.back_button = QPushButton("Back")
        self.main_layout.addWidget(self.back_button, alignment=Qt.AlignBottom | Qt.AlignRight)
        self.back_button.clicked.connect(self.go_back)

        button = QPushButton("تحديث ملفات الصلوات", self)
        button.setGeometry(self.width() - 200, 566, 115, 25)
        button.clicked.connect(lambda _: self.update_section_names())

        button = QPushButton("تحديث ملفات القطمارس", self)
        button.setGeometry(self.width() - 321, 566, 120, 25)
        button.clicked.connect(lambda _: self.update_katamars_files())

        # Load background image
        try:
            load_background_image(self.central_widget)
        except Exception as e:
            self.notification_bar.show_message(f"خطأ في تحميل الخلفية: {str(e)}")

        self.frame0 = QFrame(self)
        self.frame0.setGeometry(0, 0, 625, 70)
        self.frame0.setStyleSheet("background-color: #ffffff;")
        
        # Add the picture to frame0
        image_label = QLabel(self.frame0)
        image_label.setGeometry(0, 0, 625, 70)
        image_path = relative_path(r"Data\الصور\Untitled-2.png")
        pixmap = QPixmap(image_path)
        image_label.setPixmap(pixmap)
        image_label.setScaledContents(True)

        frame = QFrame(self)
        frame.setGeometry(20, 90, 585, 450)
        frame.setStyleSheet("QFrame { background-color: rgba(229, 182, 102, 200); border: 2px solid black; }")

        self.frame_layout = QHBoxLayout(frame)

        # Add a stretch to position the line where you want it
        self.frame_layout.addStretch(8)

        # Add a line to divide the frame
        line = QFrame(frame)
        line.setFrameShape(QFrame.VLine)
        line.setFrameShadow(QFrame.Sunken)
        line.setStyleSheet("background-color: black;")
        self.frame_layout.addWidget(line)

        # Add another stretch to fill remaining space
        self.frame_layout.addStretch(1)

        # Add photo inside the first frame
        self.image_label = QLabel(frame)
        pixmap = QPixmap(relative_path(r"Data\الصور\الفهرس.png"))
        self.image_label.setPixmap(pixmap)
        self.image_label.setGeometry(12, 25, 300, 400)  # Adjust the size to make it bigger
        self.image_label.setScaledContents(True)
        self.image_label.setAlignment(Qt.AlignCenter)  # Center align the image
        self.image_label.setStyleSheet("background-color: transparent;border: none;")
        
        # Create a nested layout for buttons and dropdowns
        self.buttons_layout = QVBoxLayout()
        self.buttons_layout.setContentsMargins(0, 0, 15, 0)  # Add right margin to the layout
        
        # Create a label with the text "الصلوات"
        self.label = QLabel("الصلوات", self)
        self.label.setStyleSheet("font-size: 20px; color: black; font-weight: bold;")  # Set the font size and color
        self.label.setAlignment(Qt.AlignCenter)  # Center align the text
        self.buttons_layout.addWidget(self.label, alignment=Qt.AlignRight)

        # Add the buttons
        self.add_buttons(["باكر و عشية", "القداس", "قداس الطفل", "الإبصلمودية السنوية", 
                          "الإبصلمودية الكيهكية", "الذكصولوجيات", "المدائح والتماجيد"], 
                         ["رفع بخور", "القداس", "قداس الطفل", "التسبحة", "تسبحة كيهك", "الذكصولوجيات", "المدائح"])

        # Create a label with the text "القطمارس"
        self.label = QLabel("القطمارس", self)
        self.label.setStyleSheet("font-size: 20px; color: black; font-weight: bold;")  # Set the font size and color
        self.label.setAlignment(Qt.AlignCenter)  # Center align the text

        # Add the label to the vertical layout before the drop-downs
        self.buttons_layout.addWidget(self.label, alignment=Qt.AlignRight)

        self.add_buttons_with_paths(
            ["الصوم الكبير", "الخماسين", "سنوي أيام - باكر", "سنوي أيام - عشية", "سنوي أيام - قداس",
             "سنوي آحاد - باكر", "سنوي آحاد - عشية", "سنوي آحاد - قداس"], 
            ["Data\القطمارس\الصوم الكبير و صوم نينوى\قطمارس الصوم الكبير.pptx", 
             "Data\القطمارس\قطمارس الخماسين (القداس).pptx", 
             "Data\القطمارس\الايام\القطمارس السنوي ايام (باكر).pptx", 
             "Data\القطمارس\الايام\القطمارس السنوي ايام (عشية).pptx", 
             "Data\القطمارس\الايام\القطمارس السنوي ايام (القداس).pptx",
             "Data\القطمارس\الاحاد\القطمارس السنوي احاد (باكر).pptx", 
             "Data\القطمارس\الاحاد\القطمارس السنوي احاد (عشية).pptx", 
             "Data\القطمارس\الاحاد\القطمارس السنوي احاد (القداس).pptx"
             ])

        # Add a scroll area for buttons and dropdowns
        scroll_area = QScrollArea()
        scroll_area.setStyleSheet("background-color: transparent; border: none; color: white;")
        scroll_area.setWidgetResizable(False)
        scroll_area.setMinimumWidth(50)
        scroll_area.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)  # Disable horizontal scroll bar
        scroll_content = QWidget()
        scroll_content.setLayout(self.buttons_layout)
        scroll_area.setWidget(scroll_content)

        # Set stylesheet for scrollbar to make it transparent
        scroll_area.verticalScrollBar().setStyleSheet(
            "QScrollBar:vertical {border: none; background: transparent; width: 10px;}"
            "QScrollBar::handle:vertical {background: rgba(255, 255, 255, 100); border-radius: 5px;}"
            "QScrollBar::add-line:vertical {background: none;}"
            "QScrollBar::sub-line:vertical {background: none;}"
            "QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical {background: rgba(0, 0, 0, 100);}"
        )

        # Create a container for the search box and checkbox
        container_widget = QWidget(self)
        container_widget.setGeometry(0, 554, 240, 40)  # Adjust position and size as needed
        
        # Add the scroll area to the main layout
        self.frame_layout.addWidget(scroll_area)

        # Ensure NotificationBar is on top
        self.notification_bar.raise_()

    def add_buttons(self, button_names, sheet_names):
        for index, button in enumerate(button_names):
            button_name = QPushButton(button)
            self.set_default_button_style(button_name)
            button_name.clicked.connect(lambda _, sheet=sheet_names[index], name=button: self.load_sheet_data(sheet, name))
            self.buttons_layout.addWidget(button_name)

    def add_buttons_with_paths(self, button_names, paths):
        for index, button in enumerate(button_names):
            button_name = QPushButton(button)
            button_name.clicked.connect(lambda _, p=paths[index]: self.katamars_button_click(p))
            self.set_default_button_style(button_name)
            self.buttons_layout.addWidget(button_name)

    def load_sheet_data(self, sheet_name, button_name):
        excel_path = relative_path(r"Files Data.xlsx")
        df = pd.read_excel(excel_path, sheet_name=sheet_name)
        column_data = df.iloc[:, 0].tolist()  # Get the first column data
        index = df.iloc[:, 2].tolist()  # Get the start index of every section

        # Create a dictionary mapping item names to their indices
        item_to_index = {item: idx for item, idx in zip(column_data, index)}

        # Create a new scroll area for the data
        self.create_scroll_area(column_data, button_name, item_to_index)

    def create_scroll_area(self, data, button_name, item_to_index):
        # Check if the scroll area already exists
        if hasattr(self, 'data_scroll_area') and self.data_scroll_area is not None:
            # Remove the existing scroll area and label
            self.frame_layout.removeWidget(self.scroll_area_container)
            self.scroll_area_container.deleteLater()
        else:
            # Remove the image if the scroll area is being created for the first time
            self.image_label.deleteLater()

        match(button_name):
            case "باكر و عشية":
                file_path = r"Data\CopyData\رفع بخور عشية و باكر.pptx"
            case "القداس":
                file_path = r"Data\CopyData\قداس.pptx"
            case "قداس الطفل":
                file_path = r"Data\CopyData\قداس الطفل.pptx"
            case "الإبصلمودية السنوية":
                file_path = r"Data\CopyData\الإبصلمودية.pptx"
            case "الإبصلمودية الكيهكية":
                file_path = r"Data\CopyData\الإبصلمودية الكيهكية.pptx"
            case "الذكصولوجيات":
                file_path = r"Data\CopyData\الذكصولوجيات.pptx"
            case "المدائح والتماجيد":
                file_path = r"Data\CopyData\كتاب المدائح.pptx"

        file_path = relative_path(file_path)

        # Create a container widget for the label and scroll area
        self.scroll_area_container = QWidget()
        container_layout = QVBoxLayout(self.scroll_area_container)

        # Create a horizontal layout for the label and search bar
        label_search_layout = QHBoxLayout()

        # Add a label with the button name on top of the scroll area
        label = QLabel(button_name, self)
        if button_name == "الإبصلمودية السنوية" or button_name == "الإبصلمودية الكيهكية":
            label.setStyleSheet("font-size: 16px; color: black; background: transparent; border: none; font-weight: bold;")  # Set the font size and color
        else:
            label.setStyleSheet("font-size: 20px; color: black; background: transparent; border: none; font-weight: bold;")
        label.setFixedHeight(40)  # Set the height of the label
        label.setAlignment(Qt.AlignCenter)  # Center align the text
        label.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Expanding )  # Set size policy to expanding        

        # Create and add the search bar
        self.search_bar = QLineEdit(self)
        self.search_bar.setPlaceholderText("بحث")
        self.search_bar.setFixedHeight(40)  # Set height to match back button
        self.search_bar.setLayoutDirection(Qt.RightToLeft)  # Set layout direction to right-to-left
        self.search_bar.setStyleSheet("""
            QLineEdit {
                text-align: center;
                border: 2px solid #c4c4c4;
                border-radius: 15px;
                padding: 5px 10px;
                background-color: #f9f9f9;
                font-size: 16px;
                color: #333333;
            }
            QLineEdit:focus {
                border-color: #a0a0ff;
                background-color: #ffffff;
            }
        """)  # Align placeholder text to center
        self.search_bar.textChanged.connect(self.filter_buttons)  # Connect the textChanged signal to the filter function
        label.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Expanding )  # Set size policy to expanding

        # Add the search bar to the layout
        label_search_layout.addWidget(self.search_bar)
        label_search_layout.addWidget(label)
        # Set the label to take as much width as it needs and allocate the remaining width to the search bar
        label_search_layout.setStretch(0, 1)
        label_search_layout.setStretch(1, 1)

        container_layout.addLayout(label_search_layout)

        # Create a new scroll area for the data
        self.data_scroll_area = QScrollArea()
        self.data_scroll_area.setFixedSize(320, 360)  # Set the height of the frame and width constraints
        self.data_scroll_area.setLayoutDirection(Qt.RightToLeft)  # Set layout direction to right-to-left
        self.data_scroll_area.setStyleSheet("background-color: transparent; border: none; color: white;")

        scroll_content = QWidget()
        scroll_layout = QVBoxLayout(scroll_content)
        scroll_layout.setAlignment(Qt.AlignTop)  # Set the alignment to top

        for item in data:
            button = QPushButton(item)
            button.clicked.connect(lambda _, i=item_to_index[item]: self.open_or_goto_slide(file_path, i))
            self.set_right_aligned_button_style(button)
            scroll_layout.addWidget(button)
        self.data_scroll_area.setWidget(scroll_content)

        # Set stylesheet for scrollbar to make it transparent
        self.data_scroll_area.verticalScrollBar().setStyleSheet(
            "QScrollBar:vertical {border: none; background: transparent; width: 10px;}"
            "QScrollBar::handle:vertical {background: rgba(255, 255, 255, 100); border-radius: 5px;}"
            "QScrollBar::add-line:vertical {background: none;}"
            "QScrollBar::sub-line:vertical {background: none;}"
            "QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical {background: rgba(0, 0, 0, 100);}"
        )
        self.data_scroll_area.horizontalScrollBar().setStyleSheet(
            "QScrollBar:horizontal {border: none; background: transparent; height: 10px;}"
            "QScrollBar::handle:horizontal {background: rgba(255, 255, 255, 100); border-radius: 5px;}"
            "QScrollBar::add-line:horizontal {background: none;}"
            "QScrollBar::sub-line:horizontal {background: none;}"
            "QScrollBar::add-page:horizontal, QScrollBar::sub-page:horizontal {background: rgba(0, 0, 0, 100);}"
        )

        # Add the scroll area to the container layout
        container_layout.addWidget(self.data_scroll_area)

        # Add the container widget to the main layout on the left side
        self.frame_layout.insertWidget(0, self.scroll_area_container)

        # Set the horizontal scroll bar to be initially on the right
        self.data_scroll_area.horizontalScrollBar().setValue(self.data_scroll_area.horizontalScrollBar().minimum())

    def open_or_goto_slide(self, ppt_path, slide_index):
        powerpoint = win32com.client.Dispatch("PowerPoint.Application")
        
        # Set up a flag to check if the presentation is already open
        is_open = False
        presentation = None

        # Iterate over the currently open presentations
        for pres in powerpoint.Presentations:
            if pres.FullName == ppt_path:  # Check if the file path matches
                is_open = True
                presentation = pres
                break
        
        if not is_open:
            presentation = powerpoint.Presentations.Open(ppt_path)
        
        #select the specified slide
        slide = presentation.Slides(slide_index)
        slide.Select()

        # Make sure PowerPoint window is visible
        powerpoint.Visible = True

    def katamars_button_click(self, path):
        # Check if the scroll area exists and remove it
        if hasattr(self, 'data_scroll_area') and self.data_scroll_area is not None:
            self.frame_layout.removeWidget(self.scroll_area_container)
            self.scroll_area_container.deleteLater()
            del self.data_scroll_area  # Remove the reference to the scroll area
            self.image_label = QLabel(self)
            pixmap = QPixmap(relative_path(r"Data\الصور\الفهرس.png"))
            self.image_label.setPixmap(pixmap)
            self.image_label.setGeometry(12, 25, 300, 400)  # Adjust the size to make it bigger
            self.image_label.setScaledContents(True)
            self.image_label.setAlignment(Qt.AlignCenter)  # Center align the image
            self.image_label.setStyleSheet("background-color: transparent;border: none;")
            self.frame_layout.insertWidget(0, self.image_label)  # Insert the image back to the left side

            # Re-add the image on the left if it doesn't already exist
        elif  hasattr(self, 'image_label') or self.image_label is None:
            pass

        powerpoint = win32com.client.Dispatch("PowerPoint.Application")
        is_open = False

        for pres in powerpoint.Presentations:
            if pres.FullName == relative_path(path):  # Check if the file path matches
                is_open = True
                break

        if is_open:
            self.notification_bar.show_message("الملف مفتوح بالفعل!")
        else:   
            # If the presentation is not open, open it
            powerpoint.Presentations.Open(relative_path(path))

    def update_katamars_files(self):
        self.show_progress_ui("جاري التحديث...")

        # Define a progress callback function
        def progress_callback(progress):
            self.progress_bar.setValue(progress)
            QApplication.processEvents()  # Ensure the UI updates

        # Run the All function with progress tracking
        self.worker_thread = WorkerThread(All, progress_callback=progress_callback)
        self.worker_thread.finished.connect(self.on_all_finished)
        self.worker_thread.start()

    def update_section_names(self):
        # Show the progress bar and overlay
        self.show_progress_ui("جاري تحديث أسماء الأقسام...")

        # Define the task function for updating section names
        def task_function(progress_callback):
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
                    (relative_path(r"Data\CopyData\كتاب المدائح.pptx"), "المدائح"),
                ]

                excel_file = relative_path(r'Files Data.xlsx')
                
                # Call extract_section_info2 with progress_callback
                extract_section_info2(file_sheet_pairs, excel_file, progress_callback)

                self.replace_presentation(progress_callback)

                # Show success message
                self.notification_bar.show_message("تم التحديث بنجاح!")
            except Exception as e:
                self.notification_bar.show_message(str(e))

        # Define a progress callback function
        def progress_callback(progress):
            self.progress_bar.setValue(progress)

        # Run the task function with progress tracking
        self.worker_thread = WorkerThread(task_function, progress_callback=progress_callback)
        self.worker_thread.finished.connect(self.on_all_finished)
        self.worker_thread.start()

    def show_progress_ui(self, message):
        # Show the overlay and progress bar
        self.overlay.setVisible(True)
        self.overlay.raise_()  # Ensure the overlay is on top of all widgets
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)  # Reset progress
        self.progress_bar.raise_()  # Ensure the progress bar is above the overlay
        self.notification_bar.setText(message)
        self.notification_bar.setVisible(True)
        self.notification_bar.raise_()  # Ensure the notification bar is above the overlay
        self.notification_bar.show()
        self.frame0.raise_()

    def on_all_finished(self):
        # Hide the overlay and progress bar, and show completion message
        self.overlay.setVisible(False)
        self.progress_bar.setVisible(False)
        self.notification_bar.hide()
        self.notification_bar.setText("")
        self.notification_bar.show_message("تم التحديث بنجاح!")

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
            "QPushButton:checked {"
            "   border-color: #a0a0ff;"  # Highlight color when pressed
            "}"
        )

    def set_right_aligned_button_style(self, button):
        button.setStyleSheet(
            "QPushButton {"
            "   background-color: #f0f0f0;"
            "   border: 1px solid #c4c4c4;"
            "   border-radius: 5px;"
            "   color: #333333;"
            "   padding: 5px 10px;"
            "   font-size: 20px;"
            "   text-align: right;"  # Align text to the right
            "}"
            "QPushButton:hover {"
            "   background-color: #e0e0e0;"
            "}"
            "QPushButton:pressed {"
            "   background-color: #d9d9d9;"
            "}"
        )

    def normalize_text(self, text):
        """Normalize text by replacing alif with hamza."""
        return text.replace('أ', 'ا').replace('ؤ', 'و').replace('إ', 'ا').replace('آ', 'ا').lower()

    def filter_buttons(self):
        search_text = self.normalize_text(self.search_bar.text())  # Normalize the search text
        if hasattr(self, 'data_scroll_area') and self.data_scroll_area is not None:
            scroll_content = self.data_scroll_area.widget()
            if scroll_content:
                layout = scroll_content.layout()
                visible_count = 0  # Track the number of visible buttons

                # Iterate through the buttons in the layout
                for i in range(layout.count()):
                    button = layout.itemAt(i).widget()
                    if isinstance(button, QPushButton):
                        # Normalize the button's text for comparison
                        button_text = self.normalize_text(button.text())
                        # Check if the button's text contains the search text
                        is_visible = search_text in button_text
                        button.setVisible(is_visible)  # Show/hide based on match
                        if is_visible:
                            visible_count += 1

                # Adjust the size of the scroll content dynamically
                scroll_content.adjustSize()

                # Dynamically adjust the scroll area height while maintaining a minimum height
                min_height = 360  # Minimum height for the scroll area
                max_height = 360  # Maximum height for the scroll area
                content_height = scroll_content.sizeHint().height()

                # Ensure the scroll area stays visible even if no buttons are visible
                if visible_count == 0:
                    # Add a placeholder widget if no buttons are visible
                    if not hasattr(self, 'placeholder_label'):
                        self.placeholder_label = QLabel("لا توجد نتائج")
                        self.placeholder_label.setAlignment(Qt.AlignCenter)
                        self.placeholder_label.setStyleSheet("color: black; font-size: 20px; font-weight: bold;")
                        layout.addWidget(self.placeholder_label)
                    self.placeholder_label.setVisible(True)
                    self.data_scroll_area.setFixedHeight(min_height)  # Set to minimum height
                else:
                    # Hide the placeholder widget if buttons are visible
                    if hasattr(self, 'placeholder_label'):
                        self.placeholder_label.setVisible(False)
                    self.data_scroll_area.setFixedHeight(max(min_height, min(max_height, content_height)))

                # Ensure buttons are always aligned to the top
                layout.setAlignment(Qt.AlignTop)
    
    def replace_presentation(self, progress_callback=None):
        from shutil import copy2
        from os import path, remove

        replace_flags = {
            'odasEltfl': True, 
            'bakerWaashya': True, 
            'tasbha': True, 
            'tasbhaKiahk': True,
            'zoksologyat': True,
            'default': True
        }

        presentations = {
            'odasEltfl': (r"قداس الطفل.pptx", r"Data\CopyData\قداس الطفل.pptx"),
            'bakerWaashya': (r"رفع بخور عشية و باكر.pptx", r"Data\CopyData\رفع بخور عشية و باكر.pptx"),
            'tasbha': (r"الإبصلمودية.pptx", r"Data\CopyData\الإبصلمودية.pptx"),
            'tasbhaKiahk': (r"الإبصلمودية الكيهكية.pptx", r"Data\CopyData\الإبصلمودية الكيهكية.pptx"),
            'zoksologyat': (r"الذكصولوجيات.pptx", r"Data\CopyData\الذكصولوجيات.pptx"),
            'default': (r"قداس.pptx", r"Data\CopyData\قداس.pptx")
        }

        total_steps = sum(replace_flags.values())
        completed_steps = 0

        for key, flag in replace_flags.items():
            if flag:
                old_presentation_path = relative_path(presentations[key][0])
                new_presentation_path = relative_path(presentations[key][1])

                try:
                    if path.exists(old_presentation_path):
                        remove(old_presentation_path)  # Delete old presentation
                        copy2(new_presentation_path, old_presentation_path)  # Copy new file
                    else:
                        copy2(new_presentation_path, old_presentation_path)

                    completed_steps += 1
                    if progress_callback:
                        progress_callback(50 + int(completed_steps / total_steps * 50))  # Second 50% progress

                except Exception as e:
                    self.notification_bar.show_message(f"Error: {e}")

    def go_back(self):
        self.close()

# from sys import argv, exit
# app = QApplication(argv)
# window = elfhrswindow()
# window.show()
# exit(app.exec_())
