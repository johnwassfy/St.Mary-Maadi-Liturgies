import os
import pandas as pd
from PyQt5.QtWidgets import QApplication, QHBoxLayout, QCheckBox, QMainWindow, QLabel, QVBoxLayout, QFrame, QLineEdit, QPushButton, QScrollArea, QWidget
from PyQt5.QtGui import QPixmap, QFont, QIcon
from PyQt5.QtCore import Qt, pyqtSignal
from PyQt5.QtGui import QMouseEvent
from NotificationBar import NotificationBar  # Assuming NotificationBar is in the same directory
import win32com.client
from commonFunctions import relative_path, load_background_image

class CustomButton(QWidget):
    clicked = pyqtSignal()  # Signal to mimic QPushButton's clicked signal

    def __init__(self, text, parent=None):
        super().__init__(parent)

        # Set up a layout for the custom button
        layout = QVBoxLayout(self)
        self.label = QLabel(text)
        self.label.setWordWrap(True)  # Enable word wrapping for long text

        # Customize the appearance
        self.label.setAlignment(Qt.AlignCenter)
        self.label.setStyleSheet("QLabel { color: black; padding: 5px; font: 16px;}")

        self.setStyleSheet("""
            QWidget {
                background-color: rgba(255, 255, 255, 170);
                border: 1px solid black;
                border-radius: 5px;
            }
            QWidget:hover {
                background-color: #f0f0f0;
            }
        """)

        # Add the label to the layout
        layout.addWidget(self.label)
        self.setLayout(layout)

    def mousePressEvent(self, event: QMouseEvent):
        if event.button() == Qt.LeftButton:
            self.clicked.emit()  # Emit the clicked signal when pressed

    def sizeHint(self):
        return self.label.sizeHint()

class Taranymwindow(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("St. Mary Maadi Liturgies")
        self.setWindowIcon(QIcon(relative_path(r"Data\الصور\Logo.ico")))
        self.setGeometry(400, 100, 625, 600)
        self.setFixedSize(625, 600)

        # Create a central widget
        self.central_widget = QLabel(self)
        self.central_widget.setAlignment(Qt.AlignCenter)
        self.setCentralWidget(self.central_widget)

        # Load background image
        try:
            load_background_image(self.central_widget)
        except Exception as e:
            self.notification_bar(f"خطأ في تحميل الخلفية: {str(e)}")

        self.frame0 = QFrame(self)
        self.frame0.setGeometry(0, 0, 625, 70)
        self.frame0.setStyleSheet("background-color: #ffffff;")

        # Add the picture to frame0
        try:
            self.frame_background_image(self.frame0, 625, 70, r"Data\الصور\Untitled-2.png")
        except Exception as e:
            self.show_error(f"خطأ في تحميل صورة الإطار: {str(e)}")

        self.frame = QFrame(self)
        self.frame.setGeometry(20, 90, 585, 450)
        self.frame.setStyleSheet("QFrame{ background-color: rgba(4, 30, 46, 240); border: 2px solid black; }")

        # Add a title on top of the frame
        title_label = QLabel("الدرة الأرثوذكسية في المدائح و التراتيل الكنسية", self.frame)
        title_label.setGeometry(0, 0, 585, 70)
        title_label.setAlignment(Qt.AlignCenter)
        title_label.setFont(QFont('Urdu Typesetting', 32))
        title_label.setStyleSheet("background-color: rgba(0, 0, 0, 0); color: rgb(253, 208, 23); border: none;")

        # Create and populate three scroll areas
        try:
            self.create_scroll_areas()
        except Exception as e:
            self.show_error(f"خطأ في إنشاء مناطق التمرير: {str(e)}")

        # Create a container for the search box and checkbox
        container_widget = QWidget(self)
        container_widget.setGeometry(20, 560, 260, 40)  # Adjust position and size as needed
        
        # Set the container background color to white
        container_widget.setStyleSheet("background-color: rgba(255, 255, 255, 250); ")
        container_layout = QHBoxLayout(container_widget)
        
        # Create and add the checkbox
        self.slideshow_checkbox = QCheckBox("slideshow", self)  # "Enter Slideshow Mode" in Arabic
        self.slideshow_checkbox.setChecked(False)  # Default to unchecked
        container_layout.addWidget(self.slideshow_checkbox)

        # Create and add the search bar
        # Update the search bar styling in Taranymwindow
        self.search_bar = QLineEdit(self)
        self.search_bar.setPlaceholderText("ابحث على المديح")
        self.search_bar.setFixedHeight(30)  # Match the height of the search bar in elfhrswindow
        self.search_bar.setLayoutDirection(Qt.RightToLeft)  # Set layout direction to right-to-left
        self.search_bar.setStyleSheet("""
            QLineEdit {
                text-align: center;
                border: 2px solid #c4c4c4;
                border-radius: 15px;
                padding: 5px 10px;
                background-color: #f9f9f9;
                font-size: 13px;
                color: #333333;
            }
            QLineEdit:focus {
                border-color: #a0a0ff;
                background-color: #ffffff;
            }
        """)  # Match the style from elfhrswindow
        self.search_bar.textChanged.connect(self.filter_buttons)  # Connect the textChanged signal to the filter function
        container_layout.addWidget(self.search_bar)

        # Add back button
        self.back_button = QPushButton("Back")
        layout = QVBoxLayout(self.central_widget)
        layout.addWidget(self.back_button, alignment=Qt.AlignBottom | Qt.AlignRight)
        self.back_button.clicked.connect(self.close)

        # Add NotificationBar for error messages
        self.notification_bar = NotificationBar(self)
        self.notification_bar.setGeometry(0, 0, 625, 50)  # Position at the top of the window

    def show_error(self, message):
        """Display an error message using the NotificationBar."""
        self.notification_bar.show_message(message, duration=5000)

    def frame_background_image(self, frame, w, h, image_relative_path):
        try:
            image_label = QLabel(frame)
            image_label.setGeometry(0, 0, w, h)
            image_path = relative_path(image_relative_path)
            pixmap = QPixmap(image_path)
            image_label.setPixmap(pixmap)
            image_label.setScaledContents(True)
        except Exception as e:
            raise RuntimeError(f"خطأ في تحميل صورة الإطار: {str(e)}")

    def create_scroll_areas(self):
        try:
            # Define the labels for each scroll area
            labels_data = [
                {"text": "المناسبات و مدائح العذراء", "x": 395, "y": 80},  # Label for Scroll Area 3 (now on the right)
                {"text": "الملائكة والقديسين", "x": 203, "y": 80},  # Label for Scroll Area 2 (in the middle)
                {"text": "الترانيم", "x": 10, "y": 80}    # Label for Scroll Area 1 (now on the left)
            ]

            # Create labels for each scroll area
            for label_data in labels_data:
                label = QLabel(label_data['text'], self.frame)
                label.setGeometry(label_data['x'], label_data['y'], 180, 30)  # Adjust the size as needed
                label.setAlignment(Qt.AlignCenter)
                label.setStyleSheet("background-color: rgba(0, 0, 0, 0); color: white; font-size: 16px;")  # You can adjust the font size and color

            # Define the scroll areas data
            scroll_areas_data = [
                {"x": 395, "y": 110, "width": 180, "height": 330},  # Swapped position of scroll area 3
                {"x": 203, "y": 110, "width": 180, "height": 330},  # Scroll area 2 remains the same
                {"x": 10, "y": 110, "width": 180, "height": 330}    # Swapped position of scroll area 1
            ]

            self.scroll_layouts = []  # Store layouts for each scroll area

            for data in scroll_areas_data:
                scroll_layout = self.create_scroll_area(data['x'], data['y'], data['width'], data['height'])
                self.scroll_layouts.append(scroll_layout)

            # Populate scroll areas
            self.load_excel_strings()
        except Exception as e:
            raise RuntimeError(f"خطأ في إنشاء مناطق التمرير: {str(e)}")

    def load_excel_strings(self):
        try:
            # Replace with the path to your Excel file
            excel_file_path = relative_path(r"Files Data.xlsx")

            # Specify the sheet name where the search should be performed
            sheet_name = 'المدائح'  # Replace with the actual sheet name

            # Read the specified sheet in the Excel file
            df = pd.read_excel(excel_file_path, sheet_name=sheet_name)

            # Get the strings and slide indices
            strings = df.iloc[1:, 0].values  # Button labels
            slide_indices = df.iloc[1:, 2].values  # Slide indices (next column)

            # Define break cells (as an example)
            break_cell_1 = "الباب الاول"
            break_cell_2 = "الباب الثاني"
            break_cell_3 = "الباب الثالث"

            # Split the content based on the break cells
            scroll_area_contents = [[], [], []]  # Three lists for scroll areas
            scroll_area_indices = [[], [], []]  # Three lists for slide indices

            current_area = -1  # To keep track of which scroll area to add content to

            for string, index in zip(strings, slide_indices):
                if string == break_cell_1:
                    current_area = 0
                elif string == break_cell_2:
                    current_area = 1
                elif string == break_cell_3:
                    current_area = 2
                elif current_area != -1:  # Only add content if a valid area is set
                    scroll_area_contents[current_area].append(string)
                    scroll_area_indices[current_area].append(index)

            # Populate each scroll area with buttons
            for i in range(3):
                self.populate_scroll_area(self.scroll_layouts[i], scroll_area_contents[i], scroll_area_indices[i])
        except FileNotFoundError:
            self.show_error("ملف Files Data غير موجود.")
        except Exception as e:
            self.show_error(f"خطأ في تحميل بيانات الإكسل: {str(e)}")

    def frame_background_image(self, frame, w, h, image_relative_path):
        image_label = QLabel(frame)
        image_label.setGeometry(0, 0, w, h)
        image_path = relative_path(image_relative_path)
        pixmap = QPixmap(image_path)
        image_label.setPixmap(pixmap)
        image_label.setScaledContents(True)

    def create_scroll_areas(self):
        # Define the labels for each scroll area
        labels_data = [
            {"text": "المناسبات و مدائح العذراء", "x": 395, "y": 80},  # Label for Scroll Area 3 (now on the right)
            {"text": "الملائكة والقديسين", "x": 203, "y": 80},  # Label for Scroll Area 2 (in the middle)
            {"text": "الترانيم", "x": 10, "y": 80}    # Label for Scroll Area 1 (now on the left)
        ]

        # Create labels for each scroll area
        for label_data in labels_data:
            label = QLabel(label_data['text'], self.frame)
            label.setGeometry(label_data['x'], label_data['y'], 180, 30)  # Adjust the size as needed
            label.setAlignment(Qt.AlignCenter)
            label.setStyleSheet("background-color: rgba(0, 0, 0, 0); color: white; font-size: 16px;")  # You can adjust the font size and color

        # Define the scroll areas data
        scroll_areas_data = [
            {"x": 395, "y": 110, "width": 180, "height": 330},  # Swapped position of scroll area 3
            {"x": 203, "y": 110, "width": 180, "height": 330},  # Scroll area 2 remains the same
            {"x": 10, "y": 110, "width": 180, "height": 330}    # Swapped position of scroll area 1
        ]

        self.scroll_layouts = []  # Store layouts for each scroll area

        for data in scroll_areas_data:
            scroll_layout = self.create_scroll_area(data['x'], data['y'], data['width'], data['height'])
            self.scroll_layouts.append(scroll_layout)

        # Populate scroll areas
        self.load_excel_strings()

    def create_scroll_area(self, x, y, width, height):
        scroll_area = QScrollArea(self.frame)
        scroll_area.setGeometry(x, y, width, height)
        scroll_area.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        scroll_area.setStyleSheet("""
            QScrollArea { background-color: rgba(0, 0, 0, 120); border: 1px solid black; }  # Semi-transparent background
            QScrollBar:vertical { background: transparent; width: 10px; }
            QScrollBar::handle:vertical { background: rgba(255, 255, 255, 150); }  # Semi-transparent scrollbar handle
        """)

        scroll_area.verticalScrollBar().setStyleSheet(
            "QScrollBar:vertical {border: none; background: transparent; width: 10px;}"
            "QScrollBar::handle:vertical {background: rgba(255, 255, 255, 100); border-radius: 5px;}"
            "QScrollBar::add-line:vertical {background: none;}"
            "QScrollBar::sub-line:vertical {background: none;}"
            "QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical {background: none;}"
        )

        # Create a widget to hold the scroll area content
        scroll_content = QWidget()
        scroll_area.setWidget(scroll_content)
        scroll_area.setWidgetResizable(True)

        # Create a vertical layout for the scroll area content
        scroll_layout = QVBoxLayout(scroll_content)
        scroll_content.setStyleSheet("background-color: rgba(255, 255, 255, 0);")  # Semi-transparent content area
        return scroll_layout

    def load_excel_strings(self):
        # Replace with the path to your Excel file
        excel_file_path = relative_path(r"Files Data.xlsx")

        # Specify the sheet name where the search should be performed
        sheet_name = 'المدائح'  # Replace with the actual sheet name

        # Read the specified sheet in the Excel file
        df = pd.read_excel(excel_file_path, sheet_name=sheet_name)

        # Get the strings and slide indices
        strings = df.iloc[1:, 0].values  # Button labels
        slide_indices = df.iloc[1:, 2].values  # Slide indices (next column)

        # Define break cells (as an example)
        break_cell_1 = "الباب الاول"
        break_cell_2 = "الباب الثاني"
        break_cell_3 = "الباب الثالث"

        # Split the content based on the break cells
        scroll_area_contents = [[], [], []]  # Three lists for scroll areas
        scroll_area_indices = [[], [], []]  # Three lists for slide indices

        current_area = -1  # To keep track of which scroll area to add content to

        for string, index in zip(strings, slide_indices):
            if string == break_cell_1:
                current_area = 0
            elif string == break_cell_2:
                current_area = 1
            elif string == break_cell_3:
                current_area = 2
            elif current_area != -1:  # Only add content if a valid area is set
                scroll_area_contents[current_area].append(string)
                scroll_area_indices[current_area].append(index)

        # Populate each scroll area with buttons
        for i in range(3):
            self.populate_scroll_area(self.scroll_layouts[i], scroll_area_contents[i], scroll_area_indices[i])

    def populate_scroll_area(self, layout, contents, indices):
        for content, index in zip(contents, indices):
            # Create the custom button with text wrapping
            button = CustomButton(content)
            button.setFixedWidth(160)  # Set fixed width

            # Connect the click signal to the method with the slide index and checkbox state
            button.clicked.connect(lambda text=content, idx=index: self.open_or_goto_slide(
                relative_path(r"كتاب المدائح.pptx"), idx, self.slideshow_checkbox.isChecked()
            ))

            # Add the button to the scroll area's layout
            layout.addWidget(button)

    def open_or_goto_slide(self, ppt_path, slide_index, slideshow):
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
            if(self.pptx_check() == False):
                self.replace_presentation()
            presentation = powerpoint.Presentations.Open(ppt_path)
        
        # Check if the presentation is already in slideshow mode
        is_showing = False
        slide_show = None
        for show in powerpoint.SlideShowWindows:
            if show.Presentation == presentation:
                is_showing = True
                slide_show = show
                break

        if is_showing:
            if not slideshow:
                # If the checkbox is unchecked and slideshow is currently open, close the slideshow
                slide_show.View.Exit()
                # Select the specified slide
                slide = presentation.Slides(slide_index)
                slide.Select()
            else:
                # If in slideshow mode and slideshow is true, just go to the specified slide
                slide_show.View.GotoSlide(slide_index)
        elif not is_showing and slideshow:
            # Start slideshow if not currently in slideshow mode
            slide_show = presentation.SlideShowSettings.Run()
            slide_show.View.GotoSlide(slide_index)

        else:
            # If slideshow is false, just select the specified slide
            slide = presentation.Slides(slide_index)
            slide.Select()

        # Make sure PowerPoint window is visible
        powerpoint.Visible = True

    def normalize_text(self, text):
        """Normalize text by replacing alif with hamza."""
        return text.replace('أ', 'ا').replace('ؤ', 'و').replace('إ', 'ا').lower()

    def filter_buttons(self):
        search_text = self.normalize_text(self.search_bar.text())  # Normalize the search text
        for layout in self.scroll_layouts:
            for i in range(layout.count()):  # Iterate through the buttons in the layout
                button = layout.itemAt(i).widget()
                if isinstance(button, CustomButton):
                    # Normalize the button's text for comparison
                    button_text = self.normalize_text(button.label.text())
                    # Check if the button's text contains the search text
                    button.setVisible(search_text in button_text)  # Show/hide based on match

    def pptx_check(self):
        from openpyxl import load_workbook
        from pptx import Presentation
        try:
            wb = load_workbook(relative_path(r"Files Data.xlsx"))
            presentation = Presentation(relative_path(r"كتاب المدائح.pptx"))
            sheet = wb["المدائح"]

            num_slides = len(presentation.slides)
            intpptx = num_slides

            # Reading the second-to-last non-empty value from column 'C'
            last_non_empty_c = None
            second_last_non_empty_c = None
            for cell in sheet['C']:
                if cell.value is not None:
                    second_last_non_empty_c = last_non_empty_c
                    last_non_empty_c = cell.value

            if second_last_non_empty_c is None:
                return False
            elif second_last_non_empty_c != intpptx:
                return False

            last_non_empty_b_cell = None
            for cell in sheet['B']:
                if cell.value is not None:
                    last_non_empty_b_cell = cell

            if last_non_empty_b_cell is None:
                return False
            elif not bool(last_non_empty_b_cell.value):  # Check if the value is False
                # Change the value to True and save the change
                last_non_empty_b_cell.value = True
                wb.save(relative_path(r"Files Data.xlsx"))
                return False

        except Exception as e:
            print(f"Error: {str(e)}")
            return None

    def replace_presentation(self):
        from shutil import copy2
        from os import path, remove
        old_presentation_path = relative_path(r"كتاب المدائح.pptx")
        new_presentation_path = relative_path(r"Data\CopyData\كتاب المدائح.pptx")
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
# window = Taranymwindow()
# window.show()
# exit(app.exec_())
