import time
from PyQt5.QtWidgets import QApplication, QMainWindow, QLabel, QPushButton, QFrame
from PyQt5.QtGui import QPixmap, QFont, QIcon
from PyQt5.QtCore import Qt, pyqtSignal
from WorkerThread import WorkerThread
from copticDate import CopticCalendar
from Season import get_season_name, get_season
from datetime import datetime
from elbas5aWindow import elbas5aWindow
from elLakanWindow import ellakanwindow
from bibleWindow import bibleWindow
from NotificationBar import NotificationBar
import asyncio
from commonFunctions import relative_path, load_background_image, open_presentation_relative_path
from sys import exit, argv
from SplashScreen import ModernSplashScreen

class ClickableFrame(QFrame):
    clicked = pyqtSignal()

    def mousePressEvent(self, event):
        self.clicked.emit()
        super().mousePressEvent(event)

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        try:
            self.current_date = datetime.now()
            self.coptic_date = CopticCalendar().gregorian_to_coptic(self.current_date)
            self.checkCopticYear(self.coptic_date[0])
            self.season = get_season(self.current_date)
            self.bishop_window = None
            self.bishop = False
            self.GuestBishop = 0
            self.setWindowTitle("St. Mary Maadi Liturgies")
            self.setWindowIcon(QIcon(relative_path(r"Data\الصور\Logo.ico")))
            self.setGeometry(400, 100, 625, 600)
            self.setFixedSize(625, 600)

            # Background label
            self.background_label = QLabel(self)
            self.background_label.setGeometry(0, 0, self.width(), self.height())
                # Load background image
            try:
                load_background_image(self.background_label)
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

            frame1 = ClickableFrame(self)
            frame1.setGeometry(20, 86, 585, 190)
            frame1.setStyleSheet("QFrame { background-color: rgba(107, 6, 6, 200); border: 2px solid black; }")
            frame1.clicked.connect(lambda: self.open_new_window())

            label1 = QLabel(self)
            label1.setObjectName("label1")
            label1.setAlignment(Qt.AlignCenter)
            label1.setGeometry(50, 0, 585, 190)
            label1.setParent(frame1)

            font = QFont()
            font.setPointSize(30)
            font.setFamily("Calibri")
            label1.setFont(font)
            label1.setStyleSheet("color: white;")

            self.image_label = QLabel(frame1)
            self.image_label.setGeometry(0, 0, 130, 190)
            self.image_label.setScaledContents(True)

            self.frame2 = QFrame(self)
            self.restore_main_frame()
            asyncio.run(self.create_button("تحديث بيانات القداسات", self.width() - 115, 566, self.update_section_names))
            asyncio.run(self.create_button("في حضور الأسقف", self.width() - 240, 566, self.open_bishop_window))
            asyncio.run(self.create_button("اضافة تعديل خاص", self.width() - 365, 566, self.open_bishop_window))
            asyncio.run(self.update_labels())
            
            # Add NotificationBar
            self.notification_bar = NotificationBar(self)
            self.notification_bar.setGeometry(0, 70, self.width(), 50)

        except Exception as e:
            # Add NotificationBars
            self.notification_bar = NotificationBar(self)
            self.notification_bar.setGeometry(0, 70, self.width(), 50)
            self.notification_bar.show_message(str(e))

    async def add_button_with_image(self, parent, image_path, geometry, text, action=None):
        x, y, width, height = geometry

        # Image Label
        image_label = QLabel(parent)
        pixmap = QPixmap(relative_path(image_path))
        image_label.setPixmap(pixmap)
        image_label.setGeometry(x, y, width, height)
        image_label.setScaledContents(True)

        # Button
        button = QPushButton(parent)
        button.setGeometry(x, y, width, height)
        button.setStyleSheet("""
            QPushButton {
                background-color: transparent;
                border: none;
            }
            QPushButton:hover {
                background-color: rgba(173, 216, 230, 50);
            }
            QPushButton:pressed {
                background-color: rgba(173, 216, 230, 100);
            }
        """)
        if text == "صلاة السجدة":
            button.clicked.connect(lambda _, p=action: open_presentation_relative_path(p))
        else :
            button.clicked.connect(action)
        # Text Label for Button
        label = QLabel(text, parent)
        label.setAlignment(Qt.AlignCenter)
        label.setGeometry(x, y + height - 5, width, 30)
        font = QFont()
        if text == "الكتاب المقدس":
            font.setPointSize(10)
        else:
            font.setPointSize(12)
        label.setFont(font)
        label.setStyleSheet("background-color: transparent; color: white; border: none; font-weight: bold;")

    async def create_button(self, text, x, y, action):
        button = QPushButton(text, self)
        button.setGeometry(x, y, 115, 30)
        button.clicked.connect(action)

    def open_bishop_window(self):
        from GuestWindow import Bishop
        self.bishop = False
        self.GuestBishop = 0
        if not self.bishop_window:
            self.bishop_window = Bishop()
            self.bishop_window.row2.line_edit.textChanged.connect(self.update_checkbox_state)
            self.bishop_window.update_button.clicked.connect(self.update_bishop_variables)
        self.bishop_window.show()

    def update_checkbox_state(self):
        # If row2's line edit has text, check the checkbox
        if self.bishop_window.row2.line_edit.text():
            self.bishop_window.checkbox1.setChecked(True)
        else:
            self.bishop_window.checkbox1.setChecked(False)

    def update_bishop_variables(self):
        # Update self.bishop based on the checkbox state in Bishop window
        self.bishop = self.bishop_window.checkbox1.isChecked()

        if self.bishop_window.row2.line_edit.text():
            self.GuestBishop += 1
        if self.bishop_window.row3.line_edit.text():
            self.GuestBishop += 1
        # Hide the Bishop window after updating variables
        self.bishop_window.hide()

    def open_tasbha_Window(self):
        from tasbhaWindow import tasbhawindow
        if self.centralWidget():
            self.clear_central_widget()
        
        tasbha_content = tasbhawindow(self.coptic_date, self.season)
        self.setCentralWidget(tasbha_content)
        self.replace_presentation(False, False, True)

    def open_elmonasbat_Window(self):
        self.frame2.deleteLater()
        self.frame2 = QFrame(self)
        self.frame2.setGeometry(20, 286, 585, 275)
        self.frame2.setStyleSheet("QFrame { background-color: rgba(107, 6, 6, 200); border: 2px solid black; }")

        # Add back button
        self.back_button = QPushButton("Back")
        asyncio.run(self.add_button_with_image(self.frame2, "Data/الصور/البصخة.jpg", (20, 20, 100, 100), "اسبوع الالام", self.open_elbas5a_window))
        asyncio.run(self.add_button_with_image(self.frame2, "Data\الصور\السجدة.jpg", (140, 20, 100, 100), "صلاة السجدة", "Data\صلاة السجدة عيد العنصرة.pptx"))
        asyncio.run(self.add_button_with_image(self.frame2, "Data\الصور\اللقان.jpg", (260, 20, 100, 100), "اللقان", self.open_ellakan_window))
        self.add_back_button(self.frame2, self.restore_main_frame)
        self.frame2.show()

    def open_elbas5a_window(self):
        if self.centralWidget():
            self.clear_central_widget()
        
        elbas5a_content = elbas5aWindow(self)
        self.setCentralWidget(elbas5a_content)

    def open_ellakan_window(self):
        if self.centralWidget():
            self.clear_central_widget()
        
        ellakan_content = ellakanwindow()
        self.setCentralWidget(ellakan_content)

    def open_bible_window(self):
        if self.centralWidget():
            self.clear_central_widget()
        
        ellakan_content = bibleWindow()
        self.setCentralWidget(ellakan_content)

    def open_elfhrs_window(self):
        from elfhrsNEWindow import elfhrswindow
        if self.centralWidget():
            self.clear_central_widget()

        elfhrs_content = elfhrswindow()
        self.setCentralWidget(elfhrs_content)

    def open_taranym_window(self):
        from TaranymWindow import Taranymwindow
        if self.centralWidget():
            self.clear_central_widget()

        elfhrs_content = Taranymwindow()
        self.setCentralWidget(elfhrs_content)

    def update_section_names(self):
        from sectionNames import extract_section_info2
        try:
            file_sheet_pairs = [
                (relative_path(r"Data\CopyData\قداس.pptx"), "القداس"),
                (relative_path(r"Data\CopyData\قداس الطفل.pptx"), "قداس الطفل"),
                (relative_path(r"Data\CopyData\باكر.pptx"), "باكر"),
                (relative_path(r"Data\CopyData\عشية.pptx"), "عشية"),
                (relative_path(r"Data\CopyData\رفع بخور عشية و باكر.pptx"), "رفع بخور"),
                (relative_path(r"Data\CopyData\الذكصولوجيات.pptx"), "الذكصولوجيات"),
                (relative_path(r"Data\CopyData\في حضور الاسقف و اساقفة ضيوف.pptx"), "في حضور الأسقف"),
                (relative_path(r"Data\CopyData\الإبصلمودية.pptx"), "التسبحة"),
                (relative_path(r"Data\CopyData\الإبصلمودية الكيهكية.pptx"), "تسبحة كيهك"),
                (relative_path(r"Data\CopyData\كتاب المدائح.pptx"), "المدائح")
            ]

            excel_file = relative_path(r'بيانات القداسات.xlsx')
            
            extract_section_info2(file_sheet_pairs, excel_file)

            # Show success message
            self.show_message("تم التحديث بنجاح!")

        except Exception as e:
            self.show_error_message(str(e))

    def season_picture(self):
        match self.season :
            case 0:
                return r"Data\الصور\Aykona.png"
            case 4 | 4.1:
                return r"Data\الصور\عيد الميلاد.jpg"
            case 10 :
                return r"Data\الصور\عرس قانا الجليل.jpg"
            case 17:
                return r"Data\الصور\الشعانين.jpg"
            case 19:
                return r"Data\الصور\خميس العهد.jpg"
            case 20 | 18:
                return r"Data\الصور\الجمعة العظيمةو البصخة.jpg"
            case 21:
                return r"Data\الصور\سبت النور.JPG"
            case 22 | 24:
                return r"Data\الصور\القيامة.jpg"
            case 23.3 | 24.1 | 25:
                return r"Data\الصور\الصعود.jpg"
            case 23.1 | 23:
                return r"Data\الصور\دخول المسيح أرض مصر.jpg"
            case 29 :
                return r"Data\الصور\التجلي.JPG"
        return r"Data\الصور\Aykona.png" 

    def open_new_window(self):
        from ChangeDateWindow import ChangeDate
        new_window = ChangeDate(self.current_date.date(), self.current_date.strftime("%I:%M %p"))
        new_window.date_updated.connect(self.update_current_date)
        new_window.exec_()

    def clear_central_widget(self):
        central_widget = self.centralWidget()
        if central_widget:
            layout = central_widget.layout()
            if layout:
                while layout.count():
                    child = layout.takeAt(0)
                    if child.widget():
                        child.widget().deleteLater()
                self.setCentralWidget(None)

    def update_current_date(self, new_date, new_time):
        try:
            self.current_date = datetime.strptime(new_date + ' ' + new_time, '%Y-%m-%d %I:%M %p')
            self.coptic_date = CopticCalendar().gregorian_to_coptic(self.current_date)
            self.season = get_season(self.current_date)
            asyncio.run(self.update_labels())
            self.close_dialog()
        except ValueError:
            self.show_error_message("التاريخ/الوقت غير صحيح.")

    def convert_to_arabic_digits(self, number):
        arabic_digits = {'0': '٠', '1': '١', '2': '٢', '3': '٣', '4': '٤', '5': '٥', '6': '٦', '7': '٧', '8': '٨', '9': '٩'}
        return ''.join(arabic_digits[digit] if digit in arabic_digits else digit for digit in str(number))

    async def update_labels(self):
        label1 = self.findChild(QLabel, "label1")
        if label1:
            sesn = get_season_name(self.season)
            m = self.getmonth(self.coptic_date[1])
            m = self.convert_to_arabic_digits(m)
            ad = self.get_arabic_month_date(self.current_date)
            ad = self.convert_to_arabic_digits(ad)
            c = f"{self.convert_to_arabic_digits(self.coptic_date[2])} {m}، {self.convert_to_arabic_digits(self.coptic_date[0])}"
            if self.current_date.time() > datetime.strptime('5:30 PM', '%I:%M %p').time():
                c = f"({c})"
            date = f"{sesn}\n{c}\n{ad}"
            label1.setText(date)
        new_pixmap = QPixmap(relative_path(self.season_picture()))
        self.image_label.setPixmap(new_pixmap)

    def close_dialog(self):
        from ChangeDateWindow import ChangeDate
        for widget in QApplication.instance().topLevelWidgets():
            if isinstance(widget, ChangeDate):
                widget.close()

    def show_error_message(self, error_message):
        self.notification_bar.show_message(f"Error: {error_message}", duration=5000)

    def show_message(self, message):
        self.notification_bar.show_message(message, duration=3000)

    def handle_qadas_button_click(self):
        from odasat import (odasElSomElkbyr, odasElsh3anyn, odasSbtLe3azr, odasElbeshara, odasEl2yama, odasEl5amasyn_2_39, 
                            odasElso3od, odasSanawy, odasElsalyb, odasEl3nsara)
        try:
            match self.season:
                case 0 | 6 | 13 | 30 | 31:
                    odasSanawy(self.coptic_date, self.season, self.bishop, self.GuestBishop)
                case 2:
                    odasElsalyb(self.coptic_date, self.bishop, self.GuestBishop)
                case 14:
                    odasElbeshara(self.bishop, self.GuestBishop)
                case 15 | 15.1 | 15.2 | 15.3 | 15.4 | 15.5 | 15.6 | 15.7 | 15.8 | 15.9 | 15.11:
                    odasElSomElkbyr(self.coptic_date, self.season, self.bishop, self.GuestBishop)
                case 16:
                    odasSbtLe3azr(self.coptic_date, self.bishop, self.GuestBishop)
                case 17:
                    odasElsh3anyn(self.coptic_date, self.bishop, self.GuestBishop)
                case 19:
                    self.notification_bar.show_message("صلوات خميس العهد متوفرة في ملف واحد: المناسبات > اسبوع الالام > خميس العهد", 10000)
                case 20:
                    self.notification_bar.show_message("لا يوجد قداس يوم الجمعة العظيمة: المناسبات > اسبوع الالام > الجمعة العظيمة", 10000)
                case 21:
                    self.notification_bar.show_message("صلوات سبت الفرح متوفرة في ملف واحد: المناسبات > اسبوع الالام > ليلة ابوغلمسيس", 10000)
                case 22:
                    odasEl2yama(self.coptic_date, self.bishop, self.GuestBishop)
                case 24:
                    odasEl5amasyn_2_39(self.coptic_date, self.bishop, self.GuestBishop)
                case 24.1:
                    odasElso3od(self.coptic_date, self.bishop, self.GuestBishop, True)
                case 25:
                    odasElso3od(self.coptic_date, self.bishop, self.GuestBishop)
                case 26:
                    odasEl3nsara(self.coptic_date, self.bishop, self.GuestBishop)
                case default :
                    self.notification_bar.show_message(f"قداس {get_season_name(self.season)} غير متوفر حاليا")
        except Exception as e:
            self.notification_bar.show_message(str(e))

    def handle_qadas_eltfl_button_click(self):
        from odasatEltfl import (odasElSomElkbyr, odasEltflSomElrosol, odasEltfl3ydElrosol, odasSanawy, 
                                 odasEltflElnayrooz, odasEltflKiahk)
        try:
            if(self.pptx_check(True) == False):
                self.replace_presentation(True)
            match self.season:
                case 0 | 6 | 30 | 31:
                    odasSanawy(self.coptic_date, self.season)
                case 1:
                    odasEltflElnayrooz(self.coptic_date)
                case 5:
                    odasEltflKiahk(self.coptic_date)
                case 15 | 15.1:
                    odasElSomElkbyr(self.coptic_date, self.season)
                case 27:
                    odasEltflSomElrosol(self.coptic_date)
                case 28:
                    odasEltfl3ydElrosol()
                case default :
                    self.notification_bar.show_message(f"قداس {get_season_name(self.season)} غير متوفر حاليا")
        except Exception as e:
            self.show_error_message(str(e))

    def handle_baker_button_click(self):
        from openpyxl import load_workbook
        from baker import baker3ydElrosol, bakerSanawy, bakerKiahk

        coptic_cal = CopticCalendar()
        copticDate = coptic_cal.coptic_to_gregorian(self.coptic_date)
        adam = False
        if copticDate.weekday() in [0, 1, 6]:
            adam = True
        try:
            match self.season :
                case 0 | 27 | 30 | 31:
                    bakerSanawy(self.season, self.coptic_date, adam, self.bishop, self.GuestBishop)
                case 5:
                    bakerKiahk(self.coptic_date, adam, self.bishop, self.GuestBishop)
                case 28:
                    baker3ydElrosol(adam)
                    open_presentation_relative_path(r"Data\لقان عيد الرسل.pptx")
        except Exception as e:
            self.show_error_message(str(e))

    def handle_3ashya_button_click(self):
        from Aashya import aashyaKiahk, aashyaSanawy
        try:
            coptic_cal = CopticCalendar()
            copticDate = coptic_cal.coptic_to_gregorian(self.coptic_date)
            adam = False
            if copticDate.weekday() in [0, 1, 6]:
                adam = True

            match (self.season) :
                case 0 | 29 | 30 | 31: 
                    aashyaSanawy(self.season, self.coptic_date, adam, self.bishop, self.GuestBishop)
                case 5 :
                    aashyaKiahk (self.coptic_date, adam, self.bishop, self.GuestBishop)

        except Exception as e :
            self.show_error_message(str(e))

    def handle_agbya_button_click(self):
        return

    def pptx_check(self, odasEltfl=False, aashya_baker=False):
        from openpyxl import load_workbook
        from pptx import Presentation
        try:
            wb = load_workbook(relative_path(r"بيانات القداسات.xlsx"))
            if odasEltfl == True:
                presentation = Presentation(relative_path(r"قداس الطفل.pptx"))
                sheet = wb["قداس الطفل"]
            elif aashya_baker == True:
                presentation = Presentation(relative_path(r"رفع بخور عشية و باكر.pptx"))
                sheet = wb["رفع بخور"]
            else:
                presentation = Presentation(relative_path(r"قداس.pptx"))
                sheet = wb["سنوي"]

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
                wb.save(relative_path(r"بيانات القداسات.xlsx"))
                return False

        except Exception as e:
            print(f"Error: {str(e)}")
            return None

    def replace_presentation(self, odasEltfl = False, baker = False, tasbha = False, aashya = False):
        from shutil import copy2
        from os import path, remove
        if(odasEltfl):    
            old_presentation_path = relative_path(r"قداس الطفل.pptx")
            new_presentation_path = relative_path(r"Data\CopyData\قداس الطفل.pptx")
        elif(baker):
            old_presentation_path = relative_path(r"باكر.pptx")
            new_presentation_path = relative_path(r"Data\CopyData\باكر.pptx")
        elif(tasbha):
            old_presentation_path = relative_path(r"الإبصلمودية.pptx")
            new_presentation_path = relative_path(r"Data\CopyData\الإبصلمودية.pptx")
        elif(aashya):
            old_presentation_path = relative_path(r"رفع بخور عشية و باكر.pptx")
            new_presentation_path = relative_path(r"Data\CopyData\رفع بخور عشية و باكر.pptx")
        else:    
            old_presentation_path = relative_path(r"قداس.pptx")
            new_presentation_path = relative_path(r"Data\CopyData\قداس.pptx")
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

    def get_arabic_month_date(self, current_date):
        # Define a dictionary to map month names from English to Arabic
        month_names_arabic = {
            'January': 'يناير',
            'February': 'فبراير',
            'March': 'مارس',
            'April': 'أبريل',
            'May': 'مايو',
            'June': 'يونيو',
            'July': 'يوليو',
            'August': 'أغسطس',
            'September': 'سبتمبر',
            'October': 'أكتوبر',
            'November': 'نوفمبر',
            'December': 'ديسمبر'
        }
        
        # Define a dictionary to map day names from English to Arabic
        day_names_arabic = {
            'Monday': 'الاثنين',
            'Tuesday': 'الثلاثاء',
            'Wednesday': 'الأربعاء',
            'Thursday': 'الخميس',
            'Friday': 'الجمعة',
            'Saturday': 'السبت',
            'Sunday': 'الأحد'
        }

        arabic_month = month_names_arabic[current_date.strftime('%B')]
        arabic_day = day_names_arabic[current_date.strftime('%A')]
        
        arabic_date_string = f"{arabic_day}، {current_date.day} {arabic_month} {current_date.year}"
        return arabic_date_string

    def getmonth(self, num):
        from openpyxl import load_workbook
        # Load the Excel file
        workbook = load_workbook(relative_path(r'Tables.xlsx'))
        sheet = workbook["المناسبات"]
        search_number = num 
        corresponding_value = None
        for row in sheet.iter_rows(values_only=True):
            if row[0] == search_number: 
                corresponding_value = row[1] 
                break
        return  corresponding_value

    def add_back_button(self, parent, action):
        # Get frame geometry
        frame_geometry = parent.geometry()
        # Calculate button position (bottom right corner)
        button_width = 100
        button_height = 30
        button_x = frame_geometry.width() - button_width - 10
        button_y = frame_geometry.height() - button_height - 10

        # Add back button
        back_button = QPushButton("Back", parent)
        back_button.setGeometry(button_x, button_y, button_width, button_height)
        back_button.clicked.connect(action)
        back_button.setStyleSheet("""
            QPushButton {
                background-color: #ff5733;
                color: white;
                border-radius: 5px;
                border: none;
            }
            QPushButton:hover {
                background-color: #ff704d;
            }
            QPushButton:pressed {
                background-color: #ff8566;
            }
        """)

    def restore_main_frame(self):
        self.frame2.deleteLater()
        self.frame2 = QFrame(self)
        self.frame2.setGeometry(20, 286, 585, 275)
        self.frame2.setStyleSheet("QFrame { background-color: rgba(107, 6, 6, 200); border: 2px solid black; }")

        # Use asyncio.run to run async methods
        asyncio.run(self.add_button_with_image(self.frame2, "Data/الصور/القداس.JPG", (13, 20, 100, 100), "القداس", self.handle_qadas_button_click))
        asyncio.run(self.add_button_with_image(self.frame2, "Data/الصور/قداس الطفل.png", (126, 20, 100, 100), "قداس الطفل", self.handle_qadas_eltfl_button_click))
        asyncio.run(self.add_button_with_image(self.frame2, "Data\الصور\باكر.jpg", (239, 20, 100, 100), "باكر", self.handle_baker_button_click))
        asyncio.run(self.add_button_with_image(self.frame2, "Data\الصور\عشية.jpg", (352, 20, 100, 100), "عشية", self.handle_3ashya_button_click))
        asyncio.run(self.add_button_with_image(self.frame2, "Data/الصور/الكتاب المقدس.png", (465, 20, 100, 100), "الكتاب المقدس", self.open_bible_window))
        asyncio.run(self.add_button_with_image(self.frame2, "Data\الصور\الأجبية.jpg", (13, 150, 100, 100), "الأجبية", self.handle_agbya_button_click))
        asyncio.run(self.add_button_with_image(self.frame2, "Data\الصور\داود 1.jpg", (126, 150, 100, 100), "الإبصلمودية", self.open_tasbha_Window))
        asyncio.run(self.add_button_with_image(self.frame2, "Data\الصور\الفهرس.jpg", (239, 150, 100, 100), "الفهرس", self.open_elfhrs_window))
        asyncio.run(self.add_button_with_image(self.frame2, "Data\الصور\المدائح2.jpg", (352, 150, 100, 100), "المدائح", self.open_taranym_window))
        asyncio.run(self.add_button_with_image(self.frame2, "Data\الصور\الصليب القبطي.jpg", (465, 150, 100, 100), "المناسبات", self.open_elmonasbat_Window))

        self.frame2.show()

    def closeEvent(self, event):
        # Check if the copticdate is a Sunday
        if self.current_date.weekday() == 6:
            
            # Check if any PowerPoint application is open
            if self.is_powerpoint_open():
                self.show_error_message(f"PowerPoint is currently open. Please close the application and try again.")
                event.ignore()
                return

            try:
                self.replace_presentation()
            except Exception as e:
                self.show_error_message(e)
                event.ignore()
                return
        event.accept()

    def is_powerpoint_open(self):
        import pythoncom
        import win32com
        """Check if any PowerPoint application is open."""
        pythoncom.CoInitialize()
        try:
            powerpoint = win32com.client.GetActiveObject("PowerPoint.Application")
            if powerpoint.Presentations.Count > 0:
                # If there's any presentation open, PowerPoint is running
                return True
        except Exception:
            # If an exception is raised, PowerPoint is not open or no active instance is found
            return False
        finally:
            pythoncom.CoUninitialize()
        return False

    def checkCopticYear(self, copticYear):
        from commonFunctions import read_excel_cell, write_to_excel_cell
        currentYear = read_excel_cell(relative_path(r"Tables.xlsx"), "المناسبات", "M2")
        if copticYear != currentYear:
            from UpdateTable import a3yad, ElsomElkbyr, katamarsEl5amasyn
            asyncio.run(write_to_excel_cell(relative_path(r"Tables.xlsx"), "المناسبات", "M2", copticYear))
            a3yad()
            ElsomElkbyr()
            katamarsEl5amasyn()
        else:
            return


def load_initial_data(progress_callback):
    # Simulate a task with 5 steps
    for i in range(1, 6):
        time.sleep(0.5)  # simulate workload
        progress_callback(i * 20)  # 20%, 40%, ..., 100%

if __name__ == "__main__":
    app = QApplication(argv)

    splash = ModernSplashScreen()
    splash.show()

    def on_progress(val):
        splash.update_progress(val)

    def on_finished():
        window = MainWindow()
        window.show()
        splash.close()

    worker = WorkerThread(load_initial_data)
    worker.progress.connect(on_progress)
    worker.finished.connect(on_finished)
    worker.start()

    exit(app.exec_())