import os
from PyQt5.QtWidgets import QLabel, QVBoxLayout, QDialog, QSpinBox, QPushButton, QMessageBox, QCalendarWidget, QHBoxLayout, QComboBox
from PyQt5.QtGui import QIcon
from PyQt5.QtCore import Qt, QDateTime, pyqtSignal, QLocale
from openpyxl import load_workbook
from datetime import datetime
from copticDate import CopticCalendar  # Import the CopticCalendar class
from commonFunctions import relative_path  # Import the function to load background image

class ChangeDate(QDialog):
    date_updated = pyqtSignal(str, str)  # Define a signal to emit the updated date and time

    def __init__(self, CurrentDate, CurrentTime):
        super().__init__()

        try:
            # Set window title and icon
            self.setWindowTitle("إختيار التاريخ")
            self.setWindowIcon(QIcon(relative_path(r"Data\الصور\Logo.ico")))

            # Load options from Excel and set up ComboBox
            self.load_options_from_excel()
            self.excel_dropdown = QComboBox()
            self.excel_dropdown.setStyleSheet("QComboBox { font-size: 14px; }")  # Set font size
            
            # Add a placeholder value to the dropdown menu
            self.excel_dropdown.addItem("المناسبات")  # Placeholder value
            self.excel_dropdown.addItems(self.options)  # Add the actual options
            self.excel_dropdown.setCurrentIndex(0)  # Set the placeholder as the default selected value

            # Create submit and current date/time buttons
            self.submit_button = QPushButton("تعين")
            self.submit_button.clicked.connect(self.submit_date_time)
            self.current_date_time_button = QPushButton("مبـــاشر")
            self.current_date_time_button.clicked.connect(self.show_current_date_time)

            # Create QLabel for displaying Coptic date
            self.coptic_date_label = QLabel()

            # Create CopticCalendar instance
            self.coptic_calendar = CopticCalendar()

            # Create the main layout
            layout = QVBoxLayout()
            layout.addWidget(QLabel("اختر التاريخ:"))

            # Create and add the calendar widget directly to the layout
            self.calendar_widget = QCalendarWidget()
            self.calendar_widget.setGridVisible(True)
            self.calendar_widget.setVerticalHeaderFormat(QCalendarWidget.NoVerticalHeader)  # Hide week numbers
            self.calendar_widget.setLocale(QLocale(QLocale.Arabic, QLocale.Egypt))  # Set Arabic locale for the calendar

            # Set the calendar to open on the provided CurrentDate
            self.calendar_widget.setSelectedDate(CurrentDate)
            self.calendar_widget.clicked.connect(self.update_coptic_date_label)  # Update date on calendar click
            layout.addWidget(self.calendar_widget)

            layout.addWidget(QLabel("اختر الوقت (اليوم القبطي ينتهي 5:30 مساء)"))

            # Create layout for time selection
            time_layout = QHBoxLayout()

            # Create SpinBoxes for hours and minutes
            self.hour_spin = QSpinBox()
            self.hour_spin.setRange(0, 13)  # 1-12 for 12-hour format
            self.hour_spin.setValue(int(CurrentTime.split(':')[0]) % 12 or 12)  # Default value
            self.hour_spin.valueChanged.connect(self.handle_hour_change)  # Connect to update AM/PM

            self.minute_spin = QSpinBox()
            self.minute_spin.setRange(0, 60)  # 0 for 00 and 1 for 30
            self.minute_spin.setValue(int(CurrentTime.split(':')[1].split()[0]))  # Default value
            self.minute_spin.valueChanged.connect(self.handle_minute_change)

            # Determine if the current time is AM or PM
            if "AM" in CurrentTime:
                self.ampm_combo = QComboBox()
                self.ampm_combo.addItems(["AM", "PM"])  # Add AM and PM options
                self.ampm_combo.setCurrentIndex(0)  # Default to AM
            else:
                self.ampm_combo = QComboBox()
                self.ampm_combo.addItems(["AM", "PM"])  # Add AM and PM options
                self.ampm_combo.setCurrentIndex(1)  # Default to PM

            # Add SpinBoxes to time layout
            time_layout.addWidget(self.hour_spin)
            time_layout.addWidget(self.minute_spin)
            time_layout.addWidget(self.ampm_combo)

            layout.addLayout(time_layout)
            layout.addWidget(QLabel("اختار المنــاسبة"))
            layout.addWidget(self.excel_dropdown)
            layout.addWidget(self.current_date_time_button)
            layout.addWidget(self.coptic_date_label)
            layout.addWidget(self.submit_button)

            self.setLayout(layout)

            # Connect ComboBox signal
            self.excel_dropdown.currentIndexChanged.connect(self.update_date_and_time_from_coptic)
            self.update_coptic_date_label()
        except Exception as e:
            self.show_error_message(str(e))

    def load_options_from_excel(self):
        try:
            # Load options from Excel
            self.options = []
            workbook_path = relative_path(r"Tables.xlsx")
            wb = load_workbook(workbook_path)
            ws = wb["المناسبات"]
            for cell in ws['G'][1:]:
                if cell.value:
                    self.options.append(cell.value)
        except Exception as e:
            self.show_error_message(str(e))

    def show_current_date_time(self):
        try:
            # Get current date and time
            current_date_time = QDateTime.currentDateTime()
            formatted_date_time = current_date_time.toString(Qt.ISODate)
            selected_date = QDateTime.fromString(formatted_date_time.split("T")[0], Qt.ISODate).date()
            self.calendar_widget.setSelectedDate(selected_date)  # Set the current date in the calendar
            self.hour_spin.setValue(current_date_time.time().hour() % 12 or 12)  # Convert to 12-hour format
            self.minute_spin.setValue(current_date_time.time().minute())
            self.ampm_combo.setCurrentIndex(0 if current_date_time.time().hour() < 12 else 1)  # 0 for AM, 1 for PM
            self.update_coptic_date_label()
        except Exception as e:
            self.show_error_message(str(e))

    def handle_hour_change(self, value):
        """Handle the hour spin box value change."""
        if value == 12:  # Check if the hour is 12
            # Change AM/PM based on current selection
            if self.ampm_combo.currentText() == "AM":
                self.ampm_combo.setCurrentIndex(1)  # Change to PM
            else:
                self.ampm_combo.setCurrentIndex(0)  # Change to AM

        # After reaching 12, reset to 1
        if value >= 13:  # This will cover cases for 12 and 1
            self.hour_spin.setValue(1)  # Reset to 1

        if value == 0:  # This will cover the case when value is reset to 0
            self.hour_spin.setValue(12)  # Reset to 12

        self.update_coptic_date_label()

    def handle_minute_change(self, value):
        if value >= 60:
            self.hour_spin.setValue(0)
            
        self.update_coptic_date_label()

    def submit_date_time(self):
        try:
            # Get the entered date and time
            date = self.calendar_widget.selectedDate().toString("yyyy-MM-dd")  # Get the selected date
            hour = self.hour_spin.value()
            minute = self.minute_spin.value()  # Convert to 00 or 30
            ampm = self.ampm_combo.currentText()  # Get selected AM/PM text

            time = f"{hour}:{minute:02d} {ampm}"

            if not date.strip() and not time.strip():
                self.close()
            else:
                self.date_updated.emit(date, time)
        except Exception as e:
            self.show_error_message(str(e))

    def update_date_and_time_from_coptic(self, index):
        try:
            selected_value = self.excel_dropdown.itemText(index)
            wb = load_workbook(relative_path(r"Tables.xlsx"))
            ws = wb["المناسبات"]
            for row in range(2, ws.max_row + 1):
                if ws.cell(row=row, column=7).value == selected_value:
                    coptic_month = ws.cell(row=row, column=4).value
                    coptic_day = ws.cell(row=row, column=6).value
                    gregorian_date = self.coptic_calendar.coptic_to_gregorian([self.coptic_calendar.current_coptic_datetime[0], coptic_month, coptic_day])
                    self.calendar_widget.setSelectedDate(gregorian_date)  # Update the calendar to show this date
                    self.hour_spin.setValue(12)  # Reset to 12
                    self.minute_spin.setValue(0)  # Reset to 00
                    self.ampm_combo.setCurrentIndex(1)  # Reset to PM
                    break
            self.update_coptic_date_label()
        except Exception as e:
            self.show_error_message(str(e))

    def update_coptic_date_label(self):
        try:
            selected_date = self.calendar_widget.selectedDate()
            hour = self.hour_spin.value()
            minute = self.minute_spin.value()
            ampm = self.ampm_combo.currentText()  # Get selected AM/PM text

            # Convert hour from 12-hour format to 24-hour format
            if ampm == "PM" and hour != 12:
                hour += 12
            elif ampm == "AM" and hour == 12:
                hour = 0

            if selected_date:
                # Convert selected date to datetime object
                gregorian_date = datetime(selected_date.year(), selected_date.month(), selected_date.day(), hour, minute)
                # Convert Gregorian date to Coptic date
                coptic_date = self.coptic_calendar.gregorian_to_coptic(gregorian_date)
                coptic_date_str = f"{coptic_date[2]:02}-{self.coptic_calendar.coptic_month_name(coptic_date[1]):02}-{coptic_date[0]:04}"
                self.coptic_date_label.setText(f"التاريخ القبطي: {coptic_date_str}")
        except Exception as e:
            self.show_error_message(str(e))

    def show_error_message(self, error_message):
        QMessageBox.critical(self, "Error", error_message)

        
# import sys
# from PyQt5.QtWidgets import QApplication

# Ensure the ChangeDate class code is already defined or imported here

# if __name__ == "__main__":
#     app = QApplication(sys.argv)

#     # Define initial date and time to display
#     initial_date = "2024-10-29"  # Example initial date
#     initial_time = "12:00 PM"    # Example initial time

#     # Create and show ChangeDate dialog
#     change_date_dialog = ChangeDate(initial_date, initial_time)
    
#     # Show the dialog
#     change_date_dialog.exec_()

#     sys.exit(app.exec_())
