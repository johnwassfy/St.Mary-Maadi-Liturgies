import os
import sys
from PyQt5.QtWidgets import QApplication, QWidget, QLabel, QCheckBox, QVBoxLayout, QHBoxLayout, QComboBox, QLineEdit, QPushButton
from PyQt5.QtGui import QIcon
from PyQt5.QtCore import Qt, pyqtSignal
from pptx import Presentation
from pptx.enum.text import PP_ALIGN
from pptx.util import Pt

class CustomRow(QWidget):
    def __init__(self, label_text):
        super().__init__()
        
        self.label = QLabel(label_text)
        self.textbox = QLabel()
        self.combo_box = QComboBox()
        self.combo_box.addItems(["الاسقف", "المطران"])
        self.line_edit = QLineEdit()
        
        layout = QHBoxLayout()
        layout.addWidget(self.textbox)
        layout.addWidget(self.combo_box)
        layout.addWidget(self.line_edit)
        layout.addWidget(self.label)
        
        self.setLayout(layout)

class Bishop(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("حضور الأسقف")
        self.setGeometry(550, 300, 80, 50)
        
        # Create label and checkbox for the first row
        self.label1 = QLabel("حضور اسقف الابراشية ")
        self.textbox1 = QLabel()
        self.checkbox1 = QCheckBox()
        self.setWindowIcon(QIcon(self.relative_path(r"Data\الصور\Logo.ico")))
        
        # Create custom rows for the second and third rows
        self.row2 = CustomRow("حضور اسقف ضيف واحد")
        self.row3 = CustomRow("حضور اسقف ضيف ثاني")
        
        # Update button
        self.update_button = QPushButton("Save")
        self.update_button.clicked.connect(self.update_powerpoint)
        
        # Layout
        layout = QVBoxLayout()
        
        hbox1 = QHBoxLayout()
        hbox1.addStretch(1)
        hbox1.addWidget(self.checkbox1)
        hbox1.addWidget(self.label1)  # Set the same stretch factor for the label
        
        layout.addLayout(hbox1)
        layout.addWidget(self.row2)
        layout.addWidget(self.row3)
        layout.addWidget(self.update_button)
        
        self.setLayout(layout)

    def relative_path(self, relative_path):
        script_directory = os.path.dirname(os.path.abspath(__file__))
        absolute_path = os.path.join(script_directory, relative_path)
        return absolute_path

    def update_powerpoint(self):
        presentation = Presentation(self.relative_path(r"Data\CopyData\في حضور الاسقف و اساقفة ضيوف.pptx"))  # Provide your PowerPoint file path

        for slide in presentation.slides:
            slide_contains_word = False  # Flag to indicate if the slide contains "(الضيف)"
            slide_contains_word2 = False  # Flag to indicate if the slide contains "(الضيف2)"
            for shape in slide.shapes:
                if shape.has_text_frame:
                    text_frame = shape.text_frame
                    for paragraph in text_frame.paragraphs:
                        for run in paragraph.runs:
                            if "(الضيف)" in run.text:
                                slide_contains_word = True  # Set the flag if "(الضيف)" is found
                            if "(الضيف2)" in run.text:
                                slide_contains_word2 = True  # Set the flag if "(الضيف2)" is found

            if slide_contains_word:  # Proceed with replacement only if the slide contains "(الضيف)"
                for shape in slide.shapes:
                    if shape.has_text_frame:
                        text_frame = shape.text_frame
                        for paragraph in text_frame.paragraphs:
                            for run in paragraph.runs:
                                if "(الضيف)" in run.text and self.row2.line_edit.text():
                                    run.text = run.text.replace("(الضيف)", self.row2.line_edit.text())
                                if self.row2.combo_box.currentText() == "المطران":
                                    run.text = run.text.replace("الأسقف", self.row2.combo_box.currentText())
                                    run.text = run.text.replace("اسقفنا", "مطراننا")
                                    run.text = run.text.replace("ايبيسكوبوس", "متروبوليتيس")
                                    run.text = run.text.replace("إيبيسكوبوتيس", "متروبوليتيس")
                                    run.text = run.text.replace("n`epickopoc", "mmytropolityc")
                                    run.text = run.text.replace("epickopoc", "mmytropolityc")
                                    run.text = run.text.replace("`epickopou tyc", "mmytropolityc")

            if slide_contains_word2:  # Proceed with replacement only if the slide contains "(الضيف2)"
                for shape in slide.shapes:
                    if shape.has_text_frame:
                        text_frame = shape.text_frame
                        for paragraph in text_frame.paragraphs:
                            for run in paragraph.runs:
                                if "(الضيف2)" in run.text and self.row3.line_edit.text():
                                    run.text = run.text.replace("(الضيف2)", self.row3.line_edit.text())
                                if self.row2.combo_box.currentText() == "المطران":
                                    run.text = run.text.replace("الأسقف", self.row3.combo_box.currentText())
                                    run.text = run.text.replace("اسقفنا", "مطراننا")
                                    run.text = run.text.replace("ايبيسكوبوس", "متروبوليتيس")
                                    run.text = run.text.replace("إيبيسكوبوتيس", "متروبوليتيس")
                                    run.text = run.text.replace("n`epickopoc", "mmytropolityc")
                                    run.text = run.text.replace("epickopoc", "mmytropolityc")
                                    run.text = run.text.replace("`epickopou tyc", "mmytropolityc")
        
        presentation.save(self.relative_path(r"Data\حضور الأسقف.pptx"))
    


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = Bishop()
    window.show()
    sys.exit(app.exec_())
