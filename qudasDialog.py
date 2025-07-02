from PyQt5.QtWidgets import (QDialog, QPushButton, QVBoxLayout, QLabel, QFrame, QHBoxLayout,
                           QScrollArea, QWidget, QLineEdit, QSizePolicy, QTabWidget)
from PyQt5.QtGui import QFont, QPixmap, QColor
from PyQt5.QtCore import Qt, QSize
from commonFunctions import relative_path, open_presentation_relative_path
import pandas as pd
import win32com.client
import pythoncom
import qtawesome as qta
import os

class SectionSelectionDialog(QDialog):
    def __init__(self, parent=None, title="القداس", sheet_name=""):
        super().__init__(parent)
        self.selected_option = None
        
        self.setWindowTitle(title)
        
        # Make dialog modal - will be attached to parent window
        self.setWindowFlags(Qt.Dialog | Qt.FramelessWindowHint | Qt.WindowSystemMenuHint | Qt.WindowTitleHint)
        self.setModal(True)
        
        # Store parameters
        self.sheet_name = sheet_name
        self.title = title
        self.excel_path = relative_path(r"Files Data.xlsx")        # Adjust size if it's رفع بخور (make it taller)
        if sheet_name == "رفع بخور":
            base_height = 550
        else:
            base_height = 480
            
        # No need to adjust height based on title anymore since we keep it on one line
        self.setFixedSize(550, base_height)
        
        # Main layout
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(0)
        
        # Set gradient background for the entire dialog
        self.setStyleSheet("""
            QDialog {
                background: qlineargradient(
                    x1: 0, y1: 0, x2: 1, y2: 1,
                    stop: 0 rgba(15, 46, 71, 220),
                    stop: 0.6 rgba(30, 91, 138, 220),
                    stop: 1 rgba(45, 130, 180, 220)
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
        content_layout = QVBoxLayout(content_container)
        content_layout.setContentsMargins(15, 10, 15, 10)
        
        # Add search bar
        self.search_bar = QLineEdit()
        self.search_bar.setPlaceholderText("بحث")
        self.search_bar.setFixedHeight(40)
        self.search_bar.setLayoutDirection(Qt.RightToLeft)
        self.search_bar.setStyleSheet("""
            QLineEdit {
                text-align: center;
                border: 2px solid #c4c4c4;
                border-radius: 15px;
                padding: 5px 10px;
                background-color: rgba(255, 255, 255, 220);
                font-size: 16px;
                color: #333333;
            }
            QLineEdit:focus {
                border-color: #a0a0ff;
                background-color: #ffffff;
            }
        """)
        content_layout.addWidget(self.search_bar)
        content_layout.addSpacing(10)
        
        # Determine PowerPoint file path
        self.determine_file_path()
        
        # Special case for رفع بخور - create two scroll areas
        if sheet_name == "رفع بخور":
            # Create a tab widget to hold both sections
            self.tab_widget = QTabWidget()
            self.tab_widget.setLayoutDirection(Qt.RightToLeft)
            self.tab_widget.setStyleSheet("""
                QTabWidget::pane {
                    border: none;
                    background: transparent;
                }
                QTabWidget::tab-bar {
                    alignment: center;
                }
                QTabBar::tab {
                    background: rgba(255, 255, 255, 120);
                    color: white;
                    padding: 8px 16px;
                    margin: 2px;
                    border-radius: 5px;
                    font-size: 14px;
                    font-weight: bold;
                }
                QTabBar::tab:selected {
                    background: rgba(255, 255, 255, 180);
                    color: #0f2e47;
                }
                QTabBar::tab:hover:!selected {
                    background: rgba(255, 255, 255, 150);
                }
            """)
            
            # Tab 1: رفع بخور عشية و باكر
            tab1 = QWidget()
            tab1_layout = QVBoxLayout(tab1)
            tab1_layout.setContentsMargins(0, 10, 0, 0)
            
            # Create scroll area for عشية و باكر
            self.scroll_area1 = QScrollArea()
            self.scroll_area1.setWidgetResizable(True)
            self.scroll_area1.setStyleSheet("""
                QScrollArea {
                    background-color: transparent;
                    border: none;
                }
            """)
            self.scroll_area1.verticalScrollBar().setStyleSheet("""
                QScrollBar:vertical {
                    border: none;
                    background: transparent;
                    width: 10px;
                }
                QScrollBar::handle:vertical {
                    background: rgba(255, 255, 255, 100);
                    border-radius: 5px;
                }
                QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
                    background: none;
                }
                QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical {
                    background: none;
                }
            """)
            
            # Create content for first scroll area
            self.scroll_content1 = QWidget()
            self.buttons_layout1 = QVBoxLayout(self.scroll_content1)
            self.buttons_layout1.setAlignment(Qt.AlignTop)
            self.buttons_layout1.setSpacing(10)
            
            # Add loading indicator
            status_label1 = QLabel("جاري تحميل البيانات...")
            status_label1.setStyleSheet("color: white; font-size: 16px; font-weight: bold; background: transparent;")
            status_label1.setAlignment(Qt.AlignCenter)
            self.buttons_layout1.addWidget(status_label1)
            
            self.scroll_area1.setWidget(self.scroll_content1)
            self.scroll_area1.setLayoutDirection(Qt.RightToLeft)
            tab1_layout.addWidget(self.scroll_area1)
            
            # Tab 2: الذكصولوجيات
            tab2 = QWidget()
            tab2_layout = QVBoxLayout(tab2)
            tab2_layout.setContentsMargins(0, 10, 0, 0)
            
            # Create scroll area for الذكصولوجيات
            self.scroll_area2 = QScrollArea()
            self.scroll_area2.setWidgetResizable(True)
            self.scroll_area2.setStyleSheet("""
                QScrollArea {
                    background-color: transparent;
                    border: none;
                }
            """)
            self.scroll_area2.verticalScrollBar().setStyleSheet("""
                QScrollBar:vertical {
                    border: none;
                    background: transparent;
                    width: 10px;
                }
                QScrollBar::handle:vertical {
                    background: rgba(255, 255, 255, 100);
                    border-radius: 5px;
                }
                QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
                    background: none;
                }
                QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical {
                    background: none;
                }
            """)
            
            # Create content for second scroll area
            self.scroll_content2 = QWidget()
            self.buttons_layout2 = QVBoxLayout(self.scroll_content2)
            self.buttons_layout2.setAlignment(Qt.AlignTop)
            self.buttons_layout2.setSpacing(10)
            
            # Add loading indicator
            status_label2 = QLabel("جاري تحميل البيانات...")
            status_label2.setStyleSheet("color: white; font-size: 16px; font-weight: bold; background: transparent;")
            status_label2.setAlignment(Qt.AlignCenter)
            self.buttons_layout2.addWidget(status_label2)
            
            self.scroll_area2.setWidget(self.scroll_content2)
            self.scroll_area2.setLayoutDirection(Qt.RightToLeft)
            tab2_layout.addWidget(self.scroll_area2)
            
            # Add tabs to tab widget
            self.tab_widget.addTab(tab1, "عشية و باكر")
            self.tab_widget.addTab(tab2, "الذكصولوجيات")
            
            # Add tab widget to content layout
            content_layout.addWidget(self.tab_widget)
            
            # Define the zoksologyat file path
            self.zoks_file_path = relative_path(r"الذكصولوجيات.pptx")
            
            # Load data for both presentations after UI is initialized
            from PyQt5.QtCore import QTimer
            QTimer.singleShot(100, self.load_dual_presentations)
            
            # Connect search bar to filter both tab contents
            self.search_bar.textChanged.connect(self.filter_dual_buttons)
            
        else:
            # Standard single scroll area for other presentations
            self.scroll_area = QScrollArea()
            self.scroll_area.setWidgetResizable(True)
            self.scroll_area.setStyleSheet("""
                QScrollArea {
                    background-color: transparent;
                    border: none;
                }
            """)
            
            # Set stylesheet for scrollbar
            self.scroll_area.verticalScrollBar().setStyleSheet("""
                QScrollBar:vertical {
                    border: none;
                    background: transparent;
                    width: 10px;
                }
                QScrollBar:horizontal {
                    border: none;
                    background: transparent;
                    height: 10px;
                }
                QScrollBar::handle:vertical {
                    background: rgba(255, 255, 255, 100);
                    border-radius: 5px;
                }
                QScrollBar::add-line:vertical {
                    background: none;
                }
                QScrollBar::sub-line:vertical {
                    background: none;
                }
                QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical {
                    background: none;
                }
            """)
            
            # Create content for the scroll area
            self.scroll_content = QWidget()
            self.buttons_layout = QVBoxLayout(self.scroll_content)
            self.buttons_layout.setAlignment(Qt.AlignTop)
            self.buttons_layout.setSpacing(10)
            
            # Load data directly from PowerPoint
            status_label = QLabel("جاري تحميل البيانات...")
            status_label.setStyleSheet("color: white; font-size: 16px; font-weight: bold; background: transparent;")
            status_label.setAlignment(Qt.AlignCenter)
            self.buttons_layout.addWidget(status_label)
            
            # Set the scroll content
            self.scroll_area.setWidget(self.scroll_content)
            self.scroll_area.setLayoutDirection(Qt.RightToLeft)
            content_layout.addWidget(self.scroll_area)
            
            # Connect search bar to filter buttons
            self.search_bar.textChanged.connect(self.filter_buttons)
            
            # Load sections after UI is initialized
            from PyQt5.QtCore import QTimer
            QTimer.singleShot(100, self.extract_sections_from_powerpoint)
        
        # Add content container to main layout
        main_layout.addWidget(content_container, 1)
        
        # Set RTL layout
        self.setLayoutDirection(Qt.RightToLeft)
        
    def load_dual_presentations(self):
        """Load sections from both presentations for رفع بخور"""
        # Load the main presentation
        self.extract_sections_from_file(self.file_path, self.buttons_layout1, 1)
        
        # Load the zoksologyat presentation
        self.extract_sections_from_file(self.zoks_file_path, self.buttons_layout2, 2)
    
    def extract_sections_from_file(self, file_path, buttons_layout, file_index):
        """Extract sections from a PowerPoint file and add them to the specified layout"""
        # Clear existing buttons
        while buttons_layout.count():
            item = buttons_layout.takeAt(0)
            widget = item.widget()
            if widget:
                widget.deleteLater()
        
        try:
            pythoncom.CoInitialize()
            # First try to get an already open instance
            try:
                powerpoint = win32com.client.GetActiveObject("PowerPoint.Application")
            except:
                # If not open, create a new instance
                powerpoint = win32com.client.Dispatch("PowerPoint.Application")
            
            # Check if the presentation is already open
            is_open = False
            presentation = None
            for pres in powerpoint.Presentations:
                if os.path.abspath(pres.FullName.lower()) == os.path.abspath(file_path.lower()):
                    is_open = True
                    presentation = pres
                    break
            
            # If not open, open it
            if not is_open:
                presentation = powerpoint.Presentations.Open(file_path, WithWindow=False)
                just_opened = True
            else:
                just_opened = False
            
            # Dictionary to store sections: title -> slide index
            section_to_slide = {}
            slide_titles = []
            
            # Check if the presentation has sections
            has_sections = False
            section_titles = []
            
            try:
                # First try to get sections
                if presentation.SectionProperties.Count > 0:
                    has_sections = True
                    for i in range(1, presentation.SectionProperties.Count + 1):
                        section_name = presentation.SectionProperties.Name(i)
                        section_titles.append(section_name)
                        # Get first slide in this section
                        first_slide_index = presentation.SectionProperties.FirstSlide(i)
                        section_to_slide[section_name] = first_slide_index
            except:
                # If sections aren't available, fall back to slide titles
                has_sections = False
            
            # If no sections, use slide titles
            if not has_sections:
                for i in range(1, presentation.Slides.Count + 1):
                    slide = presentation.Slides.Item(i)
                    # Try to get slide title
                    title = ""
                    for shape in slide.Shapes:
                        if shape.HasTextFrame:
                            if shape.TextFrame.HasText:
                                text = shape.TextFrame.TextRange.Text
                                # Use the first non-empty text as title
                                if text and len(text.strip()) > 0:
                                    title = text.strip()
                                    break
                    
                    if title:
                        # Only add if title doesn't already exist
                        if title not in section_to_slide:
                            slide_titles.append(title)
                            section_to_slide[title] = i
            
            # Close the presentation if we just opened it
            if just_opened:
                presentation.Close()
            
            # Sort section titles alphabetically
            if has_sections:
                # Add buttons for each section
                for section in section_titles:
                    button = QPushButton(section)
                    self.set_button_style(button)
                    slide_index = section_to_slide[section]
                    # Store which file to navigate to
                    button.clicked.connect(lambda _, path=file_path, idx=slide_index: 
                                         self.go_to_slide(path, idx))
                    buttons_layout.addWidget(button)
            else:
                # If using slide titles, sort them alphabetically in Arabic
                import locale
                try:
                    locale.setlocale(locale.LC_ALL, 'ar_SY.UTF-8')  # Arabic locale
                except:
                    pass  # If Arabic locale not available, use default
                
                slide_titles.sort()
                
                # Add buttons for each slide title
                for title in slide_titles:
                    button = QPushButton(title)
                    self.set_button_style(button)
                    slide_index = section_to_slide[title]
                    # Store which file to navigate to
                    button.clicked.connect(lambda _, path=file_path, idx=slide_index: 
                                         self.go_to_slide(path, idx))
                    buttons_layout.addWidget(button)
            
            # Show error if no sections or titles found
            if not has_sections and not slide_titles:
                error_label = QLabel("لم يتم العثور على أقسام أو عناوين شرائح")
                error_label.setAlignment(Qt.AlignCenter)
                error_label.setStyleSheet("color: white; font-size: 16px; font-weight: bold; background: transparent;")
                buttons_layout.addWidget(error_label)
            
        except Exception as e:
            error_label = QLabel(f"خطأ في استخراج البيانات: {str(e)}")
            error_label.setStyleSheet("color: white; font-size: 14px;")
            error_label.setAlignment(Qt.AlignCenter)
            buttons_layout.addWidget(error_label)
            
        finally:
            pythoncom.CoUninitialize()
    
    def filter_dual_buttons(self):
        """Filter buttons in both tab areas based on search text"""
        search_text = self.normalize_text(self.search_bar.text().strip())
        
        # Filter tab 1 content
        visible_count1 = 0
        for i in range(self.buttons_layout1.count()):
            widget = self.buttons_layout1.itemAt(i).widget()
            if isinstance(widget, QPushButton):
                button_text = self.normalize_text(widget.text())
                is_visible = search_text in button_text
                widget.setVisible(is_visible)
                if is_visible:
                    visible_count1 += 1
        
        # Filter tab 2 content
        visible_count2 = 0
        for i in range(self.buttons_layout2.count()):
            widget = self.buttons_layout2.itemAt(i).widget()
            if isinstance(widget, QPushButton):
                button_text = self.normalize_text(widget.text())
                is_visible = search_text in button_text
                widget.setVisible(is_visible)
                if is_visible:
                    visible_count2 += 1
        
        # Show placeholder in tab 1 if needed
        if visible_count1 == 0:
            if not hasattr(self, 'placeholder_label1'):
                self.placeholder_label1 = QLabel("لا توجد نتائج")
                self.placeholder_label1.setAlignment(Qt.AlignCenter)
                self.placeholder_label1.setStyleSheet("color: white; font-size: 18px; font-weight: bold; background: transparent;")
                self.buttons_layout1.addWidget(self.placeholder_label1)
            self.placeholder_label1.setVisible(True)
        elif hasattr(self, 'placeholder_label1'):
            self.placeholder_label1.setVisible(False)
        
        # Show placeholder in tab 2 if needed
        if visible_count2 == 0:
            if not hasattr(self, 'placeholder_label2'):
                self.placeholder_label2 = QLabel("لا توجد نتائج")
                self.placeholder_label2.setAlignment(Qt.AlignCenter)
                self.placeholder_label2.setStyleSheet("color: white; font-size: 18px; font-weight: bold; background: transparent;")
                self.buttons_layout2.addWidget(self.placeholder_label2)
            self.placeholder_label2.setVisible(True)
        elif hasattr(self, 'placeholder_label2'):
            self.placeholder_label2.setVisible(False)
        
        # If there are results in one tab but not the other, switch to the tab with results
        if self.search_bar.text() and (visible_count1 > 0 and visible_count2 == 0):
            self.tab_widget.setCurrentIndex(0)  # Switch to tab 1
        elif self.search_bar.text() and (visible_count2 > 0 and visible_count1 == 0):
            self.tab_widget.setCurrentIndex(1)  # Switch to tab 2
    
    def extract_sections_from_powerpoint(self):
        """Extract sections and slide numbers directly from PowerPoint presentation"""
        # Clear existing buttons
        while self.buttons_layout.count():
            item = self.buttons_layout.takeAt(0)
            widget = item.widget()
            if widget:
                widget.deleteLater()
        
        try:
            pythoncom.CoInitialize()
            # First try to get an already open instance
            try:
                powerpoint = win32com.client.GetActiveObject("PowerPoint.Application")
            except:
                # If not open, create a new instance
                powerpoint = win32com.client.Dispatch("PowerPoint.Application")
            
            # Check if the presentation is already open
            is_open = False
            presentation = None
            for pres in powerpoint.Presentations:
                if os.path.abspath(pres.FullName.lower()) == os.path.abspath(self.file_path.lower()):
                    is_open = True
                    presentation = pres
                    break
            
            # If not open, open it
            if not is_open:
                presentation = powerpoint.Presentations.Open(self.file_path, WithWindow=False)
                just_opened = True
            else:
                just_opened = False
            
            # Dictionary to store sections: title -> slide index
            self.section_to_slide = {}
            slide_titles = []
            
            # Check if the presentation has sections
            has_sections = False
            section_titles = []
            
            try:
                # First try to get sections
                if presentation.SectionProperties.Count > 0:
                    has_sections = True
                    for i in range(1, presentation.SectionProperties.Count + 1):
                        section_name = presentation.SectionProperties.Name(i)
                        section_titles.append(section_name)
                        # Get first slide in this section
                        first_slide_index = presentation.SectionProperties.FirstSlide(i)
                        self.section_to_slide[section_name] = first_slide_index
            except:
                # If sections aren't available, fall back to slide titles
                has_sections = False
            
            # If no sections, use slide titles
            if not has_sections:
                for i in range(1, presentation.Slides.Count + 1):
                    slide = presentation.Slides.Item(i)
                    # Try to get slide title
                    title = ""
                    for shape in slide.Shapes:
                        if shape.HasTextFrame:
                            if shape.TextFrame.HasText:
                                text = shape.TextFrame.TextRange.Text
                                # Use the first non-empty text as title
                                if text and len(text.strip()) > 0:
                                    title = text.strip()
                                    break
                    
                    if title:
                        # Only add if title doesn't already exist
                        if title not in self.section_to_slide:
                            slide_titles.append(title)
                            self.section_to_slide[title] = i
            
            # Close the presentation if we just opened it
            if just_opened:
                presentation.Close()
            
            # Sort section titles alphabetically
            if has_sections:
                # Add buttons for each section
                for section in section_titles:
                    button = QPushButton(section)
                    self.set_button_style(button)
                    slide_index = self.section_to_slide[section]
                    button.clicked.connect(lambda _, idx=slide_index: self.go_to_slide(self.file_path, idx))
                    self.buttons_layout.addWidget(button)
            else:
                # If using slide titles, sort them alphabetically in Arabic
                import locale
                try:
                    locale.setlocale(locale.LC_ALL, 'ar_SY.UTF-8')  # Arabic locale
                except:
                    pass  # If Arabic locale not available, use default
                
                slide_titles.sort()
                
                # Add buttons for each slide title
                for title in slide_titles:
                    button = QPushButton(title)
                    self.set_button_style(button)
                    slide_index = self.section_to_slide[title]
                    button.clicked.connect(lambda _, idx=slide_index: self.go_to_slide(self.file_path, idx))
                    self.buttons_layout.addWidget(button)
            
            # Show error if no sections or titles found
            if not has_sections and not slide_titles:
                error_label = QLabel("لم يتم العثور على أقسام أو عناوين شرائح")
                error_label.setAlignment(Qt.AlignCenter)
                error_label.setStyleSheet("color: white; font-size: 16px; font-weight: bold; background: transparent;")
                self.buttons_layout.addWidget(error_label)
                
                # Try to fall back to Excel data if available
                try:
                    self.load_data_from_spreadsheet()
                except:
                    pass
            
        except Exception as e:
            error_label = QLabel(f"خطأ في استخراج البيانات: {str(e)}")
            error_label.setStyleSheet("color: white; font-size: 14px;")
            error_label.setAlignment(Qt.AlignCenter)
            self.buttons_layout.addWidget(error_label)
            
            # Try to fall back to Excel data if available
            try:
                self.load_data_from_spreadsheet()
            except:
                pass
        
        finally:
            pythoncom.CoUninitialize()
    
    def go_to_slide(self, file_path, slide_index):
        """Navigate to the specified PowerPoint slide in the specified file"""
        try:
            pythoncom.CoInitialize()
            try:
                powerpoint = win32com.client.GetActiveObject("PowerPoint.Application")
            except:
                powerpoint = win32com.client.Dispatch("PowerPoint.Application")
            
            # Set up a flag to check if the presentation is already open
            is_open = False
            presentation = None
            in_slideshow = False
    
            # Iterate over the currently open presentations
            for pres in powerpoint.Presentations:
                if os.path.abspath(pres.FullName.lower()) == os.path.abspath(file_path.lower()):
                    is_open = True
                    presentation = pres
                    # Check if presentation is in slideshow view
                    try:
                        # This will throw an exception if not in slideshow mode
                        slideshow_window = pres.SlideShowWindow
                        in_slideshow = True
                    except:
                        in_slideshow = False
                    break
            
            if not is_open:
                presentation = powerpoint.Presentations.Open(file_path)
                powerpoint.Visible = True
            
            # Handle navigation based on presentation state
            if in_slideshow:
                # If in slideshow view, use slideshow methods to navigate
                slideshow = presentation.SlideShowWindow.View
                slideshow.GotoSlide(slide_index)
            else:
                # If in normal view, use normal navigation methods
                slide = presentation.Slides(slide_index)
                slide.Select()
                # Make PowerPoint visible and activate the window
                powerpoint.Visible = True
                powerpoint.ActiveWindow.Activate()
            
        except Exception as e:
            print(f"Error navigating to slide: {str(e)}")
            
            # Try alternative method if the first fails
            try:
                # Alternative approach for older PowerPoint versions
                if presentation and not in_slideshow:
                    # Start slideshow from beginning
                    presentation.SlideShowSettings.Run()
                    # Wait briefly for slideshow to start
                    import time
                    time.sleep(0.5)
                    # Then navigate to the specific slide
                    slideshow = presentation.SlideShowWindow.View
                    slideshow.GotoSlide(slide_index)
            except Exception as e2:
                print(f"Alternative navigation also failed: {str(e2)}")
                
        finally:
            pythoncom.CoUninitialize()
    
    def create_header(self):        
        header = QFrame()
        
        # Use consistent header height
        header.setFixedHeight(60)  # Standard height for all titles
            
        header.setStyleSheet("""
            QFrame {
                background: qlineargradient(
                    x1: 0, y1: 0, x2: 1, y2: 0,
                    stop: 0 #0f2e47,
                    stop: 1 #1e5b8a
                );
                border-top-left-radius: 10px;
                border-top-right-radius: 10px;
            }
        """)        
        header_layout = QHBoxLayout(header)
        header_layout.setContentsMargins(10, 0, 10, 0)  # Minimize vertical space, reduce horizontal margins
        
        # Add icon
        try:
            icon_label = QLabel()
            icon = qta.icon("fa5s.book-open", color="white").pixmap(24, 24)
            icon_label.setPixmap(icon)
            icon_label.setStyleSheet("background: transparent;")
            header_layout.addWidget(icon_label)
            header_layout.addSpacing(10)
        except:
            pass        # Title
        title_label = QLabel(self.title)
        title_font = QFont()
        
        # Dynamically calculate font size based on title length
        # Use larger fonts for shorter titles, smaller for longer ones
        if len(self.title) > 70:
            title_font.setPointSize(9)
        elif len(self.title) > 60:
            title_font.setPointSize(10)
        elif len(self.title) > 50:
            title_font.setPointSize(11)
        elif len(self.title) > 40:
            title_font.setPointSize(12)
        elif len(self.title) > 30:
            title_font.setPointSize(13)
        elif len(self.title) > 20:
            title_font.setPointSize(14)
        else:
            title_font.setPointSize(16)
            
        title_font.setBold(True)
        title_label.setFont(title_font)
        title_label.setStyleSheet("color: white; background: transparent;")
        
        # Ensure text stays on one line and is properly aligned
        title_label.setWordWrap(False)
        title_label.setAlignment(Qt.AlignVCenter | Qt.AlignLeft)
        
        # Enable text elision with ... if it's too long
        title_label.setTextFormat(Qt.PlainText)
        
        # Give the title plenty of width space
        title_label.setMinimumWidth(350)
        title_label.setMaximumWidth(440)
        
        header_layout.addWidget(title_label)
        
        # Add stretch to push the close button to the right
        header_layout.addStretch()
        
        # Close button
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
        header_layout.addWidget(close_button)
        
        return header
    
    def extract_sections_from_powerpoint(self):
        """Extract sections and slide numbers directly from PowerPoint presentation"""
        # Clear existing buttons
        while self.buttons_layout.count():
            item = self.buttons_layout.takeAt(0)
            widget = item.widget()
            if widget:
                widget.deleteLater()
        
        try:
            pythoncom.CoInitialize()
            # First try to get an already open instance
            try:
                powerpoint = win32com.client.GetActiveObject("PowerPoint.Application")
            except:
                # If not open, create a new instance
                powerpoint = win32com.client.Dispatch("PowerPoint.Application")
            
            # Check if the presentation is already open
            is_open = False
            presentation = None
            for pres in powerpoint.Presentations:
                if os.path.abspath(pres.FullName.lower()) == os.path.abspath(self.file_path.lower()):
                    is_open = True
                    presentation = pres
                    break
            
            # If not open, open it
            if not is_open:
                presentation = powerpoint.Presentations.Open(self.file_path, WithWindow=False)
                just_opened = True
            else:
                just_opened = False
            
            # Dictionary to store sections: title -> slide index
            self.section_to_slide = {}
            slide_titles = []
            
            # Check if the presentation has sections
            has_sections = False
            section_titles = []
            
            try:
                # First try to get sections
                if presentation.SectionProperties.Count > 0:
                    has_sections = True
                    for i in range(1, presentation.SectionProperties.Count + 1):
                        section_name = presentation.SectionProperties.Name(i)
                        section_titles.append(section_name)
                        # Get first slide in this section
                        first_slide_index = presentation.SectionProperties.FirstSlide(i)
                        self.section_to_slide[section_name] = first_slide_index
            except:
                # If sections aren't available, fall back to slide titles
                has_sections = False
            
            # If no sections, use slide titles
            if not has_sections:
                for i in range(1, presentation.Slides.Count + 1):
                    slide = presentation.Slides.Item(i)
                    # Try to get slide title
                    title = ""
                    for shape in slide.Shapes:
                        if shape.HasTextFrame:
                            if shape.TextFrame.HasText:
                                text = shape.TextFrame.TextRange.Text
                                # Use the first non-empty text as title
                                if text and len(text.strip()) > 0:
                                    title = text.strip()
                                    break
                    
                    if title:
                        # Only add if title doesn't already exist
                        if title not in self.section_to_slide:
                            slide_titles.append(title)
                            self.section_to_slide[title] = i
            
            # Close the presentation if we just opened it
            if just_opened:
                presentation.Close()
            
            # Sort section titles alphabetically
            if has_sections:
                # Add buttons for each section
                for section in section_titles:
                    button = QPushButton(section)
                    self.set_button_style(button)
                    slide_index = self.section_to_slide[section]
                    button.clicked.connect(lambda _, idx=slide_index: self.go_to_slide(self.file_path, idx))
                    self.buttons_layout.addWidget(button)
            else:
                # If using slide titles, sort them alphabetically in Arabic
                import locale
                try:
                    locale.setlocale(locale.LC_ALL, 'ar_SY.UTF-8')  # Arabic locale
                except:
                    pass  # If Arabic locale not available, use default
                
                slide_titles.sort()
                
                # Add buttons for each slide title
                for title in slide_titles:
                    button = QPushButton(title)
                    self.set_button_style(button)
                    slide_index = self.section_to_slide[title]
                    button.clicked.connect(lambda _, idx=slide_index: self.go_to_slide(self.file_path, idx))
                    self.buttons_layout.addWidget(button)
            
            # Show error if no sections or titles found
            if not has_sections and not slide_titles:
                error_label = QLabel("لم يتم العثور على أقسام أو عناوين شرائح")
                error_label.setAlignment(Qt.AlignCenter)
                error_label.setStyleSheet("color: white; font-size: 16px; font-weight: bold; background: transparent;")
                self.buttons_layout.addWidget(error_label)
                
                # Try to fall back to Excel data if available
                try:
                    self.load_data_from_spreadsheet()
                except:
                    pass
            
        except Exception as e:
            error_label = QLabel(f"خطأ في استخراج البيانات: {str(e)}")
            error_label.setStyleSheet("color: white; font-size: 14px;")
            error_label.setAlignment(Qt.AlignCenter)
            self.buttons_layout.addWidget(error_label)
            
            # Try to fall back to Excel data if available
            try:
                self.load_data_from_spreadsheet()
            except:
                pass
        
        finally:
            pythoncom.CoUninitialize()
    
    def load_data_from_spreadsheet(self):
        """Fallback method to load data from Excel spreadsheet"""
        try:
            # Load the Excel file
            df = pd.read_excel(self.excel_path, sheet_name=self.sheet_name)
            
            # Get the section names and indices
            column_data = df.iloc[:, 0].tolist()  # First column contains section names
            index_data = df.iloc[:, 2].tolist()   # Third column contains slide indices
            
            # Create a dictionary mapping item names to their indices
            self.section_to_slide = {item: idx for item, idx in zip(column_data, index_data) if not pd.isna(item)}
            
            # Add buttons for each section
            for item in column_data:
                if pd.isna(item):
                    continue
                button = QPushButton(item)
                self.set_button_style(button)
                button.clicked.connect(lambda _, i=self.section_to_slide[item]: self.go_to_slide(self.file_path, i))
                self.buttons_layout.addWidget(button)
                
        except Exception as e:
            error_label = QLabel(f"لم يتم العثور على البيانات: {str(e)}")
            error_label.setStyleSheet("color: white; font-size: 14px;")
            self.buttons_layout.addWidget(error_label)
    
    def determine_file_path(self):
        """Determine which PowerPoint file to use based on the sheet name"""
        match(self.sheet_name):
            case "رفع بخور":
                self.file_path = r"رفع بخور عشية و باكر.pptx"
            case "القداس":
                self.file_path = r"قداس.pptx"
            case "قداس الطفل":
                self.file_path = r"قداس الطفل.pptx"
            case "التسبحة":
                self.file_path = r"الإبصلمودية.pptx"
            case "تسبحة كيهك":
                self.file_path = r"الإبصلمودية الكيهكية.pptx"
            case "الذكصولوجيات":
                self.file_path = r"الذكصولوجيات.pptx"
            case "المدائح":
                self.file_path = r"كتاب المدائح.pptx"
            case _:
                self.file_path = r"قداس.pptx"  # Default
        
        self.file_path = relative_path(self.file_path)
    
    def set_button_style(self, button):
        button.setStyleSheet("""
            QPushButton {
                background-color: rgba(255, 255, 255, 200);
                border: none;
                border-radius: 12px;
                color: #0f2e47;
                padding: 10px;
                font-size: 16px;
                font-weight: bold;
                text-align: right;
                min-height: 40px;
            }
            QPushButton:hover {
                background-color: rgba(255, 255, 255, 230);
                color: #1e5b8a;
                border: 1px solid rgba(255, 255, 255, 50);
            }
            QPushButton:pressed {
                background-color: rgba(200, 200, 200, 250);
                padding-top: 11px;
                padding-bottom: 9px;
            }
        """)
        button.setLayoutDirection(Qt.RightToLeft)
    
    def filter_buttons(self):
        """Filter buttons based on search text"""
        search_text = self.normalize_text(self.search_bar.text())
        visible_count = 0
        
        # Iterate through buttons in the layout
        for i in range(self.buttons_layout.count()):
            widget = self.buttons_layout.itemAt(i).widget()
            if isinstance(widget, QPushButton):
                button_text = self.normalize_text(widget.text())
                is_visible = search_text in button_text
                widget.setVisible(is_visible)
                if is_visible:
                    visible_count += 1
        
        # If no buttons are visible, show a placeholder
        if visible_count == 0:
            if not hasattr(self, 'placeholder_label'):
                self.placeholder_label = QLabel("لا توجد نتائج")
                self.placeholder_label.setAlignment(Qt.AlignCenter)
                self.placeholder_label.setStyleSheet("color: white; font-size: 18px; font-weight: bold; background: transparent;")
                self.buttons_layout.addWidget(self.placeholder_label)
            self.placeholder_label.setVisible(True)
        else:
            if hasattr(self, 'placeholder_label'):
                self.placeholder_label.setVisible(False)
    
    def normalize_text(self, text):
        """Normalize text by replacing alif with hamza."""
        return text.replace('أ', 'ا').replace('ؤ', 'و').replace('إ', 'ا').replace('آ', 'ا').lower()
    
    def mousePressEvent(self, event):
        # Allow dragging the frameless window
        if event.button() == Qt.LeftButton and event.y() < 60:  # 60 is header height
            self._drag_pos = event.globalPos() - self.frameGeometry().topLeft()
            event.accept()

    def mouseMoveEvent(self, event):
        # Move the window with mouse
        if event.buttons() == Qt.LeftButton and hasattr(self, '_drag_pos'):
            self.move(event.globalPos() - self._drag_pos)
            event.accept()