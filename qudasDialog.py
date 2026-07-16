from PyQt5.QtWidgets import (QDialog, QPushButton, QVBoxLayout, QLabel, QFrame, QHBoxLayout,
                           QScrollArea, QWidget, QLineEdit, QSizePolicy, QTabWidget)
from PyQt5.QtGui import QFont, QPixmap, QColor
from PyQt5.QtCore import Qt, QSize, QTimer
from commonFunctions import relative_path, open_presentation_relative_path, get_open_presentations
import pandas as pd
import win32com.client
import pythoncom
import qtawesome as qta
import os
import re

class SectionSelectionDialog(QDialog):
    _sections_cache = {}
    _dialog_cache = {}

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
        if sheet_name in ("رفع بخور", "التسبحة", "تسبحة كيهك"):
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

        # Start monitor timer to auto-close dialog if PowerPoint closes
        self._ppt_monitor_timer = QTimer(self)
        self._ppt_monitor_timer.setSingleShot(False)
        self._ppt_monitor_timer.timeout.connect(self.check_presentation_still_open)
        self._ppt_monitor_timer.start(1000)  # Check every 1 second
        
        # Special case for رفع بخور and التسبحة and تسبحة كيهك - create two scroll areas
        if sheet_name in ("رفع بخور", "التسبحة", "تسبحة كيهك"):
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
                    min-width: 200px;
                    min-height: 20px;
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
            tab1_label = "عشية و باكر" if sheet_name == "رفع بخور" else "الإبصلمودية"
            self.tab_widget.addTab(tab1, tab1_label)
            self.tab_widget.addTab(tab2, "الذكصولوجيات")
            
            # Add tab widget to content layout
            content_layout.addWidget(self.tab_widget)
            
            # Define the zoksologyat file path
            self.zoks_file_path = relative_path(r"الذكصولوجيات.pptx")
            
            # Load data for both presentations after UI is initialized
            self._dual_load_timer = QTimer(self)
            self._dual_load_timer.setSingleShot(True)
            self._dual_load_timer.timeout.connect(self.load_dual_presentations)
            self._dual_load_timer.start(100)
            
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
            self._load_sections_timer = QTimer(self)
            self._load_sections_timer.setSingleShot(True)
            self._load_sections_timer.timeout.connect(self.extract_sections_from_powerpoint)
            self._load_sections_timer.start(100)
        
        # Add content container to main layout
        main_layout.addWidget(content_container, 1)
        
        # Set RTL layout
        self.setLayoutDirection(Qt.RightToLeft)

    @classmethod
    def get_dialog(cls, parent, title, sheet_name):
        key = (sheet_name or "")
        dialog = cls._dialog_cache.get(key)
        if dialog is None:
            dialog = cls(parent, title, sheet_name)
            cls._dialog_cache[key] = dialog
        else:
            dialog.setParent(parent)
            dialog.sheet_name = sheet_name
            dialog.set_dialog_title(title)
        return dialog

    def set_dialog_title(self, title):
        self.title = title
        self.setWindowTitle(title)
        label = getattr(self, "_title_label", None)
        if label is not None:
            label.setText(self._compute_title_text())

    def _remove_from_cache(self):
        key = (self.sheet_name or "")
        cached = self._dialog_cache.get(key)
        if cached is self:
            del self._dialog_cache[key]

    def _close_due_to_ppt_close(self):
        self._force_close = True
        self.close()

    def _notify_parent_refresh(self):
        parent = self.parentWidget()
        if parent is not None and hasattr(parent, "refresh_button_states"):
            try:
                parent.refresh_button_states(skip_timer=True)
            except Exception:
                pass

    def _compute_title_text(self):
        title_text = self.title
        if getattr(self, "source_button_label", None):
            title_text = f"{self.source_button_label} - {self.title}"
        return title_text

    @classmethod
    def _cache_key(cls, file_path):
        if not file_path:
            return None
        return os.path.abspath(file_path).lower()

    @classmethod
    def _get_cached_sections(cls, file_path):
        key = cls._cache_key(file_path)
        if not key:
            return None
        return cls._sections_cache.get(key)

    @classmethod
    def _set_cached_sections(cls, file_path, items):
        if not items:
            return
        key = cls._cache_key(file_path)
        if not key:
            return
        cls._sections_cache[key] = items

    def _render_section_buttons(self, buttons_layout, items, file_path):
        for title, slide_index in items:
            button = QPushButton(title)
            self.set_button_style(button)
            button.clicked.connect(lambda _, path=file_path, idx=slide_index: self.go_to_slide(path, idx))
            buttons_layout.addWidget(button)
        
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

        cached = self._get_cached_sections(file_path)
        if cached:
            self._render_section_buttons(buttons_layout, cached, file_path)
            return
        
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
            
            # Build list for caching and rendering
            if has_sections:
                items = [(section, section_to_slide[section]) for section in section_titles]
            else:
                import locale
                try:
                    locale.setlocale(locale.LC_ALL, 'ar_SY.UTF-8')
                except:
                    pass
                slide_titles.sort()
                items = [(title, section_to_slide[title]) for title in slide_titles]

            self._set_cached_sections(file_path, items)
            self._render_section_buttons(buttons_layout, items, file_path)
            
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

        cached = self._get_cached_sections(self.file_path)
        if cached:
            self._render_section_buttons(self.buttons_layout, cached, self.file_path)
            return
        
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
            
            if has_sections:
                items = [(section, self.section_to_slide[section]) for section in section_titles]
            else:
                import locale
                try:
                    locale.setlocale(locale.LC_ALL, 'ar_SY.UTF-8')
                except:
                    pass
                slide_titles.sort()
                items = [(title, self.section_to_slide[title]) for title in slide_titles]

            self._set_cached_sections(self.file_path, items)
            self._render_section_buttons(self.buttons_layout, items, self.file_path)
            
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
        title_label = QLabel(self._compute_title_text())
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

        self._title_label = title_label
        
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

    def reject(self):
        self.hide()
        self._notify_parent_refresh()
        super().reject()

    def showEvent(self, event):
        super().showEvent(event)
        timer = getattr(self, "_ppt_monitor_timer", None)
        if timer is not None and not timer.isActive():
            timer.start(1000)
    
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

    def closeEvent(self, event):
        for timer_name in ("_dual_load_timer", "_load_sections_timer", "_render_sections_timer", "_ppt_monitor_timer"):
            timer = getattr(self, timer_name, None)
            if timer is not None:
                timer.stop()
        self._notify_parent_refresh()
        if getattr(self, "_force_close", False):
            self._remove_from_cache()
            self._force_close = False
        super().closeEvent(event)

    def check_presentation_still_open(self):
        """Check if the associated PowerPoint presentation is still open; close dialog if not.

        Uses the centralized `get_open_presentations()` helper from commonFunctions to
        determine which presentations are currently open.
        """
        try:
            if not getattr(self, 'file_path', None):
                return

            target_path = os.path.abspath(self.file_path).lower()
            open_list = get_open_presentations() or []
            normalized = [os.path.abspath(p).lower() for p in open_list if p]
            if target_path not in normalized:
                try:
                    self._close_due_to_ppt_close()
                    self._notify_parent_refresh()
                except Exception:
                    pass
        except Exception:
            # Silent on purpose
            pass
    
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


class Elbas5aSectionSelectionDialog(SectionSelectionDialog):
    def __init__(
        self,
        parent=None,
        title="أسبوع الآلام",
        presentation_path="",
        filter_mode="",
        filter_key="",
        filter_keywords=None,
        always_include_section_ids=None,
        source_button_id="",
        source_button_label="",
    ):
        QDialog.__init__(self, parent)
        self.selected_option = None
        self.title = title
        self.sheet_name = ""
        self.excel_path = relative_path(r"Files Data.xlsx")
        self.presentation_path = os.path.abspath(presentation_path) if presentation_path else ""
        self.filter_mode = filter_mode or ""
        self.filter_key = filter_key or ""
        self.filter_keywords = filter_keywords or []
        self.always_include_section_ids = always_include_section_ids or []
        self.source_button_id = source_button_id
        self.source_button_label = source_button_label
        self.section_buttons = []
        self.section_records = []
        self.filtered_records = []
        self.active_filter_mode = "all"
        self.active_keyword = ""
        self.filter_buttons_row = []
        self.setWindowFlags(Qt.Dialog | Qt.FramelessWindowHint | Qt.WindowSystemMenuHint | Qt.WindowTitleHint)
        self.setModal(True)
        self.determine_file_path()
        # Start monitor timer to auto-close dialog if PowerPoint closes
        self._ppt_monitor_timer = QTimer(self)
        self._ppt_monitor_timer.setSingleShot(False)
        self._ppt_monitor_timer.timeout.connect(self.check_presentation_still_open)
        self._ppt_monitor_timer.start(1000)  # Check every 1 second
        
        self._setup_holyweek_ui()

    def _setup_holyweek_ui(self):
        self.setWindowTitle(self.title)
        self.setFixedSize(550, 480)
        self._position_like_parent()
        self.setLayoutDirection(Qt.RightToLeft)
        self.setStyleSheet(
            """
            QDialog {
                background: qlineargradient(
                    x1: 0, y1: 0, x2: 1, y2: 1,
                    stop: 0 rgba(107, 6, 6, 245),
                    stop: 0.6 rgba(140, 30, 30, 245),
                    stop: 1 rgba(180, 80, 80, 245)
                );
                border-radius: 10px;
                border: 1px solid rgba(200, 200, 200, 150);
            }
            """
        )

        new_main_layout = QVBoxLayout(self)
        new_main_layout.setContentsMargins(0, 0, 0, 0)
        new_main_layout.setSpacing(0)

        header = self.create_header()
        new_main_layout.addWidget(header)

        content_container = QFrame()
        content_container.setStyleSheet("background: transparent; border: none;")
        content_layout = QVBoxLayout(content_container)
        content_layout.setContentsMargins(15, 10, 15, 10)
        content_layout.setSpacing(10)

        self.search_bar = QLineEdit()
        self.search_bar.setPlaceholderText("بحث")
        self.search_bar.setFixedHeight(40)
        self.search_bar.setLayoutDirection(Qt.RightToLeft)
        self.search_bar.setStyleSheet(
            """
            QLineEdit {
                text-align: center;
                border: 2px solid rgba(255, 255, 255, 120);
                border-radius: 15px;
                padding: 5px 10px;
                background-color: rgba(255, 255, 255, 220);
                font-size: 16px;
                color: #3a0000;
            }
            QLineEdit:focus {
                border-color: #ffd4d4;
                background-color: #ffffff;
            }
            """
        )
        self.search_bar.textChanged.connect(self.filter_buttons)
        content_layout.addWidget(self.search_bar)

        self.filter_buttons_container = QFrame()
        self.filter_buttons_container.setStyleSheet("background: transparent; border: none;")
        self.filter_buttons_layout = QHBoxLayout(self.filter_buttons_container)
        self.filter_buttons_layout.setContentsMargins(0, 0, 0, 0)
        self.filter_buttons_layout.setSpacing(6)
        self.filter_buttons_layout.setAlignment(Qt.AlignCenter)
        content_layout.addWidget(self.filter_buttons_container)

        self.build_filter_buttons()

        self.status_label = QLabel("")
        self.status_label.setAlignment(Qt.AlignCenter)
        self.status_label.setStyleSheet("color: #ffe8a3; font-size: 13px; font-weight: bold; background: transparent;")
        self.status_label.hide()
        content_layout.addWidget(self.status_label)

        self.scroll_area = QScrollArea()
        self.scroll_area.setWidgetResizable(True)
        self.scroll_area.setStyleSheet(
            """
            QScrollArea {
                background-color: transparent;
                border: none;
            }
            """
        )
        self.scroll_area.verticalScrollBar().setStyleSheet(
            """
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
            """
        )

        self.scroll_content = QWidget()
        self.buttons_layout = QVBoxLayout(self.scroll_content)
        self.buttons_layout.setAlignment(Qt.AlignTop)
        self.buttons_layout.setSpacing(10)
        self.scroll_area.setWidget(self.scroll_content)
        content_layout.addWidget(self.scroll_area, 1)

        new_main_layout.addWidget(content_container, 1)

        self.show_loading_message("جاري تحميل الأقسام...")

        self._render_sections_timer = QTimer(self)
        self._render_sections_timer.setSingleShot(True)
        self._render_sections_timer.timeout.connect(self.extract_and_render_sections)
        self._render_sections_timer.start(50)


    def create_header(self):
        header = QFrame()
        header.setFixedHeight(50)
        header.setStyleSheet(
            """
            QFrame {
                background: qlineargradient(
                    x1: 0, y1: 0, x2: 1, y2: 0,
                    stop: 0 #8b0000,
                    stop: 1 #c03232
                );
                border-top-left-radius: 10px;
                border-top-right-radius: 10px;
            }
            """
        )

        header_layout = QHBoxLayout(header)
        header_layout.setContentsMargins(15, 0, 15, 0)

        title_layout = QHBoxLayout()
        try:
            icon_label = QLabel()
            icon = qta.icon("fa5s.list", color="white").pixmap(18, 18)
            icon_label.setPixmap(icon)
            icon_label.setStyleSheet("background: transparent;")
            title_layout.addWidget(icon_label)
            title_layout.addSpacing(8)
        except:
            pass

        title_text = self.title
        if self.source_button_label:
            title_text = f"{self.source_button_label} - {self.title}"
        title_label = QLabel(title_text)
        title_font = QFont()
        title_font.setPointSize(13)
        title_font.setBold(True)
        title_label.setFont(title_font)
        title_label.setStyleSheet("color: white; background: transparent;")
        title_layout.addWidget(title_label)

        back_button = QPushButton("العودة")
        back_button.setCursor(Qt.PointingHandCursor)
        back_button.setFixedHeight(34)
        back_button.setStyleSheet(
            """
            QPushButton {
                background-color: rgba(255, 255, 255, 205);
                border: none;
                border-radius: 12px;
                color: #3a0000;
                font-size: 13px;
                font-weight: bold;
                min-width: 90px;
                padding: 6px 12px;
            }
            QPushButton:hover {
                background-color: rgba(255, 240, 240, 240);
            }
            """
        )
        back_button.clicked.connect(self.reject)

        header_layout.addLayout(title_layout)
        header_layout.addStretch()
        header_layout.addWidget(back_button)
        return header

    def _position_like_parent(self):
        parent = self.parentWidget()
        if parent is not None:
            self.move(parent.frameGeometry().topLeft())

    def build_filter_buttons(self):
        while self.filter_buttons_layout.count():
            item = self.filter_buttons_layout.takeAt(0)
            widget = item.widget()
            if widget:
                widget.deleteLater()
        self.filter_buttons_row = []

        mode = (self.filter_mode or "").strip().lower()
        if mode == "by-key" and self.filter_key:
            all_button = self._create_filter_button("كل الأقسام", "all")
            key_button = self._create_filter_button(self.filter_key, "by-key")
            self.filter_buttons_layout.addWidget(all_button)
            self.filter_buttons_layout.addWidget(key_button)
            self.filter_buttons_row.extend([all_button, key_button])
        elif mode == "by-keywords" and self.filter_keywords:
            all_button = self._create_filter_button("كل الأقسام", "clear-search")
            self.filter_buttons_layout.addWidget(all_button)
            self.filter_buttons_row.append(all_button)
            for keyword in self.filter_keywords:
                if not keyword:
                    continue
                keyword_button = self._create_filter_button(str(keyword), "search-keyword", str(keyword))
                self.filter_buttons_layout.addWidget(keyword_button)
                self.filter_buttons_row.append(keyword_button)

        self._refresh_filter_button_styles()

    def _create_filter_button(self, text, mode_value, keyword_value=""):
        button = QPushButton(text)
        button.setCursor(Qt.PointingHandCursor)
        button.setProperty("modeValue", mode_value)
        button.setProperty("keywordValue", keyword_value)
        if mode_value == "search-keyword":
            button.setProperty("smallText", True)
        else:
            button.setProperty("smallText", False)
        button.clicked.connect(self.on_filter_mode_button_clicked)
        button.setStyleSheet(self._filter_button_style(False))
        return button

    def _filter_button_style(self, active):
        small_text = False
        sender_button = self.sender()
        if sender_button is not None:
            small_text = bool(sender_button.property("smallText"))
        font_size = "11px" if small_text else "12px"

        if active:
            return (
                "QPushButton {"
                "background-color: rgba(255, 255, 255, 235);"
                "border: 2px solid rgba(140, 30, 30, 210);"
                "border-radius: 10px;"
                "color: #4a0000;"
                f"font-size: {font_size};"
                "font-weight: bold;"
                "padding: 6px 10px;"
                "min-height: 28px;"
                "}"
            )
        return (
            "QPushButton {"
            "background-color: rgba(255, 255, 255, 195);"
            "border: 1px solid rgba(255, 255, 255, 120);"
            "border-radius: 10px;"
            "color: #3a0000;"
            f"font-size: {font_size};"
            "font-weight: bold;"
            "padding: 6px 10px;"
            "min-height: 28px;"
            "}"
            "QPushButton:hover {"
            "background-color: rgba(255, 240, 240, 225);"
            "}"
        )

    def _refresh_filter_button_styles(self):
        for button in self.filter_buttons_row:
            mode_value = button.property("modeValue")
            keyword_value = button.property("keywordValue")
            small_text = bool(button.property("smallText"))
            is_active = False
            if self.active_filter_mode == "all" and mode_value == "all":
                is_active = True
            elif self.active_filter_mode == "by-key" and mode_value == "by-key":
                is_active = True
            elif self.active_filter_mode == "keyword" and mode_value == "keyword" and self.normalize_text(str(keyword_value)) == self.normalize_text(self.active_keyword):
                is_active = True
            button.setStyleSheet(self._filter_button_style_with_size(is_active, small_text))

    def _filter_button_style_with_size(self, active, small_text):
        font_size = "11px" if small_text else "12px"
        if active:
            return (
                "QPushButton {"
                "background-color: rgba(255, 255, 255, 235);"
                "border: 2px solid rgba(140, 30, 30, 210);"
                "border-radius: 10px;"
                "color: #4a0000;"
                f"font-size: {font_size};"
                "font-weight: bold;"
                "padding: 6px 10px;"
                "min-height: 28px;"
                "}"
            )
        return (
            "QPushButton {"
            "background-color: rgba(255, 255, 255, 195);"
            "border: 1px solid rgba(255, 255, 255, 120);"
            "border-radius: 10px;"
            "color: #3a0000;"
            f"font-size: {font_size};"
            "font-weight: bold;"
            "padding: 6px 10px;"
            "min-height: 28px;"
            "}"
            "QPushButton:hover {"
            "background-color: rgba(255, 240, 240, 225);"
            "}"
        )

    def on_filter_mode_button_clicked(self):
        button = self.sender()
        if not button:
            return
        mode_value = button.property("modeValue")
        keyword_value = button.property("keywordValue")

        if mode_value == "search-keyword":
            keyword_text = str(keyword_value or "").strip()
            if keyword_text:
                self.search_bar.setText(keyword_text)
                self.search_bar.setFocus()
                self.search_bar.setCursorPosition(len(self.search_bar.text()))
            return

        if mode_value == "clear-search":
            self.search_bar.clear()
            self.search_bar.setFocus()
            return

        self.active_filter_mode = str(mode_value or "all")
        self.active_keyword = str(keyword_value or "")
        self._refresh_filter_button_styles()
        self.extract_and_render_sections()

    def show_loading_message(self, text):
        while self.buttons_layout.count():
            item = self.buttons_layout.takeAt(0)
            widget = item.widget()
            if widget:
                widget.deleteLater()

        self.buttons_layout.addStretch(1)
        loading = QLabel(text)
        loading.setAlignment(Qt.AlignCenter)
        loading.setStyleSheet("color: white; font-size: 16px; font-weight: bold; background: transparent;")
        self.buttons_layout.addWidget(loading)
        self.buttons_layout.addStretch(1)

    def determine_file_path(self):
        if self.presentation_path:
            self.file_path = self.presentation_path
            return
        super().determine_file_path()

    def set_button_style(self, button):
        button.setStyleSheet(
            """
            QPushButton {
                background-color: rgba(255, 255, 255, 200);
                border: none;
                border-radius: 12px;
                color: #3a0000;
                padding: 10px;
                font-size: 14px;
                font-weight: bold;
                text-align: center;
                min-height: 30px;
            }
            QPushButton:hover {
                background-color: rgba(255, 240, 240, 230);
                color: #690000;
                border: 1px solid rgba(255, 255, 255, 50);
            }
            """
        )
        button.setLayoutDirection(Qt.RightToLeft)

    def extract_and_render_sections(self):
        while self.buttons_layout.count():
            item = self.buttons_layout.takeAt(0)
            widget = item.widget()
            if widget:
                widget.deleteLater()

        self.section_buttons = []
        records = self._extract_section_records(self.file_path)
        self.section_records = records
        filtered = self._apply_context_filter(records)

        if records and not filtered:
            # If filter returns no sections, fallback to full section list (Option A).
            self.status_label.setText("لا توجد نتائج مطابقة للفلتر، يتم عرض كل الأقسام")
            self.status_label.show()
            filtered = records
        else:
            self.status_label.hide()

        self.filtered_records = filtered
        if not filtered:
            label = QLabel("لم يتم العثور على أقسام")
            label.setAlignment(Qt.AlignCenter)
            label.setStyleSheet("color: white; font-size: 16px; font-weight: bold; background: transparent;")
            self.buttons_layout.addWidget(label)
            return

        for rec in filtered:
            button = QPushButton(rec["name"])
            self.set_button_style(button)
            button.clicked.connect(lambda _, idx=rec["slide_index"]: self.go_to_slide(self.file_path, idx))
            self.buttons_layout.addWidget(button)
            self.section_buttons.append(button)

    def _extract_section_records(self, file_path):
        records = []
        try:
            pythoncom.CoInitialize()
            try:
                powerpoint = win32com.client.GetActiveObject("PowerPoint.Application")
            except:
                powerpoint = win32com.client.Dispatch("PowerPoint.Application")

            presentation = None
            just_opened = False
            for pres in powerpoint.Presentations:
                if os.path.abspath(pres.FullName.lower()) == os.path.abspath(file_path.lower()):
                    presentation = pres
                    break

            if presentation is None:
                presentation = powerpoint.Presentations.Open(file_path, WithWindow=False)
                just_opened = True

            if presentation.SectionProperties.Count > 0:
                for i in range(1, presentation.SectionProperties.Count + 1):
                    name = presentation.SectionProperties.Name(i)
                    first_slide_index = presentation.SectionProperties.FirstSlide(i)
                    slide_id = ""
                    try:
                        slide_id = str(presentation.Slides(first_slide_index).SlideID)
                    except:
                        slide_id = ""
                    section_guid = self._extract_guid_from_text(name)
                    records.append(
                        {
                            "name": name,
                            "normalized": self.normalize_text(name),
                            "slide_index": first_slide_index,
                            "section_id": section_guid,
                            "slide_id": slide_id,
                        }
                    )
            else:
                for i in range(1, presentation.Slides.Count + 1):
                    slide = presentation.Slides.Item(i)
                    text = ""
                    for shape in slide.Shapes:
                        if shape.HasTextFrame and shape.TextFrame.HasText:
                            raw = shape.TextFrame.TextRange.Text
                            if raw and raw.strip():
                                text = raw.strip()
                                break
                    if text:
                        records.append(
                            {
                                "name": text,
                                "normalized": self.normalize_text(text),
                                "slide_index": i,
                                "section_id": self._extract_guid_from_text(text),
                                "slide_id": str(slide.SlideID),
                            }
                        )

            if just_opened:
                presentation.Close()
        except Exception as e:
            self.status_label.setText(f"خطأ في قراءة الأقسام: {str(e)}")
            self.status_label.show()
        finally:
            pythoncom.CoUninitialize()

        return records

    def _apply_context_filter(self, records):
        if not records:
            return []

        filter_mode = (self.filter_mode or "").strip().lower()
        filtered = []

        if self.active_filter_mode == "all":
            filtered = list(records)
        elif self.active_filter_mode == "by-key" and filter_mode == "by-key" and self.filter_key:
            target = self.normalize_text(self.filter_key)
            filtered = [rec for rec in records if target in rec["normalized"]]
        elif filter_mode == "by-keywords":
            # For by-keywords mode we show all sections, and keyword chips just fill search bar text.
            filtered = list(records)
        else:
            filtered = list(records)

        if not self.always_include_section_ids:
            return filtered

        # Always include hardcoded constant sections even when keyword/key filtering is active.
        keep_set = set(id(rec) for rec in filtered)
        for rec in records:
            if self._matches_constant_section(rec):
                keep_set.add(id(rec))

        return [rec for rec in records if id(rec) in keep_set]

    def _matches_constant_section(self, record):
        constants = [str(x).strip() for x in self.always_include_section_ids if str(x).strip()]
        if not constants:
            return False

        name_norm = record.get("normalized", "")
        section_id = str(record.get("section_id", "") or "")
        slide_id = str(record.get("slide_id", "") or "")
        for item in constants:
            item_norm = self.normalize_text(item)
            if item == section_id or item == slide_id:
                return True
            if item_norm and item_norm in name_norm:
                return True
        return False

    def _extract_guid_from_text(self, text):
        match = re.search(r"\{[0-9A-Fa-f-]{36}\}", text or "")
        return match.group(0) if match else ""

    def filter_buttons(self):
        search_text = self.normalize_text(self.search_bar.text().strip())
        visible = 0
        for button in self.section_buttons:
            is_visible = (search_text in self.normalize_text(button.text()))
            button.setVisible(is_visible)
            if is_visible:
                visible += 1

        if visible == 0 and self.section_buttons:
            self.status_label.setText("لا توجد نتائج للبحث")
            self.status_label.show()
        elif self.status_label.text() == "لا توجد نتائج للبحث":
            self.status_label.hide()