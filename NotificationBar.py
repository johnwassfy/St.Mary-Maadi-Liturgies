from PyQt5.QtWidgets import QLabel, QGraphicsDropShadowEffect, QSizePolicy, QPushButton
from PyQt5.QtCore import Qt, QTimer, QPropertyAnimation, QEasingCurve
from PyQt5.QtGui import QFont, QColor, QCursor


class NotificationBar(QLabel):
    def __init__(self, parent=None):
        super().__init__(parent)
        
        # Initialize attributes first
        self.parent_width = 800
        self.min_width = 200
        self.max_width = 600
        self.min_height = 50
        self.max_height = 150
        self.duration = 3000
        self.margin_top = 20
        self.horizontal_padding = 40
        self.vertical_padding = 15
        
        self.setup_ui()
        self.setup_animations()
        
        self.setVisible(False)
        self.setWordWrap(True)

    def setup_ui(self):
        """Set up the visual appearance of the notification bar"""
        # Font settings
        try:
            font = QFont()
            available_fonts = QFont().families()
            preferred_fonts = ["Segoe UI", "Arial", "Helvetica", "Verdana"]
            
            for font_name in preferred_fonts:
                if font_name in available_fonts:
                    font.setFamily(font_name)
                    break
            
            # Increased font size from 14 to 16
            font.setPointSize(16)
            font.setWeight(QFont.Medium)
            self.setFont(font)
        except:
            pass
        
        # Create close button
        self.close_button = QPushButton("×", self)  # × symbol
        self.close_button.setStyleSheet("""
            QPushButton {
                color: white;
                font-size: 18px;
                font-weight: bold;
                border: none;
                background: transparent;
                padding: 0px 4px;
                margin-right: 5px;
            }
            QPushButton:hover {
                color: #ff6666;
            }
        """)
        self.close_button.setCursor(QCursor(Qt.PointingHandCursor))
        self.close_button.setFixedSize(24, 24)
        self.close_button.clicked.connect(self.hide_notification)
        
        # Style sheet for the label
        self.setStyleSheet(f"""
            QLabel {{
                background-color: rgba(40, 40, 40, 220);
                color: white;
                padding: {self.vertical_padding}px {self.horizontal_padding}px;
                border-radius: 8px;
                border: 1px solid rgba(255, 255, 255, 30);
            }}
        """)
        
        # Add shadow effect
        shadow = QGraphicsDropShadowEffect()
        shadow.setBlurRadius(15)
        shadow.setColor(QColor(0, 0, 0, 160))
        shadow.setOffset(0, 4)
        self.setGraphicsEffect(shadow)
        
        self.setAlignment(Qt.AlignCenter)
        self.setMinimumHeight(self.min_height)
        self.setMaximumHeight(self.max_height)
        self.setSizePolicy(QSizePolicy.MinimumExpanding, QSizePolicy.Fixed)

    def setup_animations(self):
        """Set up fade in/out animations"""
        self.fade_in_animation = QPropertyAnimation(self, b"windowOpacity")
        self.fade_in_animation.setDuration(300)
        self.fade_in_animation.setStartValue(0)
        self.fade_in_animation.setEndValue(1)
        self.fade_in_animation.setEasingCurve(QEasingCurve.OutCubic)

        self.fade_out_animation = QPropertyAnimation(self, b"windowOpacity")
        self.fade_out_animation.setDuration(300)
        self.fade_out_animation.setStartValue(1)
        self.fade_out_animation.setEndValue(0)
        self.fade_out_animation.setEasingCurve(QEasingCurve.InCubic)
        self.fade_out_animation.finished.connect(self.hide)

    def show_message(self, message, duration=3000):
        """Show a notification message with optional duration"""
        self.setText(message)
        self.adjust_size()
        
        # Stop any ongoing animations
        self.fade_in_animation.stop()
        self.fade_out_animation.stop()
        
        # Reset opacity and show
        self.setWindowOpacity(1)
        self.setVisible(True)
        self.raise_()  # Bring to front
        
        # Position close button
        self.close_button.move(self.width() - 30, 5)
        
        # Start animations
        self.fade_in_animation.start()
        if duration > 0:
            self.timer = QTimer.singleShot(duration, self.start_fade_out)

    def adjust_size(self):
        """Adjust the size based on message content"""
        # Calculate required size (accounting for close button space)
        text_rect = self.fontMetrics().boundingRect(
            0, 0, 
            self.max_width - 2 * self.horizontal_padding - 30,  # Space for close button
            self.max_height,
            Qt.TextWordWrap, 
            self.text()
        )
        
        # Calculate width and height with padding
        required_width = min(
            max(text_rect.width() + 2 * self.horizontal_padding + 30, self.min_width),
            self.max_width
        )
        required_height = min(
            max(text_rect.height() + 2 * self.vertical_padding, self.min_height),
            self.max_height
        )
        
        self.setFixedSize(required_width, required_height)
        self.center_in_parent()

    def center_in_parent(self):
        """Center the notification in its parent"""
        parent_width = self.parent().width() if self.parent() else self.parent_width
        x_pos = (parent_width - self.width()) // 2
        self.move(x_pos, self.margin_top)

    def start_fade_out(self):
        """Start the fade out animation"""
        self.fade_out_animation.start()

    def hide_notification(self):
        """Hide the notification immediately when close button is clicked"""
        if hasattr(self, 'timer'):
            self.timer == None  # Stop any existing timer
        self.start_fade_out()

    def set_parent_width(self, width):
        """Update the parent width reference when parent is resized"""
        self.parent_width = width
        if self.isVisible():
            self.center_in_parent()

    def resizeEvent(self, event):
        """Reposition close button when resized"""
        self.close_button.move(self.width() - 30, 5)
        super().resizeEvent(event)