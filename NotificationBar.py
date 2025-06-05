from PyQt5.QtWidgets import QLabel, QGraphicsDropShadowEffect, QPushButton, QHBoxLayout, QFrame
from PyQt5.QtCore import Qt, QTimer, QPropertyAnimation, QEasingCurve
from PyQt5.QtGui import QFont, QColor, QCursor


class NotificationBar(QFrame):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.timer = QTimer(self)
        self.timer.timeout.connect(self.hide)
        
        # Set default values
        self.horizontal_padding = 20
        self.max_width = 600  # Default max width
        
        # Apply modern styling directly to the main frame
        self.setStyleSheet("""
            QFrame {
                background-color: rgba(52, 152, 219, 180);
                border-radius: 10px;
                border: 1px solid rgba(255, 255, 255, 0.2);
                color: white;
                font-weight: bold;
                font-size: 12pt;
            }
        """)
        
        # Add shadow for depth
        shadow = QGraphicsDropShadowEffect()
        shadow.setBlurRadius(15)
        shadow.setOffset(0, 2)
        shadow.setColor(QColor(0, 0, 0, 100))
        self.setGraphicsEffect(shadow)
        
        # Create label directly on the frame
        self.label = QLabel(self)
        self.label.setAlignment(Qt.AlignCenter)
        self.label.setStyleSheet("background: transparent; border: none;")
        
        # Add close button directly on the frame
        self.close_btn = QPushButton("×", self)
        self.close_btn.setFixedSize(24, 24)
        self.close_btn.setStyleSheet("""
            QPushButton {
                background-color: transparent;
                color: white;
                font-size: 16px;
                font-weight: bold;
                border: none;
            }
            QPushButton:hover {
                background-color: rgba(255, 255, 255, 0.2);
                border-radius: 12px;
            }
        """)
        self.close_btn.clicked.connect(self.hide)
        self.hide()

    def show_message(self, message, duration=3000):
        self.label.setText(message)
        self.max_width = self.parent().width() - 40 if self.parent() else 600
        self.horizontal_padding = 20
        self.adjust_size()
        self.show()
        self.timer.start(duration)

    def adjust_size(self):
        # Get the width of the parent (main window)
        if self.parent():
            self.max_width = self.parent().width() - 40  # 20px margin on each side
        else:
            self.max_width = 600  # Default max width
            
        # Calculate the optimal width for the text
        text_width = self.label.fontMetrics().width(self.label.text())
        
        # Determine the new width (constrained by max_width)
        new_width = min(text_width + 2 * self.horizontal_padding + 30, self.max_width)
        
        # Get the parent width to center the notification
        parent_width = self.parent().width() if self.parent() else self.max_width + 40
        
        # Center the notification bar horizontally
        new_x = (parent_width - new_width) // 2
        
        # Set the new geometry
        self.setGeometry(new_x, self.y(), new_width, 50)
        
        # Position the label and close button
        self.label.setGeometry(self.horizontal_padding, 0, new_width - 2*self.horizontal_padding - 24, 50)
        self.close_btn.move(new_width - 30, 13)  # Position close button

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
        self.close_btn.move(self.width() - 30, 5)
        super().resizeEvent(event)