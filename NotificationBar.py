import logging

from PyQt5.QtWidgets import QLabel, QGraphicsDropShadowEffect, QPushButton, QFrame
from PyQt5.QtCore import Qt, QTimer, QPropertyAnimation, QEasingCurve, QAbstractAnimation
from PyQt5.QtGui import QColor


class NotificationBar(QFrame):
    def __init__(self, parent=None):
        super().__init__(parent)
        # Logger for defensive error-reporting
        self.logger = logging.getLogger(__name__)

        # Single-shot timer used to auto-hide the notification
        self.timer = QTimer(self)
        self.timer.setSingleShot(True)
        self.timer.timeout.connect(self._on_timeout)

        # Set default values
        self.horizontal_padding = 20
        self.max_width = 600  # Default max width
        self.margin_top = 10
        self.fade_duration = 250

        # Apply modern styling directly to the main frame
        self.setStyleSheet("""
            QFrame {
                background-color: rgba(15, 76, 117, 230);
                border-radius: 10px;
                border: 1px solid rgba(255, 255, 255, 0.08);
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
        self.label.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        self.label.setStyleSheet("background: transparent; border: none;")
        self.label.setWordWrap(True)
        
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
        self.close_btn.clicked.connect(self._on_close_clicked)

        # Create fade animation for nicer dismissal
        try:
            self.fade_out_animation = QPropertyAnimation(self, b"windowOpacity", self)
            self.fade_out_animation.setDuration(self.fade_duration)
            self.fade_out_animation.setEasingCurve(QEasingCurve.InOutQuad)
            self.fade_out_animation.finished.connect(self._on_fade_finished)
        except Exception:
            # If animation creation fails, log and continue without animation
            self.logger.exception("Failed to create fade animation")
            self.fade_out_animation = None

        # Start hidden and fully opaque
        self.setWindowOpacity(1.0)
        super().hide()

    def show_message(self, message, duration=3000):
        try:
            if message is None:
                message = ""
            self.label.setText(str(message))

            parent = self.parent()
            if parent and hasattr(parent, "width"):
                try:
                    self.max_width = max(200, parent.width() - 40)
                except Exception:
                    self.max_width = 600
            else:
                self.max_width = 600

            self.horizontal_padding = 20
            self.adjust_size()
            self.setWindowOpacity(1.0)
            super().show()
            # Ensure duration is a positive integer
            self.timer.start(max(100, int(duration)))
        except Exception:
            self.logger.exception("show_message failed")
            try:
                print("Notification:", message)
            except Exception:
                pass

    def adjust_size(self):
        try:
            # Get the width of the parent (main window)
            if self.parent():
                try:
                    parent_width = self.parent().width()
                    self.max_width = max(200, parent_width - 40)
                except Exception:
                    parent_width = self.max_width + 40
            else:
                parent_width = self.max_width + 40

            # Calculate the optimal width for the text
            fm = self.label.fontMetrics()
            text_width = fm.horizontalAdvance(self.label.text())

            # Determine the new width (constrained by max_width)
            new_width = min(text_width + 2 * self.horizontal_padding + 30, self.max_width)

            # Center the notification bar horizontally
            new_x = (parent_width - new_width) // 2

            # Prepare label width to compute wrapped height
            label_width = max(80, new_width - 2*self.horizontal_padding - 24)
            self.label.setFixedWidth(label_width)
            # Let the label compute its required height when wrapped
            self.label.adjustSize()
            text_height = self.label.sizeHint().height()

            # Determine the new height based on text height
            new_height = max(50, text_height + 12)

            # Set the new geometry (use a fixed top margin)
            y = getattr(self, "margin_top", 10)
            self.setGeometry(new_x, y, new_width, new_height)

            # Position the label and close button (vertically center label)
            label_y = (new_height - text_height) // 2
            self.label.setGeometry(self.horizontal_padding, label_y, label_width, text_height)
            if hasattr(self, "close_btn"):
                self.close_btn.move(new_width - 30, (new_height - self.close_btn.height()) // 2)
        except Exception:
            self.logger.exception("adjust_size failed")

    def center_in_parent(self):
        """Center the notification in its parent"""
        try:
            parent_width = self.parent().width() if self.parent() and hasattr(self.parent(), "width") else getattr(self, "parent_width", self.max_width + 40)
            x_pos = (parent_width - self.width()) // 2
            self.move(x_pos, getattr(self, "margin_top", 10))
        except Exception:
            self.logger.exception("center_in_parent failed")

    def start_fade_out(self):
        """Start the fade out animation"""
        try:
            if getattr(self, "fade_out_animation", None):
                # avoid restarting while running
                if self.fade_out_animation.state() == QAbstractAnimation.Running:
                    return
                self.fade_out_animation.setStartValue(self.windowOpacity())
                self.fade_out_animation.setEndValue(0.0)
                self.fade_out_animation.start()
            else:
                # no animation available; hide immediately
                super().hide()
        except Exception:
            self.logger.exception("start_fade_out failed")
            try:
                super().hide()
            except Exception:
                pass

    def hide_notification(self):
        """Hide the notification immediately when close button is clicked"""
        try:
            if hasattr(self, 'timer') and self.timer.isActive():
                try:
                    self.timer.stop()
                except Exception:
                    pass
            # Fade out when auto-hiding is requested
            self.start_fade_out()
        except Exception:
            self.logger.exception("hide_notification failed")
            try:
                super().hide()
            except Exception:
                pass

    def _on_close_clicked(self):
        """Immediate close when user clicks the close button."""
        try:
            # Stop timers and animations, then hide immediately
            try:
                if hasattr(self, 'timer') and self.timer.isActive():
                    self.timer.stop()
            except Exception:
                pass

            try:
                if getattr(self, 'fade_out_animation', None):
                    if self.fade_out_animation.state() == QAbstractAnimation.Running:
                        self.fade_out_animation.stop()
            except Exception:
                pass

            # Hide without animation for immediate response
            super().hide()
        except Exception:
            self.logger.exception("_on_close_clicked failed")

    def set_parent_width(self, width):
        """Update the parent width reference when parent is resized"""
        try:
            self.parent_width = width
            if self.isVisible():
                self.center_in_parent()
        except Exception:
            self.logger.exception("set_parent_width failed")

    def resizeEvent(self, event):
        """Reposition close button when resized"""
        try:
            if hasattr(self, "close_btn"):
                self.close_btn.move(self.width() - 30, 5)
        except Exception:
            self.logger.exception("resizeEvent failed")
        super().resizeEvent(event)

    def _on_timeout(self):
        try:
            self.start_fade_out()
        except Exception:
            self.logger.exception("_on_timeout failed")

    def _on_fade_finished(self):
        try:
            # If we've faded out, fully hide and reset opacity
            if getattr(self, "fade_out_animation", None) and self.windowOpacity() == 0.0:
                try:
                    self.timer.stop()
                except Exception:
                    pass
                self.setWindowOpacity(1.0)
                super().hide()
        except Exception:
            self.logger.exception("_on_fade_finished failed")

    def hide(self):
        # Ensure timer is stopped before hiding
        try:
            if hasattr(self, 'timer') and self.timer.isActive():
                try:
                    self.timer.stop()
                except Exception:
                    pass
        except Exception:
            pass
        super().hide()