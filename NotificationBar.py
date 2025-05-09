from PyQt5.QtWidgets import QLabel
from PyQt5.QtCore import Qt, QTimer, QPropertyAnimation, QEasingCurve


class NotificationBar(QLabel):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setStyleSheet("""
            QLabel {
                background-color: rgba(0, 0, 0, 180);
                color: white;
                padding: 10px;
                border-radius: 5px;
                font-size: 16px;
                font-weight: bold;
            }
        """)
        self.setAlignment(Qt.AlignCenter)
        self.setFixedHeight(50)
        self.setVisible(False)

        self.fade_in_animation = QPropertyAnimation(self, b"windowOpacity")
        self.fade_in_animation.setDuration(500)
        self.fade_in_animation.setStartValue(0)
        self.fade_in_animation.setEndValue(1)
        self.fade_in_animation.setEasingCurve(QEasingCurve.InOutQuad)

        self.fade_out_animation = QPropertyAnimation(self, b"windowOpacity")
        self.fade_out_animation.setDuration(500)
        self.fade_out_animation.setStartValue(1)
        self.fade_out_animation.setEndValue(0)
        self.fade_out_animation.setEasingCurve(QEasingCurve.InOutQuad)
        self.fade_out_animation.finished.connect(self.hide)

    def show_message(self, message, duration=3000):
        self.setText(message)
        self.adjust_size()
        self.setVisible(True)
        self.fade_in_animation.start()
        QTimer.singleShot(duration, self.start_fade_out)

    def adjust_size(self):
        # Calculate the width based on the message length
        message_width = self.fontMetrics().boundingRect(self.text()).width() + 20  # Add some padding
        self.setFixedWidth(message_width)
        self.move((self.parent().width() - self.width()) // 2, 70)  # Center the notification bar

    def start_fade_out(self):
        self.fade_out_animation.start()

