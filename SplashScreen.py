from PyQt5.QtWidgets import QWidget, QVBoxLayout, QLabel, QApplication
from PyQt5.QtGui import QPixmap, QFont, QColor
from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import QGraphicsDropShadowEffect
from commonFunctions import relative_path

class ModernSplashScreen(QWidget):
    def __init__(self):
        super().__init__()

        self.setWindowFlags(Qt.FramelessWindowHint | Qt.WindowStaysOnTopHint)
        self.setAttribute(Qt.WA_TranslucentBackground)

        # Background container
        self.container = QWidget(self)
        self.container.setStyleSheet("""
            background: qlineargradient(
                x1: 0, y1: 0, x2: 1, y2: 1,
                stop: 0 #0f2e47,
                stop: 0.6 #1e5b8a,
                stop: 1 #8cd9ff
            );
            border-radius: 20px;
        """)

        # Drop shadow
        shadow = QGraphicsDropShadowEffect()
        shadow.setBlurRadius(30)
        shadow.setOffset(0)
        shadow.setColor(QColor(0, 0, 0, 150))
        self.container.setGraphicsEffect(shadow)

        # Layout
        layout = QVBoxLayout(self.container)
        layout.setContentsMargins(30, 30, 30, 30)
        layout.setSpacing(15)

        # Image (logo)
        pix = QPixmap(relative_path(r"Data/الصور/St Mary's Liturgies Logo.png"))
        pix = pix.scaled(180, 180, Qt.KeepAspectRatio, Qt.SmoothTransformation)
        image_label = QLabel()
        image_label.setPixmap(pix)
        image_label.setAlignment(Qt.AlignCenter)

        # Text
        self.text_label = QLabel("St Mary Maadi Liturgies\nLoading...")
        self.text_label.setAlignment(Qt.AlignCenter)
        self.text_label.setStyleSheet("color: white; font-weight: bold;")
        self.text_label.setFont(QFont("Segoe UI", 14, QFont.Bold))

        layout.addWidget(image_label)
        layout.addWidget(self.text_label)

        self.resize(340, 340)
        self.container.resize(340, 340)
        self.center_on_screen()

    def center_on_screen(self):
        screen = QApplication.primaryScreen().availableGeometry()
        size = self.geometry()
        self.move(
            (screen.width() - size.width()) // 2,
            (screen.height() - size.height()) // 2
        )

    def update_progress(self):
        self.text_label.setText(f"St Mary Maadi Liturgies\nLoading...")
