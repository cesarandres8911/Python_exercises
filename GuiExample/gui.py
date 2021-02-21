from PyQt5.QtWidgets import QApplication,  QMainWindow
import sys
from PyQt5 import QtGui

class Window(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setGeometry(600, 300, 500, 400)
        self.setWindowTitle("PyQt5 Window")

        self.show()

App = QApplication(sys.argv)
window = Window()
sys.exit(App.exec())