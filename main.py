"""EvoSend — ponto de entrada PySide6."""
import sys, os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from PySide6.QtWidgets import QApplication
from PySide6.QtCore import Qt
from ui.theme import apply_theme
from ui.app import App

def main():
    app = QApplication(sys.argv)
    app.setApplicationName("EvoSend")
    apply_theme(app)
    window = App()
    window.show()
    code = app.exec()
    window.cleanup()
    sys.exit(code)

if __name__ == "__main__":
    main()
