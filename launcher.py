import ctypes
ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("Ocean.StockCount")

import sys, os
from PySide6.QtCore import Qt, QTimer
from PySide6.QtGui import QPixmap, QIcon
from PySide6.QtWidgets import QApplication, QSplashScreen

def resource_path(relative_path: str) -> str:
    if hasattr(sys, "_MEIPASS"):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), relative_path)

def main():
    app = QApplication(sys.argv)

    app.setWindowIcon(QIcon(resource_path("icon_stock_count.ico")))

    pix = QPixmap(resource_path("icon_stock_count.ico"))
    splash = QSplashScreen(pix)
    splash.showMessage("Loading...", Qt.AlignHCenter | Qt.AlignBottom)
    splash.show()
    app.processEvents()

    def start_real_app():
        import app as real_app
        w = real_app.Main()
        w.show()
        splash.finish(w)

    QTimer.singleShot(0, start_real_app)
    sys.exit(app.exec())

if __name__ == "__main__":
    main()
