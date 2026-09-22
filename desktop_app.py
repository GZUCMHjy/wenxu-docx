"""Windows GUI entry point; never starts a web server."""
import os
import sys

from PySide6.QtCore import QTimer
from PySide6.QtWidgets import QApplication

from desktop.window import MainWindow


def main():
    os.environ["WENXU_RENDERER"] = "wps"
    app = QApplication(sys.argv)
    app.setApplicationName("文序")
    app.setOrganizationName("Wenxu")
    app.setStyle("Fusion")
    window = MainWindow()
    window.show()
    if len(sys.argv) == 3 and sys.argv[1] == "--self-test":
        from desktop.smoke import start
        start(window, sys.argv[2])
    elif len(sys.argv) == 2:
        QTimer.singleShot(0, lambda: window.open_path(sys.argv[1]))
    return app.exec()


if __name__ == "__main__":
    sys.exit(main())
