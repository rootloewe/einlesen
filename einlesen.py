import sys
import time
from PyQt6.QtWidgets import QApplication
from mainwindow import MainWindow
from PyQt6.QtCore import QTranslator, QLocale, QLibraryInfo


def main():
    app = QApplication(sys.argv)

    translator = QTranslator()
    locale = QLocale.system().name()
    translator.load("qt_" + locale, QLibraryInfo.path(QLibraryInfo.LibraryPath.TranslationsPath))
    app.installTranslator(translator)

    window = MainWindow()
    window.show()
    time.sleep(0.1)
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
   

