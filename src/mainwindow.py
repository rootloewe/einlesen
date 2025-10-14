import getpass
import os
import subprocess
import platform
from PyQt6.QtWidgets import QApplication, QWidget, QMainWindow, QVBoxLayout, QLabel, QPushButton, QFileDialog, QMessageBox, QDialog
from PyQt6.QtGui import QAction
from PyQt6.QtCore import Qt
from einlesen_lib import pfad
from workersthread import WorkerThread
from pathlib import Path



class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Einlesen von Zugferd zu Excel")
        self.setFixedSize(500, 300)

        central_widget = QWidget()
        self.layout = QVBoxLayout()
        self.layout.setContentsMargins(20, 20, 20, 20)
        self.layout.setSpacing(5)
        self.layout.setAlignment(Qt.AlignmentFlag.AlignTop)
        central_widget.setLayout(self.layout)
        self.setCentralWidget(central_widget)

        self.status = self.statusBar()
        self.status.showMessage("Version: 2.6.1")

        self.__createMenuBar()

        self.label = QLabel("Hallo, ich importiere deine Rechnungen in Excel.", self)
        self.label.setStyleSheet("font-size: 11pt;")
        self.layout.addWidget(self.label)
        self.layout.addSpacing(50)
        self.label2 = QLabel("Bitte wähle deine PDF-Rechnungen aus.", self)
        self.label2.setStyleSheet("font-size: 10pt;")
        self.layout.addWidget(self.label2)

        # Button zum Öffnen des Datei-Dialogs
        self.layout.addSpacing(50)
        self.open_pdf_button = QPushButton("PDF auswählen")
        self.open_pdf_button.clicked.connect(self.open_pdf_dialog)
        self.layout.addWidget(self.open_pdf_button)

        # Buttons für "Öffnen" und "Abbrechen" (zunächst verstecken)
        self.open_button = QPushButton("Öffnen", self)
        self.open_button.clicked.connect(self.open_file)
        self.open_button.hide()
        self.layout.addWidget(self.open_button)

        self.cancel_button = QPushButton("Abbrechen", self)
        self.cancel_button.clicked.connect(self.close)
        self.cancel_button.hide()
        self.layout.addWidget(self.cancel_button)

        self.pdf_files = []


    def __createMenuBar(self):
        menubar = self.menuBar()

        # Datei-Menü
        self.file_menu = menubar.addMenu("&Datei")
        self.open_action = QAction("&Öffnen", self)
        self.open_action.setShortcut("Ctrl+O")
        self.open_action.triggered.connect(self.open_pdf_dialog)
        self.file_menu.addAction(self.open_action)
        self.file_menu.addSeparator()
        self.exit_action = QAction("&Beenden", self)
        self.exit_action.setShortcut("Ctrl+Q")
        self.exit_action.setStatusTip("Exit application")
        self.exit_action.triggered.connect(QApplication.instance().quit)
        self.file_menu.addAction(self.exit_action)

        # Hilfe-Menü
        help_menu = menubar.addMenu("&Hilfe")
        about_action = QAction("&Über", self)
        about_action.triggered.connect(self.__ueber)
        help_menu.addAction(about_action)


    def start_worker(self):
        self.worker = WorkerThread(self.pdf_files)
        self.worker.noXmlFound.connect(self.handle_no_xml)
        self.worker.finished.connect(self.on_finished)
        self.worker.start()


    def on_finished(self):
        # Nach Abschluss des Hintergrundprozesses Buttons anzeigen und Text anpassen
        self.open_button.show()
        self.cancel_button.show()
        self.open_pdf_button.hide()


    def open_pdf_dialog(self):
        # pfad = Path(r"C:\Users\goli1\OneDrive\Desktop\Datanorm Sonepar")
        pfad = r"V:\Zugferd Sonepar"

        if not Path(pfad).exists():
            pfad = ""

        # wegen windows, weil sonst englisch
        options = QFileDialog.Option.DontUseNativeDialog

        filenames, _ = QFileDialog.getOpenFileNames(self, "PDF auswählen", pfad, "PDF Dateien (*.pdf)", options=options)

        if filenames:
            self.pdf_files = filenames
            self.open_action.setEnabled(False)

            self.worker = WorkerThread(self.pdf_files)
            self.worker.finished.connect(self.on_finished)
            self.start_worker()
            self.label2.setStyleSheet("font-size: 10pt;")
            self.label2.setText("PDFs werden verarbeitet ... fertig.")


    def open_file(self):
        dateipfad = pfad()

        dateiname = str(dateipfad / "Überweisung Großhandel.xltx")

        if not os.path.exists(dateiname):
            self.label.setText(f"Datei nicht gefunden:\n{dateiname}")
            return

        try:
            if platform.system() == "Linux":
                subprocess.run(["libreoffice", dateiname])

            elif platform.system() == "Windows":
                os.startfile(dateiname)
            elif platform.system() == "Darwin":
                subprocess.run(["open", dateiname])
            else:
                print("Unsupported OS")

        except Exception as e:
            self.label.setText(f"Fehler beim Öffnen:\n{e}")

        finally:
            self.close()
            QApplication.instance().quit()


    def handle_no_xml(self):
        ret = QMessageBox.question(
            self,
            "Keine XML-Dateien gefunden!",
            "Während der Verarbeitung wurden keine zugehörigen XML-Dateien erzeugt.\n"
            "Möchtest Du eine neue PDF-Auswahl treffen (Neustart) oder abbrechen?",
            QMessageBox.StandardButton.Retry | QMessageBox.StandardButton.Cancel,
            QMessageBox.StandardButton.Retry,
        )

        if ret == QMessageBox.StandardButton.Retry:
            self.open_pdf_dialog()  # oder deine Methode, die neue Auswahl ermöglicht
        else:
            self.label2.setText("Verarbeitung abgebrochen.")
            QApplication.quit()


    def __ueber(self):
        # Dialogfenster erzeugen, selbstständig aber modal (blockiert das Hauptfenster)
        about_dialog = QDialog(self)
        about_dialog.setWindowTitle("Über dieses Programm:")
        about_dialog.setFixedSize(300, 150)
        about_dialog.setModal(True)  # Modal: blockiert Hauptfenster, bis geschlossen

        layout = QVBoxLayout()

        label1 = QLabel("Zugferd einlesen in Excel v2.5.4")
        label1.setStyleSheet("font-weight: bold; font-size: 12pt;")
        layout.addWidget(label1, alignment=Qt.AlignmentFlag.AlignCenter)

        label2 = QLabel("Erstellt von: Ingenieurüro Inh. Armin Herzner")
        layout.addWidget(label2, alignment=Qt.AlignmentFlag.AlignCenter)


        label3 = QLabel("© 2025, Alle Rechte vorbehalten")
        layout.addWidget(label3, alignment=Qt.AlignmentFlag.AlignCenter)


        close_button = QPushButton("Schließen")
        close_button.clicked.connect(about_dialog.accept)  # Dialog schließt sich
        layout.addWidget(close_button, alignment=Qt.AlignmentFlag.AlignCenter)

        about_dialog.setLayout(layout)
        about_dialog.exec()  # Zeigt den Dialog modal an und wartet bis er geschlossen wirdpass

