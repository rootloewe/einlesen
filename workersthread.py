from PyQt6.QtCore import QThread, pyqtSignal
from xltx_bin import binärstr, xltx_base64
from einlesen_lib import pfad, xltx_bauen, xml_erstellen, csv_erstellen, xltx_erstellen


class WorkerThread(QThread):
    noXmlFound = pyqtSignal()
    finished = pyqtSignal()
    error_signal = pyqtSignal(str)


    def __init__(self, pdf_files, parent=None):
        super().__init__(parent)
        self.pdf_files = pdf_files


    def run(self):
        dateipfad = pfad()

        b64str = binärstr(xltx_base64)
        xltx_bauen(dateipfad, b64str)

        pdf_files = self.pdf_files
        xml_erstellen(dateipfad, pdf_files)

        erg = csv_erstellen(dateipfad)

        if not erg:
            self.noXmlFound.emit()

        else:
            xltx_erstellen(dateipfad)
            self.finished.emit()
