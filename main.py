from PyQt5.QtWidgets import QApplication
from dialog_fill_xlsx import dialog_fill_xlsx
import sys

if __name__ == "__main__":

    app = QApplication([])

    d = dialog_fill_xlsx()
    d.exec()

    sys.exit(app.exec())
