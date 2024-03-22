
from PyQt5.QtWidgets import QDialog, QWidget , QFileDialog
from ui_dialog_fill_xlsx import Ui_dialog_fill_xlsx


class dialog_fill_xlsx(QDialog ,Ui_dialog_fill_xlsx):
    def __init__(self, parent: QWidget = None) -> None:
        super().__init__(parent)
        self.setupUi(self)
        self.btn_select_fp.clicked.connect(self._select_file)
        self.btn_start.clicked.connect(self._start)


    
    def _select_file(self):
        fp,_ = QFileDialog.getOpenFileName(self,"选择Excel文件", None ,"*.xlsx")

        if len(fp):
            self.edit_fp.setText(fp)
    
    def _start(self):
        from utils import Locationinformation
        Locationinformation(self.edit_fp.text())

