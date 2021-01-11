import sys
from PyQt5.QtWidgets import QApplication, QWidget, QPushButton, QVBoxLayout, QPlainTextEdit

import openpyxl


class gui(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        ta = QPlainTextEdit(self)
        ta.setPlaceholderText('내용 붙여넣기')

        btn = QPushButton(self)
        btn.setText('생성')
        btn.clicked.connect(self.makeExcel)

        vbox = QVBoxLayout()
        vbox.addWidget(ta)
        vbox.addWidget(btn)

        self.setLayout(vbox)
        self.setWindowTitle('QPushButton')
        self.setGeometry(300, 300, 300, 200)
        self.show()

    def makeExcel(self):
        # A 65 / a 97
        startCol = 65
        col = 3
        # endcol = startCol + col - 1
        print('makeExcel start..')

        wb = openpyxl.Workbook()
        sheet = wb.active

        sheet.merge_cells('A1:E4')
        sheet['A1'] = 'testtest'

        wb.save("changed_row.xlsx")
        print('makeExcel end..')

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = gui()
    sys.exit(app.exec_())

# pip install openpyxl
# pip install PyQt5