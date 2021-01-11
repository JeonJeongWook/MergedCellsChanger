import sys
from PyQt5.QtWidgets import QApplication, QWidget, QPushButton, QVBoxLayout, QPlainTextEdit, QSpinBox

import openpyxl


class gui(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        colCount = QSpinBox(self)

        btn = QPushButton(self)
        btn.setText('생성')
        btn.clicked.connect(lambda :self.makeExcel(colCount.value()))

        ta = QPlainTextEdit(self)
        ta.setPlaceholderText('내용 붙여넣기')

        vbox = QVBoxLayout()
        vbox.addWidget(colCount)
        vbox.addWidget(ta)
        vbox.addWidget(btn)

        self.setLayout(vbox)
        self.setWindowTitle('QPushButton')
        self.setGeometry(300, 300, 300, 200)
        self.show()

    def makeExcel(self, colCount):
        # A 65 / a 97
        startCol = 65
        colCount = colCount
        print(colCount)



"""
        print('makeExcel start..')

        wb = openpyxl.Workbook()
        sheet = wb.active

        sheet.merge_cells('A1:E4')
        sheet['A1'] = 'testtest'

        # wb.save("changed_row.xlsx")
        print('makeExcel end..')
"""
if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = gui()
    sys.exit(app.exec_())

# pip install openpyxl
# pip install PyQt5