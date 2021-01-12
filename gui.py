import sys
from PyQt5.QtWidgets import QApplication, QWidget, QPushButton, QVBoxLayout, QPlainTextEdit, QSpinBox, QMessageBox

import openpyxl

class gui(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        vbox = QVBoxLayout()
        self.colCount = QSpinBox(self)
        self.ta = QPlainTextEdit(self)
        btn_make = QPushButton(self)
        # btn_clear
        # 병합할 열 개수 | spinbox | 미리보기(A:A, A:B)

        vbox.addWidget(self.colCount)
        vbox.addWidget(self.ta)
        vbox.addWidget(btn_make)

        self.colCount.setMinimum(1)

        self.ta.setPlaceholderText('내용 붙여넣기')
        self.ta.setPlainText("")

        btn_make.setText('생성하기')
        btn_make.clicked.connect(self.makeExcel)

        self.setLayout(vbox)
        self.setWindowTitle('QPushButton')
        self.setGeometry(300, 300, 300, 200)
        self.show()


    # A 65 / a 97
    def makeExcel(self):
        merged_col = self.colCount.value()
        text = self.ta.toPlainText().strip()
        row_count = self.ta.document().lineCount()

        print('merged_col \t> ', merged_col)
        print('text----------\n', text)
        print("row_count \t> ", row_count)

        if text == "":  # NULL
            QMessageBox.about(self, "오류", '내용을 입력하세요')
        else:  # Not NULL
            print('makeExcel start..')
            start_col   = 65
            end_col     = start_col + merged_col - 1
            wb = openpyxl.Workbook()
            sheet = wb.active

            for i in range(0, row_count):
                result = chr(start_col) + repr(i+1) + ':' + chr(end_col) + repr(i+1)    # A1:b1
                idx = str(chr(start_col)+repr(i+1))     #A1


                    '''
                    # Merging first 3 columns of 1st 3 rows
                        r1 = 1
                        r2 = 3
                        c1 = 1
                        c2 = 3
                        sheet.merge_cells(start_row=r1, start_column=c1, end_row=r2, end_column=c2)
                        https://medium.com/swlh/working-with-spreadsheets-using-python-903202509407
                    '''
                # sheet.merge_cells('a1:b1')
                # sheet['a1'] = "check"
                sheet.merge_cells(self, result)
                sheet['a1'] = "check"

            print('end')
            wb.save("changed_row.xlsx")

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = gui()
    sys.exit(app.exec_())

# pip install openpyxl
# pip install PyQt5
