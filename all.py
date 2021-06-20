from PyQt5.QtWidgets import QApplication, QMainWindow, QFileDialog, QMessageBox
import sys
from PyQt5.QtGui import QIcon
from PyQt5.QtWidgets import (QTextEdit, QAction)
from main import main_foo
import pandas as pd

def main_foo(x):
    book = pd.read_excel(x, engine='openpyxl')
    column_one = book.columns.tolist()[0]
    column_two = book.columns.tolist()[1]
    column_three = book.columns.tolist()[2]
    column_four = book.columns.tolist()[3]
    num_slice = []
    for i in range(book.shape[0]):
        if str(book[column_one][i]).replace(" ", '').lower() == 'материалы' or str(book[column_two][i]).replace(" ", '').lower() == 'материалы' or str(book[column_three][i]).replace(" ", '').lower() == 'материалы':
            num_slice.append(i)
        elif str(book[column_one][i]).replace(" ", '').lower() == 'итого"материалы"' or str(book[column_two][i]).replace(" ", '').lower() == 'итого"материалы"':
            num_slice.append(i)
            break

    ws = []

    for i in range(num_slice[0] + 1, num_slice[1]):
        ws.append([str(book[column_one][i]), str(book[column_two][i]), str(book[column_three][i]), str(book[column_four][i])])
    df = pd.DataFrame(ws, columns=['code', 'name', 'unit', 'amount'])

    electrodes = 0
    paintwork = 0
    bitumen = 0
    fertilizers = 0
    seeds = 0
    pipes = 0

    pipe = pd.read_excel(r'справочник массы.xlsx', sheet_name='pipe')
    piperpov = []
    for i in range(pipe.shape[0]):
        piperpov.append(pipe['code'][i])

    'расчет электродов'

    for index, row in df.iterrows():
        if str(row['name']).lower().find('электрод') == 0:
            if str(row['unit']).lower() == 'кг':
                electrodes += float(row['amount']) / 1000
            elif str(row['unit']).lower() == 'т':
                electrodes += float(row['amount'])

    electrodes = round(electrodes, 3)
    seeds = round(seeds, 3)
    fertilizers = round(fertilizers, 3)
    pipes = round(pipes, 3)
    message = f"Масса электродов: {electrodes} т\n" \
              f"Масса семян: {seeds} т\n" \
              f"Масса удорбрений: {fertilizers} т\n" \
              f"Масса стальных труб: {pipes} т"

    return message

class Resource_sheet(QMainWindow):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        self.textEdit = QTextEdit()
        self.setCentralWidget(self.textEdit)
        self.textEdit.setStyleSheet('background-color: rgb(16, 24, 32); color: white;')
        self.statusBar()
        self.statusBar().showMessage('version 0.1 Beta')
        self.statusBar().setStyleSheet("font-family: arial; text-align: right;")

        openFile = QAction(QIcon('open.png'), 'Открыть...', self)
        openFile.setShortcut('Ctrl+O')
        openFile.setStatusTip('Открыть файл')
        openFile.triggered.connect(self.showDialog)
        menubar = self.menuBar()
        menubar.setObjectName('menu')
        menubar.setStyleSheet("""
                            #menu {
                                background-color: rgb(16, 24, 32);
                                color: white;
                                padding: 3px;
                                font-size: 9pt;
                                font-family: arial;
                                border: 1px solid black;
                            }
                            #menu:hover {
                                background-color: rgb(16, 24, 32);
                                color: white;
                            }
                            """
            )
        fileMenu = menubar.addMenu('&Файл')
        fileMenu.setStyleSheet('font-family: arial; font-size: 11pt')
        fileMenu.setObjectName('filest')
        fileMenu.setStyleSheet("""
                            #filest {
                                background-color: rgb(16, 24, 32);
                                color: white;
                            }
                            #filest:hover {
                                background-color: rgb(16, 24, 32);
                                color: white;
                            }
        """)
        fileMenu.addAction(openFile)

        self.setGeometry(300, 300, 350, 300)
        self.setWindowTitle('Разбор сводной ресурсной ведомости')
        self.show()

    def showDialog(self):
        fname = QFileDialog.getOpenFileName(self, 'Open file', '', "Все файлы Excel (*.xlsx)")[0]
        try:
            a = str(main_foo(fname))
            """Тут можно будет в дальнейшем вставить код для анализа текста, который мы получаем из главной функции и выгрузки его допустим в word"""
            self.textEdit.setText(a)

        except FileNotFoundError:
            pass
        except UnboundLocalError:
            msg = QMessageBox()
            msg.setIcon(QMessageBox.Critical)
            msg.setText('Код ошибки №2')
            msg.setInformativeText('Выбран неверный формат файла')
            msg.setStyleSheet("width")
            msg.setWindowTitle('Ошибка')
            msg.setGeometry(350, 350, 500, 500)
            msg.exec_()
        except IndexError:
            msg = QMessageBox()
            msg.setIcon(QMessageBox.Critical)
            msg.setText('Ошибка')
            msg.setInformativeText('Возможно, выбран неверный формат файла')
            msg.setWindowTitle('Ошибка')
            msg.setGeometry(350, 350, 500, 500)
            msg.exec_()

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = Resource_sheet()
    sys.exit(app.exec_())

input()