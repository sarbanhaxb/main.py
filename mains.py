from PyQt5 import QtWidgets
from PyQt5.QtWidgets import QApplication, QMainWindow, QFileDialog
import sys, os
import xlsxwriter

from main import lol

class Example(QMainWindow):
    def __init__(self):
        super().__init__()
        self.application()
        self.add_button()

    def add_button(self):
        # Кнопка добавления файла ресурсной ведомости
        fname = QFileDialog.getOpenFileName()[0]
        x = r"'" + str(os.path.abspath(fname)) + "'"
        print(x)

    def application(self):
        app = QApplication(sys.argv)
        window = QMainWindow()

        window.setWindowTitle('Разбор ресурсной ведомости')
        window.setGeometry(250, 250, 300, 250)

        main_btn = QtWidgets.QPushButton(window)
        main_btn.setText('Загрузить ресурсную ведомость')
        main_btn.move(70, 100)
        main_btn.adjustSize()
        main_btn.clicked.connect(self)

        window.show()
        sys.exit(app.exec_())



if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = Example()
    sys.exit(app.exec_())