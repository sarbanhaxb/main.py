from PyQt5.QtWidgets import QApplication, QMainWindow, QFileDialog, QMessageBox
import sys
from PyQt5.QtGui import QIcon
from PyQt5.QtWidgets import (QTextEdit, QAction)
from main import main_foo


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
        import xlsxwriter
        fname = QFileDialog.getOpenFileName(self, 'Open file', '', "Все файлы Excel (*.xl*;*.xlsx;*.xlsm;*.xlsb;*.xlam;*.xlts;*.xltm;*.xls;*htm;*.html;*.mht;*.mhtml;*.xml;*.xla;*.xlm;*.xlw;*.odc;*.ods)")[0]
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
        except xlsxwriter.exceptions.FileCreateError:
            msg = QMessageBox()
            msg.setIcon(QMessageBox.Critical)
            msg.setText('Код ошибки №3')
            msg.setInformativeText('Файл use.xlsx занят другим процессом')
            msg.setWindowTitle('Ошибка')
            msg.setGeometry(350, 350, 500, 500)
            msg.exec_()

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = Resource_sheet()
    sys.exit(app.exec_())