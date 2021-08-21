import sys,Ui_untitled,ExcelVerification,threading
from PyQt5.QtWidgets import QApplication, QFileDialog, QMainWindow, QPushButton


def GetFile():
    name = QFileDialog.getOpenFileName(None,"选择文件", "/", "xlsx files (*.xlsx);;xls files (*.xls);;all files (*)")
    if name[0] != "":
        t = threading.Thread(target=ExcelVerification.main,args=(name[0],),daemon = True)
        t.start()


def main():
    app = QApplication(sys.argv)
    MainWindow = QMainWindow()
    ui = Ui_untitled.Ui_MainWindow()
    ui.setupUi(MainWindow)
    MainWindow.show()
    ui.getfilebutton.clicked.connect(GetFile)
    ui.exitbutton.clicked.connect(sys.exit)
    sys.exit(app.exec_())


if __name__ == '__main__':
    main()