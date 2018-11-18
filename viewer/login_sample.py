import sys
from PyQt5.QtWidgets import *
import win32com.client
from PyQt5.QtGui import *
from PyQt5.QAxContainer import *
import luncher.daishin.cybos_luncher


class MyWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Login Sample")
        self.setGeometry(300, 300, 300, 150)

        btn1 = QPushButton("Login", self)
        btn1.move(20, 20)
        btn1.clicked.connect(self.btn1_clicked)

        btn2 = QPushButton("Check state", self)
        btn2.move(20, 70)
        btn2.clicked.connect(self.btn2_clicked)

    def btn1_clicked(self):
        luncher.daishin.cybos_luncher.cybos_login()

    def btn2_clicked(self):
        obj_cybos = win32com.client.Dispatch("CpUtil.CpCybos")
        if obj_cybos.IsConnect == 1:
            self.statusBar().showMessage("연결됨")
        else:
            self.statusBar().showMessage("연결되지 않음")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    myWindow = MyWindow()
    myWindow.show()
    app.exec_()


