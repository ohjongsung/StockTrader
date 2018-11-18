import sys
from luncher.daishin import cybos_luncher
from cybos import cp_util
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from PyQt5 import uic


form_class = uic.loadUiType("main_window.ui")[0]


class MyWindow(QMainWindow, form_class):
    def __init__(self):
        super().__init__()
        self.setupUi(self)
        cybos_luncher.cybos_login()
        self.cp_cybos = cp_util.CpCybos()

        self.timer = QTimer(self)
        self.timer.start(1000)
        self.timer.timeout.connect(self.timeout)

    def timeout(self):
        current_time = QTime.currentTime()
        text_time = current_time.toString("hh:mm:ss")
        time_msg = "현재시간: " + text_time

        if self.cp_cybos.is_connect():

            state_msg = "서버 연결 중"
        else:
            state_msg = "서버 연결 안됨"

        self.statusBar().showMessage(state_msg + " | " + time_msg)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    myWindow = MyWindow()
    myWindow.show()
    app.exec_()
