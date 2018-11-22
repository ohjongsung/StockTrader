import sys
from luncher.daishin import cybos_luncher
from cybos import cp_util
from cybos import cp_trade
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from PyQt5 import uic


form_class = uic.loadUiType("main_window.ui")[0]


class MyWindow(QMainWindow, form_class):
    def __init__(self):
        super().__init__()
        self.setupUi(self)
        self.initTrade = False
        # cybos_luncher.cybos_login()
        self.CpCybos = cp_util.CpCybos()
        self.CpStockCode = cp_util.CpStockCode()
        self.CpTdUtil = cp_trade.CpTdUtil()
        self.CpTdOrder = cp_trade.CpTdOrder()
        self.CpCancelOrder = cp_trade.CpTdCancelOrder()
        self.CpUpdateOrder = cp_trade.CpTdUpdateOrder()
        self.CpConclusion = cp_trade.CpConclusion()

        self.timer = QTimer(self)
        self.timer.start(1000)
        self.timer.timeout.connect(self.timeout)

        self.lineEdit.textChanged.connect(self.code_changed)
        self.pushButton.clicked.connect(self.send_order)

    def timeout(self):
        current_time = QTime.currentTime()
        text_time = current_time.toString("hh:mm:ss")
        time_msg = "현재시간: " + text_time

        if self.CpCybos.is_connect():
            state_msg = "서버 연결 됨"
            trade_init = "Need to execute Trade Init"
            if self.initTrade is False:
                init_check = self.CpTdUtil.trade_init()
                if init_check == 0:
                    account_list = [self.CpTdUtil.get_account_number()[0]]
                    self.comboBox.addItems(account_list)
                    self.initTrade = True
            else:
                trade_init = "Trade Init executed!!"
        else:
            state_msg = "서버 연결 안됨"

        self.statusBar().showMessage(state_msg + " | " + time_msg + " | " + trade_init)

    def code_changed(self):
        code = self.lineEdit.text()
        name = self.CpStockCode.code_to_name(code)
        self.lineEdit_2.setText(name)

    def send_order(self):
        order_type_lookup = {'매도': '1', '매수': '2', '매도취소': '3', '매수취소': '4'}
        hoga_lookup = {'지정가': '00', '시장가': '03 '}

        account = self.comboBox.currentText()
        order_type = self.comboBox_2.currentText()
        code = self.lineEdit.text()
        hoga = self.comboBox_3.currentText()
        num = self.spinBox.value()
        price = self.spinBox_2.value()

        self.CpTdOrder.set_input_value(0, order_type_lookup[order_type])
        self.CpTdOrder.set_input_value(1, account)
        self.CpTdOrder.set_input_value(3, code)
        self.CpTdOrder.set_input_value(4, num)
        self.CpTdOrder.set_input_value(5, price)
        self.CpTdOrder.set_input_value(8, hoga_lookup[hoga])

        self.CpTdOrder.block_request()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    myWindow = MyWindow()
    myWindow.show()
    app.exec_()
