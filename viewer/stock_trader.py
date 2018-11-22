import sys
from luncher.daishin import cybos_luncher
from cybos import cp_util
from cybos import cp_trade
from cybos import cp_account
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
        self.CpAccount = cp_account.CpAccount()

        self.timer = QTimer(self)
        self.timer.start(1000)
        self.timer.timeout.connect(self.timeout)

        self.timer2 = QTimer(self)
        self.timer2.start(1000*10)
        self.timer2.timeout.connect(self.timeout2)

        self.lineEdit.textChanged.connect(self.code_changed)
        self.pushButton.clicked.connect(self.send_order)
        self.pushButton_2.clicked.connect(self.check_balance)

    def do_after_trade_init(self):
        account = self.CpTdUtil.get_account_number()[0]
        self.comboBox.addItems([account])
        self.CpAccount.set_input_value(0, account)
        self.CpAccount.set_input_value(2, 50)
        self.CpAccount.set_input_value(3, 1)

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
                    self.initTrade = True
                    self.do_after_trade_init()
            else:
                trade_init = "Trade Init executed!!"
        else:
            state_msg = "서버 연결 안됨"

        self.statusBar().showMessage(state_msg + " | " + time_msg + " | " + trade_init)

    def timeout2(self):
        if self.checkBox.isChecked():
            self.check_balance()

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

    def check_balance(self):
        self.CpAccount.block_request()

        # 통신 및 통신 에러 처리
        request_status = self.CpAccount.get_dib_status()
        request_result = self.CpAccount.get_dib_mgs1()
        print("통신상태", request_status, request_result)
        if request_status != 0:
            return False

        row = []
        row.append(self.CpAccount.get_header_value(0))  # 계좌명
        row.append(self.CpAccount.get_header_value(1))  # 결제잔고수량
        row.append(self.CpAccount.get_header_value(2))  # 체결잔고수량
        row.append(self.CpAccount.get_header_value(3))  # 평가금액
        row.append(self.CpAccount.get_header_value(4))  # 평가손익
        row.append(self.CpAccount.get_header_value(8))  # 수익률 rate of return to investment
        row.append(self.CpAccount.get_header_value(9))  # D+2 예상예수금
        for i in range(len(row)):
            item = QTableWidgetItem(row[i])
            item.setTextAlignment(Qt.AlignVCenter | Qt.AlignRight)
            self.tableWidget.setItem(0, i, item)

        self.tableWidget.resizeRowsToContents()

        cnt = self.CpAccount.get_header_value(7)
        print(cnt)

        if cnt is None:
            return False

        self.tableWidget_2.setRowCount(cnt)

        for i in range(cnt):
            row = []
            row.append(self.CpAccount.get_data_value(12, i))  # 종목코드
            row.append(self.CpAccount.get_data_value(0, i))  # 종목명
            row.append(self.CpAccount.get_data_value(1, i))  # 신용구분
            row.append(self.CpAccount.get_data_value(7, i))  # 체결잔고수량
            row.append(self.CpAccount.get_data_value(17, i))  # 체결장부단가
            row.append(self.CpAccount.get_data_value(9, i))  # 평가금액(천원미만은 절사)
            row.append(self.CpAccount.get_data_value(10, i))  # 평가손익
            row.append(self.CpAccount.get_data_value(11, i))  # 수익률
            for j in range(len(row)):
                item = QTableWidgetItem(row[j])
                item.setTextAlignment(Qt.AlignVCenter | Qt.AlignRight)
                self.tableWidget_2.setItem(i, j, item)

        self.tableWidget_2.resizeRowsToContents()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    myWindow = MyWindow()
    myWindow.show()
    app.exec_()
