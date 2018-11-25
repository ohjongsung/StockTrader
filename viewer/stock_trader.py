import sys
from service import order
from service import account
from service import initial
from service import stock
from service import watch
from service import slack
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from PyQt5 import uic


form_class = uic.loadUiType("main_window.ui")[0]


class MyWindow(QMainWindow, form_class):
    def __init__(self):
        super().__init__()
        self.setupUi(self)
        self.InitialService = initial.InitialService()
        self.StockService = stock.StockService()
        self.OrderService = order.OrderService()
        self.AccountService = account.AccountService()
        self.WatchService = watch.WatchService()
        self.Slack = slack.Slack()

        self.timer = QTimer(self)
        self.timer.start(1000)
        self.timer.timeout.connect(self.timeout)

        self.timer2 = QTimer(self)
        self.timer2.start(1000*10)
        self.timer2.timeout.connect(self.timeout2)

        self.timer3 = QTimer(self)
        self.timer3.start(1000*10)
        self.timer3.timeout.connect(self.timeout3)

        self.comboBox.addItems([self.AccountService.get_account_num()])
        self.lineEdit.textChanged.connect(self.code_changed)
        self.pushButton.clicked.connect(self.send_order)
        self.pushButton_2.clicked.connect(self.show_balance)
        self.pushButton_3.clicked.connect(self.get_unusual_stock)

    def timeout(self):
        current_time = QTime.currentTime()
        text_time = current_time.toString("hh:mm:ss")
        time_msg = "현재시간: " + text_time
        state_msg = self.InitialService.check_trader_status()
        self.statusBar().showMessage(state_msg + " | " + time_msg)

    def timeout2(self):
        if self.checkBox.isChecked():
            self.show_balance()

    def timeout3(self):
        if self.checkBox_2.isChecked():
            self.get_unusual_stock()

    def code_changed(self):
        code = self.lineEdit.text()
        name = self.StockService.code_to_name(code)
        self.lineEdit_2.setText(name)

    def send_order(self):
        order_type_lookup = {'매도': '1', '매수': '2', '매도취소': '3', '매수취소': '4'}
        order_type = self.comboBox_2.currentText()

        code = self.lineEdit.text()
        amount = self.spinBox.value()
        price = self.spinBox_2.value()
        if order_type_lookup[order_type] == '1':
            self.OrderService.buy(code, price, amount)
        elif order_type_lookup[order_type] == '2':
            self.OrderService.sell(code, price, amount)

    def show_balance(self):
        balance = self.AccountService.check_balance()
        if balance is False:
            return False

        total = balance['total']
        stock = balance['stock']
        for i in range(len(total)):
            item = QTableWidgetItem(total[i])
            item.setTextAlignment(Qt.AlignVCenter | Qt.AlignRight)
            self.tableWidget.setItem(0, i, item)

        self.tableWidget.resizeRowsToContents()

        cnt = len(stock)
        if cnt is 0:
            return False

        self.tableWidget_2.setRowCount(cnt)

        for i in range(cnt):
            row = stock[i]
            for j in range(len(row)):
                item = QTableWidgetItem(row[j])
                item.setTextAlignment(Qt.AlignVCenter | Qt.AlignRight)
                self.tableWidget_2.setItem(i, j, item)

        self.tableWidget_2.resizeRowsToContents()

    def get_unusual_stock(self):
        unusual_stock_list = self.WatchService.get_unusual_stock()
        if unusual_stock_list is False:
            print('조회된 특징주가 없습니다.')
            return False

        cnt = len(unusual_stock_list)
        if cnt is 0:
            print('조회된 특징주가 없습니다.')
            return False

        self.tableWidget_4.setRowCount(cnt)

        for i in range(cnt):
            row = unusual_stock_list[i]
            for j in range(len(row)):
                item = QTableWidgetItem(row[j])
                item.setTextAlignment(Qt.AlignVCenter | Qt.AlignRight)
                self.tableWidget_4.setItem(i, j, item)

        self.tableWidget_4.resizeRowsToContents()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    myWindow = MyWindow()
    myWindow.show()
    app.exec_()
