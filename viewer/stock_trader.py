import sys
from service import order
from service import account
from service import initial
from service import stock
from service import watch
from service import slack
from service import strategy
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
        self.StrategyService = strategy.StrategyService()
        self.my_strategy_list = {}
        self.strategy_stocks = []

        self.timer = QTimer(self)
        self.timer.start(1000)
        self.timer.timeout.connect(self.timeout)

        self.timer2 = QTimer(self)
        self.timer2.start(1000*10)
        self.timer2.timeout.connect(self.timeout2)

        self.comboBox.addItems([self.AccountService.get_account_num()])
        self.lineEdit.textChanged.connect(self.code_changed)
        self.pushButton.clicked.connect(self.send_order)
        self.pushButton_2.clicked.connect(self.show_balance)
        self.pushButton_3.clicked.connect(self.get_unusual_stock)
        self.comboBox_3.currentIndexChanged.connect(self.change_strategy)
        # 전략리스트 조회
        self.list_up_my_strategy()

    def list_up_my_strategy(self):
        self.comboBox_3.addItem('전략선택없음')
        self.my_strategy_list = self.StrategyService.get_my_strategy()

        for k, v in self.my_strategy_list.items():
            self.comboBox_3.addItem(k)

    def change_strategy(self):
        strategy_name = self.comboBox_3.currentText()
        print(strategy_name)
        if strategy_name == '전략선택없음':
            return

        # 1: 기존 감시 중단 (중요)
        # 종목검색 실시간 감시 개수 제한이 있어, 불필요한 감시는 중단이 필요
        self.StrategyService.monitoring_all_stop()

        # 2 - 종목검색 조회: CpSysDib.CssStgFind
        item = self.my_strategy_list[strategy_name]
        id = item['ID']
        name = item['전략명']

        result, self.strategy_stocks = self.StrategyService.search_strategy_stock(id)
        if result is False:
            return

        for item in self.strategy_stocks:
            print(item)
        print('검색전략:', id, '전략명:', name, '검색종목수:', len(self.strategy_stocks))

        if len(self.strategy_stocks) >= 200:
            print('검색종목이 200 을 초과할 경우 실시간 감시 불가 ')
            return

        #####################################################
        # 실시간 요청
        # 3 - 전략의 감시 일련번호 요청 : CssWatchStgSubscribe
        result, monitor_id = self.StrategyService.get_monitoring_id(id)
        if result is False:
            return
        print('감시일련번호', monitor_id)

        # 4 - 전략 감시 시작 요청 - CpSysDib.CssWatchStgControl
        result, status = self.StrategyService.request_monitoring(id, monitor_id, True)
        if result is False:
            return

        return

    def timeout(self):
        current_time = QTime.currentTime()
        text_time = current_time.toString("hh:mm:ss")
        time_msg = "현재시간: " + text_time
        state_msg = self.InitialService.check_trader_status()
        self.statusBar().showMessage(state_msg + " | " + time_msg)

    def timeout2(self):
        if self.checkBox.isChecked():
            self.show_balance()

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
