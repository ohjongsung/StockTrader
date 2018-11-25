from cybos import cp_trade
from service import account


# 주문 서비스 클래스
class OrderService:
    def __init__(self):
        self.CpTdUtil = cp_trade.CpTdUtil()
        self.CpTdOrder = cp_trade.CpTdOrder()
        self.CpCancelOrder = cp_trade.CpTdCancelOrder()
        self.CpUpdateOrder = cp_trade.CpTdUpdateOrder()
        self.CpConclusion = cp_trade.CpConclusion()
        self.AccountService = account.AccountService()

        self.accountNumber = self.CpTdUtil.get_account_number()[0]
        self.acc_flag = self.CpTdUtil.goods_list(self.accountNumber, 1)[0]

    # 주식 매수 주문
    def buy(self, code, price, amount):
        self.CpTdOrder.set_input_value(0, '2')
        self.CpTdOrder.set_input_value(1, self.accountNumber)
        self.CpTdOrder.set_input_value(2, self.acc_flag)
        self.CpTdOrder.set_input_value(3, code)
        if amount > 0:
            self.CpTdOrder.set_input_value(4, amount)
        else:
            self.CpTdOrder.set_input_value(4, self.AccountService.calculate_buy_stock_amount(price, code))
        self.CpTdOrder.set_input_value(5, int(price))
        self.CpTdOrder.set_input_value(7, '0')
        self.CpTdOrder.set_input_value(8, '01')

        self.CpTdOrder.block_request()

        # 통신 및 통신 에러 처리
        if self.CpTdOrder.get_communication_status() is False:
            return exit()

    # 주식 매도 주문
    def sell(self, code, amount, price):
        self.CpTdOrder.set_input_value(0, "1")  # 1: 매도
        self.CpTdOrder.set_input_value(1, self.accountNumber)  # 계좌번호
        self.CpTdOrder.set_input_value(2, self.acc_flag)  # 상품구분 - 주식 상품 중 첫번째
        self.CpTdOrder.set_input_value(3, code)  # 종목코드
        self.CpTdOrder.set_input_value(4, amount)  # 매도수량
        self.CpTdOrder.set_input_value(5, int(price))  # 주문단가
        self.CpTdOrder.set_input_value(7, "0")  # 주문 조건 구분 코드, 0: 기본
        self.CpTdOrder.set_input_value(8, "01")  # 주문호가 구분코드 - 01: 지정가

        # 매도 주문 요청
        self.CpTdOrder.block_request()

        # 통신 및 통신 에러 처리
        if self.CpTdOrder.get_communication_status() is False:
            pass

    # 주문 취소
    def cancel_order(self, order_number, code):
        self.CpCancelOrder.set_input_value(1, order_number)
        self.CpCancelOrder.set_input_value(2, self.accountNumber)
        self.CpCancelOrder.set_input_value(3, self.acc_flag)
        self.CpCancelOrder.set_input_value(4, code)
        self.CpCancelOrder.set_input_value(5, 0)

        self.CpCancelOrder.block_request()
