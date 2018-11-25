from cybos import cp_trade
from cybos import cp_account
from service import stock


# 주문 서비스 클래스
class OrderService:
    def __init__(self):
        self.CpTdUtil = cp_trade.CpTdUtil()
        self.CpTdOrder = cp_trade.CpTdOrder()
        self.CpCancelOrder = cp_trade.CpTdCancelOrder()
        self.CpUpdateOrder = cp_trade.CpTdUpdateOrder()
        self.CpConclusion = cp_trade.CpConclusion()
        self.CpAvailableBuy = cp_account.CpAvailableBuy()
        self.StockService = stock.StockService()

        self.accountNumber = self.CpTdUtil.get_account_number()[0]
        self.acc_flag = self.CpTdUtil.goods_list(self.accountNumber, 1)[0]
        self.order_type = {
            'sell': 1,
            'buy': 2,
        }

    # 주식 매수 주문
    def buy(self, code, price, amount):
        self.CpTdOrder.set_input_value(0, '2')
        self.CpTdOrder.set_input_value(1, self.accountNumber)
        self.CpTdOrder.set_input_value(2, self.acc_flag)
        self.CpTdOrder.set_input_value(3, code)

        order_price = price
        if order_price is 0 or order_price is None:
            order_price = self.calculate_order_stock_price(code, self.order_type['buy'])
            self.CpTdOrder.set_input_value(5, order_price)
        else:
            self.CpTdOrder.set_input_value(5, order_price)

        order_amount = amount
        if order_amount is 0 or order_amount is None:
            order_amount = self.calculate_buy_stock_amount(order_price, code)
            self.CpTdOrder.set_input_value(4, order_amount)
        else:
            self.CpTdOrder.set_input_value(4, order_amount)

        self.CpTdOrder.set_input_value(7, '0')
        self.CpTdOrder.set_input_value(8, '01')

        self.CpTdOrder.block_request()

        # 통신 및 통신 에러 처리
        if self.CpTdOrder.get_communication_status() is False:
            return exit()

    # 주식 매도 주문
    def sell(self, code, price, amount):
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

    def calculate_order_stock_price(self, code, order_type):
        price_info = self.StockService.get_current_price(code)
        if price_info is False:
            return False

        if order_type is 1:
            return price_info['sell2']  # 2호가 매도
        elif order_type is 2:
            return price_info['buy2']  # 2호가 매수

    # 주문 가능 수량 계산
    def calculate_buy_stock_amount(self, price, code):
        self.CpAvailableBuy.set_input_value(0, self.accountNumber)  # 계좌번호
        self.CpAvailableBuy.set_input_value(1, self.acc_flag)
        self.CpAvailableBuy.set_input_value(2, code)  # 종목코드
        self.CpAvailableBuy.set_input_value(3, '01')  # 보통가(지정가)
        self.CpAvailableBuy.set_input_value(4, int(price))  # 가격
        self.CpAvailableBuy.set_input_value(6, 2)  # 수량 조회

        self.CpAvailableBuy.block_request()

        money_to_buy = self.CpAvailableBuy.get_header_value(18)  # 현금 주문 가능수량
        amount = self.CpAvailableBuy.get_header_value(45)  # 잔고 호출

        # 매수수량, 잔고 확인 및 리턴
        print(amount)
        print(money_to_buy)

        return money_to_buy
