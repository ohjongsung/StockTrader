from cybos import cp_util
from cybos import cp_trade
from luncher.daishin import cybos_luncher


# 초기화 서비스 클래스
class InitialService:
    def __init__(self):
        self.is_executed_trade_init = False
        self.CpCybos = cp_util.CpCybos()
        self.CpTdUtil = cp_trade.CpTdUtil()
        self.trade_init()
        # cybos_luncher.cybos_login()

    def check_trader_status(self):
        if self.is_connect():
            return "서버 연결 됨"
        else:
            return "서버 연결 안됨"

    def trade_init(self):
        init_check = self.CpTdUtil.trade_init()
        if init_check != 0:
            print("주문 초기화 실패")
            exit()
        else:
            print("주문 초기화 성공")
            self.is_executed_trade_init = True

    def get_trade_init_status(self):
        return self.is_executed_trade_init

    def is_connect(self):
        return self.CpCybos.is_connect()
