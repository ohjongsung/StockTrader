import win32com.client


# 주문 오브젝트를 사용하기 위해 필요한 초기화 클래스
class CpTdUtil(object):
    def __init__(self):
        self.obj = win32com.client.Dispatch('CpTrade.CpTdUtil')

    # 사용자의 U-CYBOS 로 사인온한 복수계좌목록을스트링 배열로 받아온다.
    def get_account_number(self):
        return self.obj.AccountNumber()

    # 주문을 하기 위한 예비과정을 수행한다.
    def trade_init(self):
        self.obj.TradeInit()

