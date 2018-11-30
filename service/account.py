from cybos import cp_trade
from cybos import cp_account


# 계정 정보 서비스 클래스
class AccountService:
    def __init__(self):
        self.CpTdUtil = cp_trade.CpTdUtil()
        self.CpAccount = cp_account.CpAccount()
        self.accountNumber = self.CpTdUtil.get_account_number()[0]
        self.acc_flag = self.CpTdUtil.goods_list(self.accountNumber, 1)[0]

    # type  value
    # 0     계좌번호
    # 1     상품관리구분코드
    # 2     요청건수 [default:14] - 최대 50개
    # 3     수익률구분코드  - ( "1" : 100% 기준, "2": 0% 기준)
    def check_balance(self):
        self.CpAccount.set_input_value(0, self.accountNumber)
        self.CpAccount.set_input_value(1, self.acc_flag)
        self.CpAccount.set_input_value(2, 50)
        self.CpAccount.set_input_value(3, 1)
        self.CpAccount.block_request()

        # 통신 및 통신 에러 처리
        if self.CpAccount.get_communication_status() is False:
            return False

        balance = {}
        total_list = []
        total_list.append(self.CpAccount.get_header_value(0))  # 계좌명
        total_list.append(self.accountNumber)
        total_list.append(self.CpAccount.get_header_value(1))  # 결제잔고수량
        total_list.append(self.CpAccount.get_header_value(2))  # 체결잔고수량
        total_list.append(self.CpAccount.get_header_value(3))  # 평가금액
        total_list.append(self.CpAccount.get_header_value(4))  # 평가손익
        cnt = self.CpAccount.get_header_value(7)
        total_list.append(cnt)  # 보유 주식 개수
        total_list.append(self.CpAccount.get_header_value(8))  # 수익률 rate of return to investment
        total_list.append(self.CpAccount.get_header_value(9))  # D+2 예상예수금
        balance['total'] = total_list

        print('보유 주식 : ', cnt)
        stock_list = []
        if cnt is None:
            balance['stock'] = stock_list
            return balance

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
            stock_list.append(row)

        balance['stock'] = stock_list
        return balance

    def get_account_num(self):
        return self.accountNumber

