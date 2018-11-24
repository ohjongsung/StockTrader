import win32com.client
from cybos import cp_util


# 주문 오브젝트를 사용하기 위해 필요한 초기화 클래스
class CpTdUtil(cp_util.Core):
    def __init__(self):
        self.obj = win32com.client.Dispatch('CpTrade.CpTdUtil')
        super(CpTdUtil, self).__init__(self.obj)

    # 사용자의 U-CYBOS 로 사인온한 복수계좌목록을스트링 배열로 받아온다.
    def get_account_number(self):
        return self.obj.AccountNumber

    # 주문을 하기 위한 예비과정을 수행한다.
    def trade_init(self):
        return self.obj.TradeInit()

    # 사인온 한 계좌에 대해서 필터 값에 따른 계좌목록을 배열로 반환한다.
    # -1 : 전체, 1 : 주식, 2 : 선물/옵션 16 : EUREX, 64 : 해외선물
    def goods_list(self, account, index):
        return self.obj.GoodsList(account, index)


# 장내주식/코스닥주식/ELW 현금주문을 위한 클래스
# https://money2.daishin.com/e5/mboard/ptype_basic/HTS_Plus_Helper/DW_Basic_Read_Page.aspx?boardseq=291&seq=159&page=2&searchString=&p=&v=&m=
class CpTdOrder(cp_util.Core):
    def __init__(self):
        self.obj = win32com.client.Dispatch('CpTrade.CpTd0311')
        super(CpTdOrder, self).__init__(self.obj)

    # type 에 해당하는 입력 데이터를 value 값으로 지정합니다.
    # type  value
    # 0     주문종류코드 : 1 - 매도, 2 - 매수
    # 1     계좌번호
    # 2     상품관리구분코드
    # 3     종목코드
    # 4     주문수량
    # 5     주문단가
    # 7     주문조건구분코드 : 0 - 없음, 1 - IOC, 2 - FOK
    # 8     주문호가구분코드(type 6 대체) : 01 - 보통(지정가), 02 - 임의, 03 - 시장가, 05 - 조건부지정가...더 있는데 내가 사용할거 같지 않음
    def set_input_value(self, data_type, value):
        self.obj.SetInputValue(data_type, value)

    # type 에 해당하는 헤더데이터를 반환합니다.
    # type  value
    # 0     주문종류코드 : 1 - 매도, 2 - 매수
    # 1     계좌번호
    # 2     상품관리구분코드
    # 3     종목코드
    # 4     주문수량
    # 5     주문단가
    # 8     주문번호
    # 9     계좌명
    # 10    종목명
    # 12    주문조건구분코드 : 0 - 없음, 1 - IOC, 2 - FOK
    # 13    주문호가구분코드 : 01 - 보통(default), 02 - 임의, 03 - 시장가, 05 - 조건부지정가...더 있는데 내가 사용할거 같지 않음
    def get_header_value(self, data_type):
        self.obj.GetHeaderValue(data_type)

    # hts 장내주식 현금주문 관련 데이터요청. Blocking Mode
    def block_request(self):
        self.obj.BlockRequest()


# 실시간 주문 체결 수신 클래스
class CpConclusion:
    def __init__(self):
        self.obj = win32com.client.Dispatch("DsCbo1.CpConclusion")

    def subscribe(self):
        win32com.client.WithEvents(self.obj, CpConclusionHandler(self.obj))
        self.obj.Subscribe()

    def unsubscribe(self):
        self.obj.Unsubscribe()


# 실시간 주문 체결 이벤트 핸들러
# https://money2.daishin.com/e5/mboard/ptype_basic/HTS_Plus_Helper/DW_Basic_Read_Page.aspx?boardseq=291&seq=155&page=2&searchString=&p=&v=&m=
class CpConclusionHandler:
    def __init__(self, client):
        self.client = client

    def on_received(self):
        conclusion = {}
        pass
        # 계좌명
        conclusion[1] = self.client.GetHeaderValue(1)
        # 종목명
        conclusion[2] = self.client.GetHeaderValue(2)
        # 체결수량
        conclusion[3] = self.client.GetHeaderValue(3)
        # 주문번호
        conclusion[4] = self.client.GetHeaderValue(4)
        # 계좌명
        conclusion[5] = self.client.GetHeaderValue(5)
        # 계좌명
        conclusion[6] = self.client.GetHeaderValue(6)
        # 계좌명
        conclusion[7] = self.client.GetHeaderValue(7)
        # 계좌명
        conclusion[8] = self.client.GetHeaderValue(8)
        # 종목코드
        conclusion[9] = self.client.GetHeaderValue(9)
        # 매매구분코드
        conclusion[12] = '매도' if self.client.GetHeaderValue(12) == '1' else '매수'
        # 체결구분코드
        if self.client.GetHeaderValue(14) == '1':
            conclusion[14] = '체결'
        elif self.client.GetHeaderValue(14) == '2':
            conclusion[14] = '확인'
        elif self.client.GetHeaderValue(14) == '3':
            conclusion[14] = '거부'
        elif self.client.GetHeaderValue(14) == '4':
            conclusion[14] = '접수'
        # 15 - 신용대출구분코드 (사용할 일이 없어서 제외)
        # 정정취소구분코드
        if self.client.GetHeaderValue(16) == '1':
            conclusion[16] = '정상주문'
        elif self.client.GetHeaderValue(16) == '2':
            conclusion[16] = '정정주문'
        elif self.client.GetHeaderValue(16) == '3':
            conclusion[16] = '취소주문'
        # 현금신용대용구분코드
        if self.client.GetHeaderValue(17) == '1':
            conclusion[17] = '현금'
        elif self.client.GetHeaderValue(17) == '2':
            conclusion[17] = '신용'
        elif self.client.GetHeaderValue(17) == '3':
            conclusion[17] = '선물대용'
        elif self.client.GetHeaderValue(17) == '4':
            conclusion[17] = '공매도'
        # 주문호가구분코드 (01~79 중 필요한 것만 작성)
        if self.client.GetHeaderValue(18) == '01':
            conclusion[18] = '보통'
        elif self.client.GetHeaderValue(18) == '02':
            conclusion[18] = '임의'
        elif self.client.GetHeaderValue(18) == '03':
            conclusion[18] = '시장가'
        elif self.client.GetHeaderValue(18) == '05':
            conclusion[18] = '조건부지정가'
        # 주문조건구분코드
        if self.client.GetHeaderValue(19) == '1':
            conclusion[19] = '없음'
        elif self.client.GetHeaderValue(19) == '2':
            conclusion[19] = 'IOC'
        elif self.client.GetHeaderValue(19) == '3':
            conclusion[19] = 'FOK'
        # 대출일
        conclusion[20] = self.client.GetHeaderValue(20)
        # 장부가
        conclusion[21] = self.client.GetHeaderValue(21)
        # 매도가능수량
        conclusion[22] = self.client.GetHeaderValue(22)
        # 체결기준잔고수량
        conclusion[23] = self.client.GetHeaderValue(23)
        print(conclusion)


# 장내주식/코스닥주식/ELW 현금주문 취소를 위한 클래스
# https://money2.daishin.com/e5/mboard/ptype_basic/HTS_Plus_Helper/DW_Basic_Read_Page.aspx?boardseq=291&seq=162&page=1&searchString=&p=&v=&m=
class CpTdCancelOrder(cp_util.Core):
    def __init__(self):
        self.obj = win32com.client.Dispatch('CpTrade.CpTd0314')
        super(CpTdCancelOrder, self).__init__(self.obj)

    # type 에 해당하는 입력 데이터를 value 값으로 지정합니다.
    # type  value
    # 1     원주문번호
    # 2     계좌번호
    # 3     상품관리구분코드
    # 4     종목코드
    # 5     취소수량 (0 입력 시, 가능수량 자동계산됨)
    def set_input_value(self, data_type, value):
        self.obj.SetInputValue(data_type, value)

    # type 에 해당하는 헤더데이터를 반환합니다.
    # type  value
    # 1     원주문번호
    # 2     계좌번호
    # 3     상품관리구분코드
    # 4     종목코드
    # 5     취소수량
    # 6     주문번호
    # 7     계좌명
    # 8     종목명
    def get_header_value(self, data_type):
        self.obj.GetHeaderValue(data_type)

    # hts 장내주식 현금주문 관련 데이터요청. Blocking Mode
    def block_request(self):
        self.obj.BlockRequest()


# 장내주식/코스닥주식/ELW 현금주문 정정을 위한 클래스
# https://money2.daishin.com/e5/mboard/ptype_basic/HTS_Plus_Helper/DW_Basic_Read_Page.aspx?boardseq=291&seq=161&page=2&searchString=&p=&v=&m=
class CpTdUpdateOrder(cp_util.Core):
    def __init__(self):
        self.obj = win32com.client.Dispatch('CpTrade.CpTd0313')
        super(CpTdUpdateOrder, self).__init__(self.obj)

    # type 에 해당하는 입력 데이터를 value 값으로 지정합니다.
    # type  value
    # 1     원주문번호
    # 2     계좌번호
    # 3     상품관리구분코드
    # 4     종목코드
    # 5     주문수량 (0 입력 시, 가능수량 자동계산됨)
    # 5     주문단가
    def set_input_value(self, data_type, value):
        self.obj.SetInputValue(data_type, value)

    # type 에 해당하는 헤더데이터를 반환합니다.
    # type  value
    # 1     원주문번호
    # 2     계좌번호
    # 3     상품관리구분코드
    # 4     종목코드
    # 5     주문수량
    # 6     주문단가
    # 7     주문번호
    # 8     계좌명
    # 9     종목명
    def get_header_value(self, data_type):
        self.obj.GetHeaderValue(data_type)

    # hts 장내주식 현금주문 관련 데이터요청. Blocking Mode
    def block_request(self):
        self.obj.BlockRequest()

