import win32com.client
from cybos import cp_util


# 주식종목의 현재가에 관련된 데이터(10차 호가 포함) 조회 클래스
# https://money2.daishin.com/e5/mboard/ptype_basic/HTS_Plus_Helper/DW_Basic_Read_Page.aspx?boardseq=284&seq=3&page=1&searchString=DsCbo1.StockMst&p=8839&v=8642&m=9508
class CpStockMst(cp_util.Core):
    def __init__(self):
        self.obj = win32com.client.Dispatch('Dscbo1.StockMst')
        super(CpStockMst, self).__init__(self.obj)

    # type 에 해당하는 입력 데이터를 value 값으로 지정합니다.
    # type  value
    # 0     종목코드
    def set_input_value(self, data_type, value):
        self.obj.SetInputValue(data_type, value)

    # type 에 해당하는 헤더데이터를 반환합니다.
    # type  value
    #  0   종목코드
    #  1   종목명
    #  2   대신업종코드
    #  3   그룹코드
    #  4   시간
    #  5   소속구분(문자열)
    #  6   대형,중형,소형
    #  8   상한가
    #  9   하한가
    #  10  전일종가
    #  11  현재가
    #  12  전일대비
    #  13  시가
    #  14  고가
    #  15  저가
    #  16  매도호가
    #  17  매수호가
    #  18  거래량
    #  19  누적거래대금
    #  추가적으로 필요한건 나중에 필요할때 하자
    def get_header_value(self, data_type):
        self.obj.GetHeaderValue(data_type)

    # type 종류의 index 번째에 해당하는 데이터를 반환합니다.
    # type  value
    # 0     매도호가
    # 1     매수호가
    # 2     매도잔량
    # 3     매수잔량
    # 4     매도잔량대비
    # 5     매수잔량대비
    def get_data_value(self, data_type, index):
        self.obj.GetDataValue(data_type, index)

    def block_request(self):
        self.obj.BlockRequest()
