import win32com.client
from cybos import cp_util


# 주식, 업종, ELW 의 차트 데이터를 수신하는 클래스
# https://money2.daishin.com/e5/mboard/ptype_basic/HTS_Plus_Helper/DW_Basic_Read_Page.aspx?boardseq=284&seq=102&page=1&searchString=StockChart&p=8839&v=8642&m=9508
class CpData(cp_util.Core):
    def __init__(self):
        self.obj = win32com.client.Dispatch('CpSysDib.StockChart')
        super(CpData, self).__init__(self.obj)

    # type 에 해당하는 입력 데이터를 value 값으로 지정합니다.
    # type  value
    # 0     종목코드
    # 1     요청 구분 : 1 - 기간, 2 - 개수
    # 2     요청 종료일 : YYYYMMDD
    # 3     요청 시작일 : YYYYMMDD
    # 4     요청 개수
    # 5     요청 필드 : 홈피 참고
    # 6     차트 구분 : D -일, W - 주, M - 월, m - 분, T - 틱
    # 9     수정 주가 : 0 - 무수정, 1 - 수정
    # 10    거래량 구분 : 1 - 시간외거래량모두포함[default], 2 - 장종료시간외거래량만포함, 3 - 시간외거래량모두제외, 4 - 장전시간외거래량만 포함
    def set_input_value(self, data_type, value):
        self.obj.SetInputValue(data_type, value)

    # type 에 해당하는 헤더데이터를 반환합니다.
    # type  value
    #  3   수신 개수
    def get_header_value(self, data_type):
        return self.obj.GetHeaderValue(data_type)

    # type 종류의 index 번째에 해당하는 데이터를 반환합니다.
    # type 요청한필드의 index - 필드는 요청한 필드 값으로 오름차순으로정 렬되어 있음
    # index  요청한종목의 index
    def get_data_value(self, data_type, index):
        return self.obj.GetDataValue(data_type, index)

    def block_request(self):
        self.obj.BlockRequest()
