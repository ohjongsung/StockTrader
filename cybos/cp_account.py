import win32com.client


# 계좌별 잔고 및 주문체결 평가 현황 데이터 조회를 위한 클래스
# https://money2.daishin.com/e5/mboard/ptype_basic/HTS_Plus_Helper/DW_Basic_Read_Page.aspx?boardseq=286&seq=176&page=3&searchString=&p=&v=&m=
class CpAccount(object):
    def __init__(self):
        self.obj = win32com.client.Dispatch('CpTrade.CpTd6033')

    # type 에 해당하는 입력 데이터를 value 값으로 지정합니다.
    # type  value
    # 0     계좌번호
    # 1     상품관리구분코드
    # 2     요청건수 [default:14] - 최대 50개
    # 3     수익률구분코드  - ( "1" : 100% 기준, "2": 0% 기준)
    def set_input_value(self, data_type, value):
        self.obj.SetInputValue(data_type, value)

    # type 에 해당하는 헤더데이터를 반환합니다.
    # type  value
    # 0     계좌명
    # 1     결제잔고수량
    # 2     체결잔고수량
    # 3     평가금액
    # 4     평가손익
    # 5     사용하지 않음
    # 6     대출금액
    # 7     수신개수
    # 8     수익률
    # 9     D+2 예상예수금
    # 10    대주평가금액
    # 11    잔고평가금액
    # 12    대주금액
    def get_header_value(self, data_type):
        self.obj.GetHeaderValue(data_type)

    # type 종류의 index 번째에 해당하는 데이터를 반환합니다.
    # type  value
    # 0     종목명
    # 1     신용구분 : Y - 신용융자/유통융자, D - 신용대주/유통대주, B - 담보대출, M - 매입담보대출, P - 플러스론대출, I - 자기융자/유통융자
    # 2     대출일
    # 3     결제잔고수량
    # 4     결제장부단가
    # 5     전일체결수량
    # 6     금일체결수량
    # 7     체결잔고수량
    # 9     평가금액(단위:원) - 천원미만은내림
    # 10    평가손익(단위:원) - 천원미만은내림
    # 11    수익률
    # 12    종목코드
    # 13    주문구분
    # 15    매도가능수량
    # 16    만기일
    # 17    체결장부단가
    # 18    손익단가
    def get_data_value(self, data_type, index):
        self.obj.GetDataValue(data_type, index)

    def block_request(self):
        self.obj.BlockRequest()

    def get_dib_status(self):
        return self.obj.GetDibStatus()

    def get_dib_mgs1(self):
        return self.obj.GetDibMsg1()

