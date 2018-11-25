import win32com.client
from cybos import cp_util


# 특정 주식종목이나 주식 전종목에 대한 특징주 포착 데이터 Pub/Sub 방식
# https://money2.daishin.com/e5/mboard/ptype_basic/HTS_Plus_Helper/DW_Basic_Read_Page.aspx?boardseq=284&seq=104&page=1&searchString=CpMarketWatch&p=8839&v=8642&m=9508
class CpMarketWatchPubSub(cp_util.Core):
    def __init__(self):
        self.obj = win32com.client.Dispatch('CpSysDib.CpMarketWatchS')
        super(CpMarketWatchPubSub, self).__init__(self.obj)

    # type 에 해당하는 입력 데이터를 value 값으로 지정합니다.
    # type  value
    # 0     종목코드 : * - 전종목
    def set_input_value(self):
        self.obj.SetInputValue(0, '*')

    # type 에 해당하는 헤더데이터를 반환합니다.
    # type  value
    # 0     종목코드
    # 1     종목명
    # 2     Count
    def get_header_value(self, data_type):
        return self.obj.GetHeaderValue(data_type)

    # type 종류의 index 번째에 해당하는 데이터를 반환합니다.
    # type  value
    # 0     시간
    # 1     작업구분 : n - 신규, c - 취소
    # 2     특이사항 코드 : unusual_status_dic 참고
    def get_data_value(self, data_type, index):
        self.obj.GetDataValue(data_type, index)

    def subscribe(self):
        win32com.client.WithEvents(self.obj, CpMarketWatchHandler(self.obj))
        self.obj.Subscribe()

    def unsubscribe(self):
        self.obj.Unsubscribe()


# 특징주 포착 이벤트 핸들러
class CpMarketWatchHandler:
    def __init__(self, client):
        self.client = client
        self.unusual_status_dic = {
            10: '외국계증권사창구첫매수',
            11: '외국계증권사창구첫매도',
            12: '외국인순매수',
            13: '외국인순매도',
            21: '전일거래량갱신',
            22: '최근5일거래량최고갱신',
            23: '최근5일매물대돌파',
            24: '최근60일매물대돌파',
            28: '최근5일첫상한가',
            29: '최근5일신고가갱신',
            30: '최근5일신저가갱신',
            31: '상한가직전',
            32: '하한가직전',
            41: '주가 5MA 상향돌파',
            42: '주가 5MA 하향돌파',
            43: '거래량 5MA 상향돌파',
            44: '주가데드크로스(5MA < 20MA)',
            45: '주가골든크로스(5MA > 20MA)',
            46: 'MACD 매수-Signal(9) 상향돌파',
            47: 'MACD 매도-Signal(9) 하향돌파',
            48: 'CCI 매수-기준선(-100) 상향돌파',
            49: 'CCI 매도-기준선(100) 하향돌파',
            50: 'Stochastic(10,5,5)매수- 기준선상향돌파',
            51: 'Stochastic(10,5,5)매도- 기준선하향돌파',
            52: 'Stochastic(10,5,5)매수- %K%D 교차',
            53: 'Stochastic(10,5,5)매도- %K%D 교차',
            54: 'Sonar 매수-Signal(9) 상향돌파',
            55: 'Sonar 매도-Signal(9) 하향돌파',
            56: 'Momentum 매수-기준선(100) 상향돌파',
            57: 'Momentum 매도-기준선(100) 하향돌파',
            58: 'RSI(14) 매수-Signal(9) 상향돌파',
            59: 'RSI(14) 매도-Signal(9) 하향돌파',
            60: 'Volume Oscillator 매수-Signal(9) 상향돌파',
            61: 'Volume Oscillator 매도-Signal(9) 하향돌파',
            62: 'Price roc 매수-Signal(9) 상향돌파',
            63: 'Price roc 매도-Signal(9) 하향돌파',
            64: '일목균형표매수-전환선 > 기준선상향교차',
            65: '일목균형표매도-전환선 < 기준선하향교차',
            66: '일목균형표매수-주가가선행스팬상향돌파',
            67: '일목균형표매도-주가가선행스팬하향돌파',
            68: '삼선전환도-양전환',
            69: '삼선전환도-음전환',
            70: '캔들패턴-상승반전형',
            71: '캔들패턴-하락반전형',
            81: '단기급락후 5MA 상향돌파',
            82: '주가이동평균밀집-5%이내',
            83: '눌림목재상승-20MA 지지'
        }

    def on_received(self):
        code = self.client.GetHeaderValue(0)
        name = self.client.GetHeaderValue(1)
        cnt = self.client.GetHeaderValue(2)
        unusual_stock_list = []
        for i in range(cnt):
            row = []
            time = self.client.GetDataValue(0, i)
            h, m = divmod(time, 100)
            update = self.client.GetDataValue(1, i)
            cate = self.client.GetDataValue(2, i)

            row.append('%02d:%02d' % (h, m))
            row.append(code)
            row.append(name)
            row.append(self.unusual_status_dic[cate])

            if update == ord('c'):
                new_cancel = '[취소]'
            else:
                new_cancel = '[신규]'
            if cate in self.unusual_status_dic:
                row.append(new_cancel + self.unusual_status_dic[cate])
            else:
                row.append(new_cancel + '')

            unusual_stock_list.append(row)
            print(row)

        return unusual_stock_list


# 특정 주식종목이나 주식 전종목에 대한 특징주 포착 데이터 Req/Res 방식
# https://money2.daishin.com/e5/mboard/ptype_basic/HTS_Plus_Helper/DW_Basic_Read_Page.aspx?boardseq=284&seq=103&page=1&searchString=CpMarketWatch&p=8839&v=8642&m=9508
class CpMarketWatchReqRes(cp_util.Core):
    def __init__(self):
        self.obj = win32com.client.Dispatch('CpSysDib.CpMarketWatch')
        super(CpMarketWatchReqRes, self).__init__(self.obj)

    # type 에 해당하는 입력 데이터를 value 값으로 지정합니다.
    # type  value
    # 0     종목코드 : * - 전종목
    # 1     수신 항목 구분 목록 (구분자 ,) unusual_status_dic 참고
    # 2     시작시간 (요청 시작 시간, 0 [default] - 처음부터)
    def set_input_value(self, data_type, value):
        self.obj.SetInputValue(data_type, value)

    # type 에 해당하는 헤더데이터를 반환합니다.
    # type  value
    # 0     수신 항목 구분 목록
    # 1     시작시간
    # 2     수신개수
    def get_header_value(self, data_type):
        self.obj.GetHeaderValue(data_type)

    # type 종류의 index 번째에 해당하는 데이터를 반환합니다.
    # type  value
    # 0     시간
    # 1     종목코드
    # 2     종목명
    # 3     항목구분
    # 4     내용
    def get_data_value(self, data_type, index):
        self.obj.GetDataValue(data_type, index)

    def block_request(self):
        self.obj.BlockRequest()
