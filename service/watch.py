from cybos import cp_watch


# 특징주 포착 서비스 클래스
class WatchService:
    def __init__(self):
        self.CpMarketWatchPubSub = cp_watch.CpMarketWatchPubSub()
        self.CpMarketWatchReqRes = cp_watch.CpMarketWatchReqRes()
        self.unusual_status_dic = {
            1: '종목뉴스',
            2: '공시정보',
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

    def start_watch(self):
        self.CpMarketWatchPubSub.unsubscribe()

    def stop_watch(self):
        self.CpMarketWatchPubSub.subscribe()

    def get_unusual_stock(self):
        self.CpMarketWatchReqRes.set_input_value(0, '*')
        # 1: 종목 뉴스 2: 공시정보 10: 외국계 창구첫매수, 11:첫매도 12 외국인 순매수 13 외국인 순매도
        request_field = '10,11,12,13'
        self.CpMarketWatchReqRes.set_input_value(1, request_field)
        self.CpMarketWatchReqRes.set_input_value(2, 0)

        self.CpMarketWatchReqRes.block_request()

        # 통신 및 통신 에러 처리
        if self.CpMarketWatchReqRes.get_communication_status() is False:
            return False

        unusual_stock_list = []
        cnt = self.CpMarketWatchReqRes.get_header_value(2)  # 수신 개수
        if cnt is None:
            return unusual_stock_list

        for i in range(cnt):
            row = []
            time = self.CpMarketWatchReqRes.get_data_value(0, i)
            h, m = divmod(time, 100)
            row.append('%02d:%02d' % (h, m))
            row.append(self.CpMarketWatchReqRes.get_data_value(1, i))
            row.append(self.CpMarketWatchReqRes.get_data_value(2, i))
            cate = self.CpMarketWatchReqRes.get_data_value(3, i)
            row.append(self.unusual_status_dic[cate])
            row.append(self.CpMarketWatchReqRes.get_data_value(4, i))
            unusual_stock_list.append(row)
            print(row)

        return unusual_stock_list
