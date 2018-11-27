import win32com.client
from cybos import cp_strategy
from service import stock
from service import slack


# 주식 정보 서비스 클래스
class StrategyService:
    def __init__(self):
        self.CpCssStgList = cp_strategy.CpCssStgList()
        self.CpCssStgFind = cp_strategy.CpCssStgFind()
        self.CpCssWatchStgSubscribe = cp_strategy.CpCssWatchStgSubscribe()
        self.CpCssWatchStgControl = cp_strategy.CpCssWatchStgControl()
        self.CpCssAlert = CpCssAlert()
        self.StockService = stock.StockService()

        self.isMonitoring = False

    def get_my_strategy(self):
        self.CpCssStgList.set_input_value(0, 1)
        self.CpCssStgList.block_request()

        self.CpCssStgList.get_communication_status()

        cnt = self.CpCssStgList.get_header_value(0)
        print("검색된 나의 전략 개수:", cnt)
        if cnt is None:
            exit()

        my_strategy_list = {}
        for i in range(cnt):
            item = {}
            item['전략명'] = self.CpCssStgList.get_data_value(0, i)
            item['ID'] = self.CpCssStgList.get_data_value(1, i)
            item['전략등록일시'] = self.CpCssStgList.get_data_value(2, i)
            item['작성자필명'] = self.CpCssStgList.get_data_value(3, i)
            item['평균종목수'] = self.CpCssStgList.get_data_value(4, i)
            item['평균승률'] = self.CpCssStgList.get_data_value(5, i)
            item['평균수익'] = self.CpCssStgList.get_data_value(6, i)
            my_strategy_list[item['전략명']] = item
            print(item)

        return my_strategy_list

    def search_strategy_stock(self, strategy_id):
        self.CpCssStgFind.set_input_value(0, strategy_id)
        self.CpCssStgFind.block_request()

        self.CpCssStgFind.get_communication_status()

        cnt = self.CpCssStgFind.get_header_value(0)
        total_cnt = self.CpCssStgFind.get_header_value(1)
        search_time = self.CpCssStgFind.get_header_value(2)
        print('검색된 종목수:', cnt, '전체종목수:', total_cnt, '검색시간:', search_time)

        stock_list = []
        for i in range(cnt):
            item = {}
            item['code'] = self.CpCssStgFind.get_data_value(0, i)
            item['종목명'] = self.StockService.code_to_name(item['code'])
            stock_list.append(item)

        return stock_list

    def get_monitoring_id(self, strategy_id):
        self.CpCssWatchStgSubscribe.set_input_value(0, strategy_id)
        self.CpCssWatchStgSubscribe.block_request()

        self.CpCssWatchStgSubscribe.get_communication_status()

        monitoring_id = self.CpCssWatchStgSubscribe.get_header_value(0)
        if monitoring_id is 0:
            print("감시 일련번호 얻기 실패")
            return False
        elif monitoring_id is 1:
            print("현재 감시중인 전략이 있습니다.")
            return False

        return monitoring_id

    def monitoring_start(self, strategy_id, monitoring_id):
        if self.isMonitoring:
            print('이미 전략 감시중입니다.')
            return False

        self.CpCssWatchStgControl.set_input_value(0, strategy_id)
        self.CpCssWatchStgControl.set_input_value(1, monitoring_id)
        self.CpCssWatchStgControl.set_input_value(2, 1)

        self.CpCssWatchStgControl.block_request()

        self.CpCssWatchStgControl.get_communication_status()

        status = self.CpCssWatchStgControl.get_header_value(0)
        if status == 0:
            print('전략감시상태: 초기상태')
        elif status == 1:
            print('전략감시상태: 감시중')
        elif status == 2:
            print('전략감시상태: 감시중단')
        elif status == 3:
            print('전략감시상태: 등록취소')

        self.CpCssAlert.subscribe()
        self.isMonitoring = True
        return True

    def monitoring_stop(self, strategy_id, monitoring_id):
        if self.isMonitoring is False:
            print('이미 전략 감시 중단상태입니다.')
            return False

        self.CpCssWatchStgControl.set_input_value(0, strategy_id)
        self.CpCssWatchStgControl.set_input_value(1, monitoring_id)
        self.CpCssWatchStgControl.set_input_value(2, 3)

        self.CpCssWatchStgControl.block_request()

        self.CpCssWatchStgControl.get_communication_status()

        status = self.CpCssWatchStgControl.get_header_value(0)
        if status == 0:
            print('전략감시상태: 초기상태')
        elif status == 1:
            print('전략감시상태: 감시중')
        elif status == 2:
            print('전략감시상태: 감시중단')
        elif status == 3:
            print('전략감시상태: 등록취소')

        self.CpCssAlert.unsubscribe()
        self.isMonitoring = False
        return True


# 종목검색 실시간 신호 감지 클래스
# https://money2.daishin.com/e5/mboard/ptype_basic/HTS_Plus_Helper/DW_Basic_Read_Page.aspx?boardseq=284&seq=241&page=1&searchString=%EC%A0%84%EB%9E%B5&p=8839&v=8642&m=9508
class CpCssAlert:
    def __init__(self):
        self.obj = win32com.client.Dispatch('CpSysDib.CssAlert')

    # type 에 해당하는 입력 데이터를 value 값으로 지정합니다.
    # type  value
    # 0     전략 ID
    # 1     감시 일련 번호
    # 2     단축종목코드
    # 3     전입/전출 구분 : 1 - 전입, 2 - 전출
    # 4     신호발생시각 HHMMSS(거래소 체결시세 시간)
    # 5     현재가
    def set_input_value(self, data_type, value):
        self.obj.SetInputValue(data_type, value)

    def subscribe(self):
        win32com.client.WithEvents(self.obj, CpStrategyWatchHandler(self.obj))
        self.obj.Subscribe()

    def unsubscribe(self):
        self.obj.Unsubscribe()


# 전략주 실시간 이벤트 핸들러
class CpStrategyWatchHandler:
    def __init__(self, client):
        self.client = client
        self.StockService = stock.StockService()
        self.Slack = slack.Slack()

    def on_received(self):
        stocks = {}
        stocks['전략ID'] = self.client.GetHeaderValue(0)
        stocks['감시일련번호'] = self.client.GetHeaderValue(1)
        code = stocks['code'] = self.client.GetHeaderValue(2)
        stocks['종목명'] =  self.StockService.code_to_name(code)

        in_out_flag = self.client.GetHeaderValue(3)
        if ord('1') is in_out_flag:
            stocks['INOUT'] = '진입'
        elif ord('2') is in_out_flag:
            stocks['INOUT'] = '퇴출'
        stocks['시각'] = self.client.GetHeaderValue(4)
        stocks['현재가'] = self.client.GetHeaderValue(5)

        message = stocks['code'] + '/' + stocks['종목명'] + ' ' + stocks['INOUT'] + ' ' + stocks['시각'] + ' ' + stocks['현재가']
        self.Slack.push(message)
