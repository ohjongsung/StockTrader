import win32com.client
from cybos import cp_strategy
from service import stock
from service import slack


# 주식 정보 서비스 클래스
class StrategyService:
    def __init__(self):
        self.CpCssStgList = cp_strategy.CpCssStgList()
        self.CpCssStgFind = cp_strategy.CpCssStgFind()
        self.CpCssWatchStgControl = cp_strategy.CpCssWatchStgControl()
        self.CpCssAlert = CpCssAlert()
        self.StockService = stock.StockService()
        self.strategy_monitor_ids = {}
        self.isMonitoring = False

    def get_my_strategy(self):
        self.CpCssStgList.set_input_value('0', '1')
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
            item['time'] = search_time
            item['code'] = self.CpCssStgFind.get_data_value(0, i)
            item['종목명'] = self.StockService.code_to_name(item['code'])
            stock_list.append(item)

        return True, stock_list

    def get_monitoring_id(self, strategy_id):
        CpCssWatchStgSubscribe = cp_strategy.CpCssWatchStgSubscribe()
        CpCssWatchStgSubscribe.set_input_value(0, strategy_id)
        CpCssWatchStgSubscribe.block_request()

        CpCssWatchStgSubscribe.get_communication_status()

        monitoring_id = CpCssWatchStgSubscribe.get_header_value(0)
        if monitoring_id is 0:
            print("감시 일련번호 얻기 실패")
            return False
        elif monitoring_id is 1:
            print("현재 감시중인 전략이 있습니다.")
            return False

        return True, monitoring_id

    def request_monitoring(self, strategy_id, monitoring_id, is_start, caller):
        self.CpCssWatchStgControl.set_input_value(0, strategy_id)
        self.CpCssWatchStgControl.set_input_value(1, monitoring_id)

        if is_start is True:
            self.CpCssWatchStgControl.set_input_value(2, ord('1'))  # 감시시작
            print('전략감시 시작 요청 ', '전략 ID:', strategy_id, '감시일련번호', monitoring_id)
        else:
            self.CpCssWatchStgControl.set_input_value(2, ord('3'))  # 감시취소
            print('전략감시 취소 요청 ', '전략 ID:', strategy_id, '감시일련번호', monitoring_id)

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

        if self.isMonitoring is False:
            self.CpCssAlert.subscribe(caller)
            self.isMonitoring = True

        # 진행 중인 전략들 저장
        if is_start is True:
            self.strategy_monitor_ids[strategy_id] = monitoring_id
        else:
            if strategy_id in self.strategy_monitor_ids:
                del self.strategy_monitor_ids[strategy_id]

        return True, status

    def monitoring_all_stop(self):
        self.stop_all_strategy_control()
        if self.isMonitoring:
            self.CpCssAlert.unsubscribe()
            self.isMonitoring = False

    def stop_all_strategy_control(self):
        delitem = []
        for strategy_id, monitoring_id in self.strategy_monitor_ids.items():
            delitem.append((strategy_id, monitoring_id))

        for item in delitem:
            self.request_monitoring(item[0], item[1], False, None)

        print(len(self.strategy_monitor_ids))


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
    def get_header_value(self, data_type):
        self.obj.GetHeaderValue(data_type)

    def subscribe(self, caller):
        self.obj.Unsubscribe()
        handler = win32com.client.WithEvents(self.obj, CpStrategyWatchHandler)
        handler.set_client(self.obj, caller)
        self.obj.Subscribe()

    def unsubscribe(self):
        self.obj.Unsubscribe()


# 전략주 실시간 이벤트 핸들러
class CpStrategyWatchHandler:
    def __init__(self):
        self.StockService = stock.StockService()
        self.Slack = slack.Slack()

    def set_client(self, client, caller):
        self.client = client
        self.caller = caller

    def on_received(self):
        stock = {}
        stock['전략ID'] = self.client.get_header_value(0)
        stock['감시일련번호'] = self.client.get_header_value(1)
        code = stock['code'] = self.client.get_header_value(2)
        stock['종목명'] = self.StockService.code_to_name(code)

        in_out_flag = self.client.get_header_value(3)
        if ord('1') is in_out_flag:
            stock['INOUT'] = '진입'
        elif ord('2') is in_out_flag:
            stock['INOUT'] = '퇴출'
        stock['time'] = self.client.get_header_value(4)
        stock['현재가'] = self.client.get_header_value(5)

        message = stock['code'] + '/' + stock['종목명'] + ' ' + stock['INOUT'] + ' ' + stock['time'] + ' ' + stock['현재가']
        print(message)
        self.Slack.push(message)
        self.caller.change_strategy_stocks(stock)
