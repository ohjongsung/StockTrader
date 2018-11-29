import win32com.client
from cybos import cp_util


# 전략 목록 조회 클래스
# https://money2.daishin.com/e5/mboard/ptype_basic/HTS_Plus_Helper/DW_Basic_Read_Page.aspx?boardseq=284&seq=237&page=1&searchString=%EC%A0%84%EB%9E%B5&p=8839&v=8642&m=9508
class CpCssStgList(cp_util.Core):
    def __init__(self):
        self.obj = win32com.client.Dispatch('CpSysDib.CssStgList')
        super(CpCssStgList, self).__init__(self.obj)

    # type 에 해당하는 입력 데이터를 value 값으로 지정합니다.
    # type  value
    # 0     전략 종류 : 0 - 예제전략, 1 - 나의전략
    def set_input_value(self, data_type, value):
        self.obj.SetInputValue(data_type, value)

    # type 에 해당하는 헤더데이터를 반환합니다.
    # type  value
    # 0     전략 목록 수
    # 1     요청구분
    def get_header_value(self, data_type):
        return self.obj.GetHeaderValue(data_type)

    # type 종류의 index 번째에 해당하는 데이터를 반환합니다.
    # type  value
    # 0     전략명
    # 1     전략ID
    # 2     전략 등록 일시
    # 2     작성자 필명
    # 2     평균 종목 수
    # 2     평균 승률
    # 2     평균 수익
    # 2     전략 URL 주소( 해당 url로 웹 페이지에서 전략에 맞게 검색된 종목리스트를 가져올 수 있다)
    def get_data_value(self, data_type, index):
        return self.obj.GetDataValue(data_type, index)

    # 데이터요청. Blocking Mode
    def block_request(self):
        self.obj.BlockRequest()


# 전략에 맞는 종목 검색 클래스
# https://money2.daishin.com/e5/mboard/ptype_basic/HTS_Plus_Helper/DW_Basic_Read_Page.aspx?boardseq=284&seq=238&page=1&searchString=%EC%A0%84%EB%9E%B5&p=8839&v=8642&m=9508
class CpCssStgFind(cp_util.Core):
    def __init__(self):
        self.obj = win32com.client.Dispatch('CpSysDib.CssStgFind')
        super(CpCssStgFind, self).__init__(self.obj)

    # type 에 해당하는 입력 데이터를 value 값으로 지정합니다.
    # type  value
    # 0     전략 ID
    def set_input_value(self, data_type, value):
        self.obj.SetInputValue(data_type, value)

    # type 에 해당하는 헤더데이터를 반환합니다.
    # type  value
    # 0     검색된 결과 종목 수
    # 1     총 검색 종목 수
    # 2     검색 시간
    def get_header_value(self, data_type):
        return self.obj.GetHeaderValue(data_type)

    # type 종류의 index 번째에 해당하는 데이터를 반환합니다.
    # type  value
    # 0     종목 코드
    def get_data_value(self, data_type, index):
        return self.obj.GetDataValue(data_type, index)

    # 데이터요청. Blocking Mode
    def block_request(self):
        self.obj.BlockRequest()


# 종목검색 전략 감시 일련번호 요청 클래스
# https://money2.daishin.com/e5/mboard/ptype_basic/HTS_Plus_Helper/DW_Basic_Read_Page.aspx?boardseq=284&seq=239&page=1&searchString=%EC%A0%84%EB%9E%B5&p=8839&v=8642&m=9508
class CpCssWatchStgSubscribe(cp_util.Core):
    def __init__(self):
        self.obj = win32com.client.Dispatch('CpSysDib.CssWatchStgSubscribe')
        super(CpCssWatchStgSubscribe, self).__init__(self.obj)

    # type 에 해당하는 입력 데이터를 value 값으로 지정합니다.
    # type  value
    # 0     전략 ID
    def set_input_value(self, data_type, value):
        self.obj.SetInputValue(data_type, value)

    # type 에 해당하는 헤더데이터를 반환합니다.
    # type  value
    # 0     해당 전략ID 감시 일련 번호
    def get_header_value(self, data_type):
        return self.obj.GetHeaderValue(data_type)

    # 데이터요청. Blocking Mode
    def block_request(self):
        self.obj.BlockRequest()


# 종목검색 전략 감시 시작/중지 클래스
# https://money2.daishin.com/e5/mboard/ptype_basic/HTS_Plus_Helper/DW_Basic_Read_Page.aspx?boardseq=284&seq=240&page=1&searchString=%EC%A0%84%EB%9E%B5&p=8839&v=8642&m=9508
class CpCssWatchStgControl(cp_util.Core):
    def __init__(self):
        self.obj = win32com.client.Dispatch('CpSysDib.CssWatchStgControl')
        super(CpCssWatchStgControl, self).__init__(self.obj)

    # type 에 해당하는 입력 데이터를 value 값으로 지정합니다.
    # type  value
    # 0     전략 ID
    # 1     감시 일련 번호
    # 2     작업 구분코드 : 1 - 감시시작, 3 - 등록취소
    def set_input_value(self, data_type, value):
        self.obj.SetInputValue(data_type, value)

    # type 에 해당하는 헤더데이터를 반환합니다.
    # type  value
    # 0     감시상태 : 0 - 초기상태, 1 - 감시중, 2 - 감시중단, 3 - 등록취소
    def get_header_value(self, data_type):
        return self.obj.GetHeaderValue(data_type)

    # 데이터요청. Blocking Mode
    def block_request(self):
        return self.obj.BlockRequest()

