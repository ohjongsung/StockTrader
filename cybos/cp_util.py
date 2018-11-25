import win32com.client


# CYBOS 통신 결과 수신 코어 클래스
class Core(object):
    def __init__(self, obj):
        self.obj = obj

    def get_dib_status(self):
        return self.obj.GetDibStatus()

    def get_dib_mgs1(self):
        return self.obj.GetDibMsg1()

    def get_communication_status(self):
        # 통신 및 통신 에러 처리
        request_status = self.get_dib_status()
        request_result = self.get_dib_mgs1()
        print("통신상태", request_status, request_result)
        if request_status != 0:
            return False
        else:
            return True


# CYBOS 의 상태 값 확인 클래스
class CpCybos(Core):
    # limitType: 요쳥에 대한 제한타입
    # 주문관련 RQ 요청
    LT_TRADE_REQUEST = 0
    # 시세관련 RQ 요청
    LT_NONTRADE_REQUEST = 1
    # 시세관련 SB
    LT_SUBSCRIBE = 2

    def __init__(self):
        self.obj = win32com.client.Dispatch('CpUtil.CpCybos')
        super(CpCybos, self).__init__(self.obj)

    # CYBOS 의 통신연결상태를 반환합니다.
    # 0 - 연결 끊김, 1 - 연결 정상
    def is_connect(self):
        return True if self.obj.IsConnect == 1 else False

    # 연결된 서버종류를 반환합니다.
    # 0 - 연결 끊김, 1 - cybosplus 서버, 2 - HTS 보통서버(cybosplus 서버제외)
    def get_server_type(self):
        return self.obj.ServerType

    # 요청 개수를 재계산하기까지 남은 시간을 반환합니다.
    # 즉 리턴한 시간동안 남은 요청 개수보다 더 요청하면 요청제한이 됩니다.
    # 요청 개수를 재계산하기까지 남은 시간 (단위:milisecond)
    def get_limit_request_remain_time(self):
        return self.obj.LimitRequestRemainTime

    # limitType 에 대해 제한을 하기전까지 남은 요청 개수를 반환합니다.
    def get_limit_remain_count(self, limit_type):
        return self.obj.GetLimitRemainCount(limit_type)


# 주식 코드 조회 클래스
class CpStockCode(Core):
    def __init__(self):
        self.obj = win32com.client.Dispatch('CpUtil.CpStockCode')
        super(CpStockCode, self).__init__(self.obj)

    # code 에 해당하는 종목 명을 반환합니다.
    def code_to_name(self, code):
        return self.obj.CodeToName(code)

    # name 에 해당하는 종목 코드를 반환합니다.
    def name_to_code(self, name):
        return self.obj.NameToCode(name)

    # code 에 해당하는 full code 를 반환합니다.
    def code_to_full_code(self, code):
        return self.obj.CodeToFullCode(code)

    def full_code_to_name(self, full_code):
        return self.obj.FullCodeToName(full_code)

    def full_code_to_code(self, full_code):
        return self.obj.FullCodeToCode(full_code)

    # code 에 해당하는 종목 index 를 반환합니다.
    def code_to_index(self, code):
        return self.obj.CodeToIndex(code)

    # 종목 코드 수를 반환한다.
    def get_count(self):
        return self.obj.GetCount()

    # 해당 index 의 종목 데이터를 구한다.
    # data_type : 0 - code, 1 - 종목명, 2 - full code
    # idx : 종목 index
    def get_data(self, data_type, idx):
        return self.obj.GetData(data_type, idx)

    # 주식/ETF/ELW 의 호가 단위를 구한다.
    # base_price : 기준 가격
    # direction_up : 상승의 단위인가의 여부(뭔소리여??)  (boolean 값, default = true)
    def get_price_unit(self, code, base_price, direction_up):
        return self.obj.GetPriceUnit(code, base_price, direction_up)


# 주식 코드 정보 및 코드 리스트 조회 클래스
class CpCodeMgr(Core):
    def __init__(self):
        self.obj = win32com.client.Dispatch('CpUtil.CpCodeMgr')
        super(CpCodeMgr, self).__init__(self.obj)

    # code 에 해당하는 주식/선물/옵션 종목명을 반환한다.
    def code_to_name(self, code):
        return self.obj.CodeToName(code)

    # code 에 해당하는 주식매수 증거금율을 반환한다.
    def get_stock_margin_rate(self, code):
        return self.obj.GetStockMarginRate(code)

    # code 에 해당하는 주식매매 거래단위 주식수를 반환한다.
    def get_stock_meme_min(self, code):
        return self.obj.GetStockMemeMin(code)

    # code 에 해당하는 증권전산업종코드를 반환한다.
    def get_stock_industry_code(self, code):
        return self.obj.GetStockIndustryCode(code)

    # code 에 해당하는 소속부를 반환한다.
    CPC_MARKET_NULL = 0  # 구분없음
    CPC_MARKET_KOSPI = 1  # 거래소
    CPC_MARKET_KOSDAQ = 2  # 코스닥
    CPC_MARKET_FREEBOARD = 3  # K-OTC
    CPC_MARKET_KRX = 4  # KRX
    CPC_MARKET_KONEX = 5  # KONEX

    def get_stock_market_kind(self, code):
        return self.obj.GetStockMarketKind(code)

    # code 에 해당하는 감리구분을 반환한다.
    CPC_CONTROL_NONE = 0  # 정상
    CPC_CONTROL_ATTENTION = 1  # 주의
    CPC_CONTROL_WARNING = 2  # 경고
    CPC_CONTROL_DANGER_NOTICE = 3  # 위험예고
    CPC_CONTROL_DANGER = 4  # 위험

    def get_stock_control_kind(self, code):
        return self.obj.GetStockControlKind(code)

    # code 에 해당하는 관리구분을 반환한다.
    CPC_SUPERVISION_NONE = 0
    CPC_SUPERVISION_NORMAL = 1

    def get_stock_supervision_kind(self, code):
        return self.obj.GetStockSupervisionKind(code)

    # code 에 해당하는 주식상태를 반환한다
    CPC_STOCK_STATUS_NORMAL = 0
    CPC_STOCK_STATUS_STOP = 1
    CPC_STOCK_STATUS_BREAK = 2

    def get_stock_status_kind(self, code):
        return self.obj.GetStockStatusKind(code)

    # code 에 해당하는 자본금 규모구분을 반환한다.
    CPC_CAPITAL_NULL = 0
    CPC_CAPITAL_LARGE = 1
    CPC_CAPITAL_MIDDLE = 2
    CPC_CAPITAL_SMALL = 3

    def get_stock_capital(self, code):
        return self.obj.GetStockCapital(code)

    # code 에 해당하는 결산기를 반환한다.
    def get_stock_fiscal_month(self, code):
        return self.obj.GetStockFiscalMonth(code)

    # code 에 해당하는 그룹(계열사)코드를 반환한다.
    def get_stock_group_code(self, code):
        return self.obj.GetStockGroupCode(code)

    # code 에 해당하는 KOSPI200 종목여부를 반환한다.
    CPC_KOSPI200_NONE = 0  # 미채용
    CPC_KOSPI200_CONSTRUCTIONS_MACHINERY = 1  # 건설기계
    CPC_KOSPI200_SHIPBUILDING_TRANSPORTATION = 2  # 조선운송
    CPC_KOSPI200_STEELS_METERIALS = 3  # 철강소재
    CPC_KOSPI200_ENERGY_CHEMICALS = 4  # 에너지화학
    CPC_KOSPI200_IT = 5  # 정보통신
    CPC_KOSPI200_FINANCE = 6  # 금융
    CPC_KOSPI200_CUSTOMER_STAPLES = 7  # 필수소비재
    CPC_KOSPI200_CUSTOMER_DISCRETIONARY = 8  # 자유소비재

    def get_stock_kospi200_kind(self, code):
        return self.obj.GetStockKospi200Kind(code)

    # code 에 해당하는 부구분코드를 반환한다.
    CPC_KSE_SECTION_KIND_NULL = 0  # 구분없음
    CPC_KSE_SECTION_KIND_ST = 1  # 주권
    CPC_KSE_SECTION_KIND_MF = 2  # 투자회사
    CPC_KSE_SECTION_KIND_RT = 3  # 부동산투자회사
    CPC_KSE_SECTION_KIND_SC = 4  # 선박투자회사
    CPC_KSE_SECTION_KIND_IF = 5  # 사회간접자본투융자회사
    CPC_KSE_SECTION_KIND_DR = 6  # 주식예탁증서
    CPC_KSE_SECTION_KIND_SW = 7  # 신수인수권증권
    CPC_KSE_SECTION_KIND_SR = 8  # 신주인수권증서
    CPC_KSE_SECTION_KIND_ELW = 9  # 주식워런트증권
    CPC_KSE_SECTION_KIND_ETF = 10  # 상장지수펀드(ETF)
    CPC_KSE_SECTION_KIND_BC = 11  # 수익증권
    CPC_KSE_SECTION_KIND_FETF = 12  # 해외ETF
    CPC_KSE_SECTION_KIND_FOREIGN = 13  # 외국주권
    CPC_KSE_SECTION_KIND_FU = 14  # 선물
    CPC_KSE_SECTION_KIND_OP = 15  # 옵션

    def get_stock_section_kind(self, code):
        return self.obj.GetStockSectionKind(code)

    # code 에 해당하는 락구분코드를 반환한다.
    CPC_LAC_NORMAL = 0  # 구분없음
    CPC_LAC_EX_RIGHTS = 1  # 권리락
    CPC_LAC_EX_DIVIDEND = 2  # 배당락
    CPC_LAC_EX_DISTRI_DIVIDEND = 3  # 분배락
    CPC_LAC_EX_RIGHTS_DIVIDEND = 4  # 권배락
    CPC_LAC_INTERIM_DIVIDEND = 5  # 중간배당락
    CPC_LAC_EX_RIGHTS_INTERIM_DIVIDEND = 6  # 권리중간배당락
    CPC_LAC_ETC = 99  # 기타

    def get_stock_lac_kind(self, code):
        return self.obj.GetStockLacKind(code)

    # code 에 해당하는 상장일을 반환한다.
    def get_stock_listed_date(self, code):
        return self.obj.GetStockListedDate(code)

    # code 에 해당하는 상한가를 반환한다
    def get_stock_max_price(self, code):
        return self.obj.GetStockMaxPrice(code)

    # code 에 해당하는 하한가를 반환한다.
    def get_stock_min_price(self, code):
        return self.obj.GetStockMinPrice(code)

    # code 에 해당하는 액면가를 반환한다.
    def get_stock_par_price(self, code):
        return self.obj.GetStockParPrice(code)

    # code 에 해당하는 권리락 등으로 인한 기준가를 반환한다.
    def get_stock_std_price(self, code):
        return self.obj.GetStockStdPrice(code)

    # code 에 해당하는 전일시가를 반환한다.
    def get_stock_yd_open_price(self, code):
        return self.obj.GetStockYdOpenPrice(code)

    # code 에 해당하는 전일고가를 반환한다.
    def get_stock_yd_high_price(self, code):
        return self.obj.GetStockYdHighPrice(code)

    # code 에 해당하는 전일저가를 반환한다.
    def get_stock_yd_low_price(self, code):
        return self.obj.GetStockYdLowPrice(code)

    # code 에 해당하는 전일종가를 반환한다.
    def get_stock_yd_close_price(self, code):
        return self.obj.GetStockYdClosePrice(code)

    # code 에 해당하는 신용가능종목여부를 반환한다.
    def is_stock_credit_enable(self, code):
        return self.obj.IsStockCreditEnable(code)

    # code 에 해당하는 액면정보코드를 반환한다.
    CPC_PARPRICE_CHANGE_NONE = 0  # 해당없음
    CPC_PARPRICE_CHANGE_DIVIDE = 1  # 액면분할
    CPC_PARPRICE_CHANGE_MERGE = 2  # 액면병합
    CPC_PARPRICE_CHANGE_ETC = 99  # 기타

    def get_stock_par_price_change_type(self, code):
        return self.obj.GetStockParPriceChageType(code)

    # code 에 해당하는 SPAC 종목여부를 반환한다.
    def is_spac(self, code):
        return self.obj.IsSPAC(code)

    # Elw 기초자산코드리스트얻기 (바스켓)
    def get_stock_elw_basket_code_list(self, code):
        return self.obj.GetStockElwBasketCodeList(code)

    # Elw 기초자산비율리스트얻기 (바스켓)
    def get_stock_elw_basket_comp_list(self, code):
        return self.obj.GetStockElwBasketCompList(code)

    # 시장구분에 따른 주식종목배열을 반환하다.
    def get_stock_list_by_market(self, code):
        return self.obj.GetStockListByMarket(code)

    # 관심종목(700 ~799 ) 및업종코드(GetIndustryList 참고)에 해당하는 종목배열을 반환한다.
    def get_group_code_list(self, code):
        return self.obj.GetGroupCodeList(code)

    # 관심종목(700 ~799 ) 및업종코드에 해당하는 명칭을 반환한다.
    def get_group_name(self, code):
        return self.obj.GetGroupName(code)

    # 증권전산업종 코드리스트를 반환한다.
    def get_industry_name(self, code):
        return self.obj.GetIndustryName(code)

    # 거래원코드(회원사)의 코드리스트를 반환한다.
    def get_member_list (self):
        return self.obj.GetMemberList()

    # 거래원코드(회원사)코드에 해당하는 거래원코드명을 반환한다.
    def get_member_name(self, code):
        return self.obj.GetMemberName(code)

    # 코스닥산업별 코드리스트를 반환한다.
    def get_kosdaq_industry1_list(self):
        return self.obj.GetKosdaqIndustry1List()

    # 코스닥지수업종 코드리스트를 반환한다.
    def get_kosdaq_industry2_list(self):
        return self.obj.GetKosdaqIndustry2List()

    # 장시작시각얻기 (ex 9시일 경우 리턴값 900)
    def get_market_start_time(self):
        return self.obj.GetMarketStartTime()

    # 장마감시각얻기 (ex: 오후 3시30분 일경우 리턴값 1530,수능일 1630)
    def get_market_end_time(self):
        return self.obj.GetMarketEndTime()


