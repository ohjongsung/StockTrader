from cybos import cp_util


# 주식 정보 서비스 클래스
class StockService:
    def __init__(self):
        self.CpStockCode = cp_util.CpStockCode()

    def code_to_name(self, code):
        return self.CpStockCode.code_to_name(code)
