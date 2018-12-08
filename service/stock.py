from cybos import cp_util
from cybos import cp_stock


# 주식 정보 서비스 클래스
class StockService:
    def __init__(self):
        self.CpStockCode = cp_util.CpStockCode()
        self.StockMst = cp_stock.CpStockMst()

    def code_to_name(self, code):
        return self.CpStockCode.code_to_name(code)

    def get_current_price(self, code):
        self.StockMst.set_input_value(0, code)
        self.StockMst.block_request()

        if self.StockMst.get_communication_status() is False:
            return False

        price_info = dict()
        price_info['code'] = code
        price_info['name'] = self.StockMst.get_header_value(1)
        price_info['current'] = self.StockMst.get_header_value(11)

        # 호가 만들기
        for i in range(10):
            sell_price = 'sell%d' % (i+1)
            buy_price = 'buy%d' % (i+1)
            price_info[sell_price] = self.StockMst.get_data_value(0, i)
            price_info[buy_price] = self.StockMst.get_data_value(1, i)

        return price_info
