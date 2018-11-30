from cybos import cp_data
from cybos import cp_util
import pandas as pd
import datetime
import time
import sqlalchemy as db


# 주식 데이타 수집 클래스
class DataService:
    def __init__(self):
        self.CpData = cp_data.CpData()
        self.CpCodeMgr = cp_util.CpCodeMgr()
        self.CpCybos = cp_util.CpCybos()
        today = datetime.datetime.now()
        self.today_stamp = today.year * 10000 + today.month * 100 + today.day
        with open('../.ignores/database.txt') as f:
            items = list(f.readlines())
            self.host = items[1].strip()
            self.pw = items[2].strip()
        connect_string = 'mysql+pymysql://{}:{}@{}:{}/{}?charset=utf8mb4'.format('ohjongsung', self.pw, self.host, '3306', 'STOCK')
        self.engine = db.create_engine(connect_string)
        self.metadata = db.MetaData()
        self.daily_charts = db.Table('daily_charts', self.metadata, autoload=True, autoload_with=self.engine)
        print(self.daily_charts.columns.keys())
        print(repr(self.metadata.tables['daily_charts']))

    def get_stock_data_dwm(self, code, last_date):
        from_date = open_date = int(self.CpCodeMgr.get_stock_listed_date(code))

        if last_date is not None and last_date > open_date:
            from_date = last_date + 1
        else:
            pass

        self.CpData.set_input_value(0, code)
        self.CpData.set_input_value(1, ord('1'))
        self.CpData.set_input_value(2, self.today_stamp)
        self.CpData.set_input_value(3, from_date)
        self.CpData.set_input_value(5, [0, 2, 3, 4, 5, 8, 9])
        self.CpData.set_input_value(6, ord('D'))
        self.CpData.set_input_value(9, 1)

        self.CpData.block_request()
        cnt = self.CpData.get_header_value(3)

        data = []
        for i in range(cnt):
            temp_date={}
            temp_date['code'] = code
            temp_date['date'] = self.CpData.get_data_value(0, i)
            temp_date['open'] = self.CpData.get_data_value(1, i)
            temp_date['high'] = self.CpData.get_data_value(2, i)
            temp_date['low'] = self.CpData.get_data_value(3, i)
            temp_date['close'] = self.CpData.get_data_value(4, i)
            temp_date['trade_volume'] = self.CpData.get_data_value(5, i)
            temp_date['trade_amount'] = self.CpData.get_data_value(6, i)
            data.append(temp_date)

        stock_data = pd.DataFrame(data, columns=['code', 'date', 'open', 'high', 'low', 'close', 'trade_volume', 'trade_amount'])
        return stock_data

    # 대기시간 체크(TR 제한이 15초에 30개이다)
    def check_remain_time(self):
        remain_time = self.CpCybos.get_limit_request_remain_time()
        remain_count = self.CpCybos.get_limit_remain_count(1)  # 시세 제한

        if remain_count <= 0:
            time_start = time.time()
            while remain_time <= 0:
                time.sleep(remain_time+1000/1000)
                remain_count = self.CpCybos.get_limit_remain_count(1)  # 시세 제한
                remain_time = self.CpCybos.get_limit_request_remain_time
            elapsed_time = time.time() - time_start
            print("시간 지연: %.2f" %elapsed_time, "시간:", remain_time)

    def collect_data_dwm(self):
        kospi = self.CpCodeMgr.get_stock_list_by_market(1)
        kosdaq = self.CpCodeMgr.get_stock_list_by_market(22)
        code_list = list(kospi + kosdaq)
        conn = self.engine.connect()
        for i in range(len(code_list)):
            last_date = self.get_last_collect_day(code_list[i], conn)
            stock_data = self.get_stock_data_dwm(code_list[i], last_date)
            try:
                stock_data.to_sql(con=self.engine, name='daily_charts', if_exists='append', index=False)
            except Exception as e:
                print('code %s not saved. reason : %s' % (code_list[i], e))
            finally:
                print("[%s / %s] %s 수집 완료" % (i+1, len(code_list), code_list[i]))
                self.check_remain_time()

        conn.close()

    def get_last_collect_day(self, code, conn):
        query = db.select([db.func.max(self.daily_charts.columns.date)]).where(db.and_(self.daily_charts.columns.code == code))
        result = conn.execute(query).scalar()
        return result

