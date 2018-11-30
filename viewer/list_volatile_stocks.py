import os
import sys
import time
import win32com.client


# 변동성이 큰 순서대로 종목 나열하기
# 최고가격 - 최저가격) / 전일종가 * 100 이 큰 순서대로 종목코드를 나열
HOW_MANY_SHOW = 30



inCpCybos   = win32com.client.Dispatch('CpUtil.CpCybos')
inCpCodeMgr = win32com.client.Dispatch('CpUtil.CpCodeMgr')
inMarketEye = win32com.client.Dispatch('CpSysDib.MarketEye')

codes = inCpCodeMgr.GetStockListByMarket(1) + inCpCodeMgr.GetStockListByMarket(2)
inMarketEye.SetInputValue(0, (0, 6, 7, 23))  #0-stock code, 6-highest price, 7-lowest price, 23-closing price.

start, step = 0, 200
volatile = [ ]
while True:
    #MarketEye allows maximum 200 stock codes per request.
    #make small chunk of tuple including 200 stock codes.
    target_code = codes[start:start+step]
    if not target_code:
        break
    inMarketEye.SetInputValue(1, target_code)

    if not inCpCybos.GetLimitRemainCount(1):
        time.sleep((inCpCybos.LimitRequestRemainTime + 10) / 1000)
    inMarketEye.BlockRequest2(1)

    for i in range(inMarketEye.GetHeaderValue(2)):
        try:
            vol = (inMarketEye.GetDataValue(1, i) - inMarketEye.GetDataValue(2, i)) / inMarketEye.GetDataValue(3, i) * 100
            volatile.append(round(vol, 2))
        except (TypeError, ZeroDivisionError):
            volatile.append(0)
            print('can\'t get a volatility of {}. set to 0.'.format(inMarketEye.GetDataValue(0, i)))

    start = start + step

volatility = dict(zip(codes, volatile))
sorted_volatility = tuple((code, volatility[code])  for code in sorted(volatility, key=volatility.get, reverse=True))
[print('{} {}%'.format(*sorted_volatility[i]))  for i in range(HOW_MANY_SHOW)]
