import pandas as pd
import datetime
from CreonPy import *
from dt_alimi import *
import time

bot.sendMessage(myId, "단타킹 종가매매 매수 시작")

today = datetime.datetime.now().date()
data_path = "C:\\CloudStation\\dt_data\\buy_list"

file = data_path + "\\buy_" + today.strftime("%y%m%d") + ".csv"
items = pd.read_csv(file)
items = items['code'].to_list()
items2 = list(set(items))  # 중복요소 제거
print(items2)

InitPlusCheck()

account = g_objCpTrade.AccountNumber[1]  # 종가매매 계좌번호
unit_price = 1000000  # 종목당 매수금액

cp_order = CpRPOrder(account)
cp_price = CpRPCurrentPrice()

msg = []
for cd in items2:
    item = cp_price.Request(cd)
    print(item)
    if item['예상플래그'] == '1':
        curPrice = item['cur']

        amount = divmod(unit_price, curPrice)[0]
        while cp_order.buy_order(cd, amount) is False:
            time.sleep(2)

        temp = "%s %s %i" % (cd, item['종목명'], amount)
        print(temp)
        msg.append(temp + '\n')

if msg:
    bot.sendMessage(myId, "".join(msg))

bot.sendMessage(myId, "단타킹 종가매매 매수 종료")
