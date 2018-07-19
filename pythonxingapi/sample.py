# -*- coding:utf-8 -*-

import time

import datetime as dt
import pandas as pd
from dateutil import relativedelta

import pythonxingapi.xinglogin as xinglogin
import pythonxingapi.xingrequest as xingrequest

# 0.basic info
resfilenm = "C:\\eBEST\\xingAPI\\Res\\tmpnm.res"  # res file dir + file_name
ebestsec_id = ""  # your id
ebestsec_pw = ""  # your password
cert_pw = ""  # no need for demo investing
login_gb = "0"  # demo investing
account_pw = "0000"  # demo investing

# 1.login
login = xinglogin.ConnectXing(xing_id=ebestsec_id, xing_pw=ebestsec_pw,
                              cert_pw=cert_pw, login_gb=login_gb)
login.login_xing()
account_num = login.get_account_num()[0]  # account number
print(account_num)

# 2.request
request = xingrequest.RequestXing(res_file_nm=resfilenm)


# trade result
trade_res = request.request2_account_trade_result(account_num=account_num,
                                                  order_pw=account_pw,
                                                  trd_dt="20180704")
print(trade_res)

time.sleep(1)
# account portfolio status
account_res = request.request2_account_result(account_num=account_num,
                                              order_pw=account_pw)
print(account_res)

# 3. trade
stk_list = pd.read_excel("sample_stklist.xlsx", "Sheet1", header=0)[
    "GICODE"].tolist()
print(stk_list)

# test(It will be modified.)


def time_to_sec(x):
    """
    Time format %H%M%S to seconds

    :param x:
    :return:
    """
    h, m, s = x.split(':')
    return int(h) * 3600 + int(m) * 60 + int(s)

"""
trading_end_gb = 0
trading_result_gb = 0

start_time = dt.datetime.now()
start_time_hms = start_time.strftime("%H%M%S")
start_time_plus_ten = start_time + relativedelta.relativedelta(seconds=601)
request_num = 0

if ("090000" <= start_time_hms < "154500") and (trading_end_gb == 0):
    # trading time
    print("Trading time.")
    '''
    Trading Logic Function
    '''
    request_num += 1

    # for xing API requests can only be sent 200 times within 10 minutes.
    if request_num >= 199:
        sleep_sec = (start_time_plus_ten - dt.datetime.now())
        sleep_sec = time_to_sec(str(sleep_sec).split('.')[0])
        time.sleep(sleep_sec)

    if start_time_hms >= "153001":
        trading_end_gb = 1

elif ("154500" <= start_time_hms < "160000") and trading_result_gb == 0:
    # report time
    print("Reporting time.")
    '''
    Reporting Function
    '''
    trading_result_gb = 1

else:
    pass
"""
