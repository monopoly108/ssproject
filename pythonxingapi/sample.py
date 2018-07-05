# -*- coding:utf-8 -*-

import time

import pandas as pd

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
