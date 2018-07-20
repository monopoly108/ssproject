# -*- coding:utf-8 -*-
"""
Xing API (info) request module

request function: 단일데이터 조회
request2 function: 반복데이터 조회
"""

import time

import pandas as pd
import pythoncom
import win32com.client


class RequestSessionEventHandler(object):
    """
    Class for request sessioon verification
    """
    query_state = 0

    @staticmethod
    def OnReceiveData(code):
        """
        The name of the method should be "OnReceiveData" according to
        Xing API reference.

        Parameters
        ----------
        :param code:

        Returns
        _______
        :return: check query success / failure
        """
        print("Query Success!")
        RequestSessionEventHandler.query_state = 1
        

class RequestXing(object):
    """
    Class for Xing API (info) request

    Parameters
    ----------
    _QUERY_NM
    _IN_BLOCK
    _IN_BLOCK1
    _OUT_BLOCK1
    _OUT_BLOCK2
    _OUT_BLOCK3
    """
    _QUERY_NM = "XA_DataSet.XAQuery"
    _IN_BLOCK = "tmpnmInBlock"
    _IN_BLOCK1 = "tmpnmInBlock1"
    _OUT_BLOCK = "tmpnmOutBlock"
    _OUT_BLOCK1 = "tmpnmOutBlock1"
    _OUT_BLOCK2 = "tmpnmOutBlock2"
    _OUT_BLOCK3 = "tmpnmOutBlock3"

    def __init__(self, res_file_nm):
        """
        Full path for res file
        :param res_file_nm:
            C:\\eBEST\\xingAPI\\Res\\tmpnm.res
        """
        self.res_file_nm = res_file_nm

    def request_stk_prc(self, gicode="", res_id="t1102"):
        """
        Load name, price, trading volumne of a certain stock.

        Parameters
        ----------
        :param gicode: stock code
        :param res_id: t1102

        Returns
        -------
        :return: load stock name, current price, current cumulative volume
        """
        prc_df = pd.DataFrame(columns=["name", "prc", "volume"])

        query = win32com.client.DispatchWithEvents(self._QUERY_NM,
                                                   RequestSessionEventHandler)
        query.ResFileName = self.res_file_nm.replace("tmpnm", res_id)
        query.SetFieldData(self._IN_BLOCK.replace("tmpnm", res_id), "shcode",
                           0, gicode[1:])
        query.Request(0)

        while RequestSessionEventHandler.query_state == 0:
            pythoncom.PumpWaitingMessages()

        name = query.GetFieldData(self._OUT_BLOCK.replace("tmpnm", res_id),
                                  "hname", 0)
        price = query.GetFieldData(self._OUT_BLOCK.replace("tmpnm", res_id),
                                   "price", 0)
        volume = query.GetFieldData(self._OUT_BLOCK.replace("tmpnm", res_id),
                                    "volume", 0)

        prc_df.loc[0, "name"] = name
        prc_df.loc[0, "price"] = price
        prc_df.loc[0, "volume"] = volume
        RequestSessionEventHandler.query_state = 0

        return prc_df

    def request_stk_quote(self, gicode="", res_id="t1101"):
        """
        Load 10 price and order amount of bid and ask respectively.
        10 price and order amount of bid and ask respectively
        Parameters
        ----------
        :param gicode: stock code
        :param res_id:  t1101

        Returns
        -------
        :return: 10 price and order amount of bid and ask respectively
        """
        quote_df = pd.DataFrame(columns=["offerrem", "offerho",
                                         "bidho", "bidrem"])

        query = win32com.client.DispatchWithEvents(self._QUERY_NM,
                                                   RequestSessionEventHandler)
        query.ResFileName = self.res_file_nm.replace("tmpnm", res_id)
        query.SetFieldData(self._IN_BLOCK.replace("tmpnm", res_id), "shcode",
                           0, gicode[1:])
        query.Request(0)

        while RequestSessionEventHandler.query_state == 0:
            pythoncom.PumpWaitingMessages()

        offerrem1 = query.GetFieldData(self._OUT_BLOCK.replace("tmpnm",
                                                               res_id),
                                       "offerrem1", 0)
        offerho1 = query.GetFieldData(self._OUT_BLOCK.replace("tmpnm", res_id),
                                      "offerho1", 0)
        bidho1 = query.GetFieldData(self._OUT_BLOCK.replace("tmpnm", res_id),
                                    "bidho1", 0)
        bidrem1 = query.GetFieldData(self._OUT_BLOCK.replace("tmpnm", res_id),
                                     "bidrem1", 0)

        offerrem2 = query.GetFieldData(self._OUT_BLOCK.replace("tmpnm",
                                                               res_id),
                                       "offerrem2", 0)
        offerho2 = query.GetFieldData(self._OUT_BLOCK.replace("tmpnm", res_id),
                                      "offerho2", 0)
        bidho2 = query.GetFieldData(self._OUT_BLOCK.replace("tmpnm", res_id),
                                  "bidho2", 0)
        bidrem2 = query.GetFieldData(self._OUT_BLOCK.replace("tmpnm", res_id),
                                     "bidrem2", 0)

        offerrem3 = query.GetFieldData(self._OUT_BLOCK.replace("tmpnm", res_id),
                                       "offerrem3", 0)
        offerho3 = query.GetFieldData(self._OUT_BLOCK.replace("tmpnm", res_id),
                                      "offerho3", 0)
        bidho3 = query.GetFieldData(self._OUT_BLOCK.replace("tmpnm", res_id),
                                  "bidho3", 0)
        bidrem3 = query.GetFieldData(self._OUT_BLOCK.replace("tmpnm", res_id),
                                  "bidrem3", 0)

        offerrem4 = query.GetFieldData(self._OUT_BLOCK.replace("tmpnm", res_id),
                                       "offerrem4", 0)
        offerho4 = query.GetFieldData(self._OUT_BLOCK.replace("tmpnm", res_id),
                                      "offerho4", 0)
        bidho4 = query.GetFieldData(self._OUT_BLOCK.replace("tmpnm", res_id),
                                    "bidho4", 0)
        bidrem4 = query.GetFieldData(self._OUT_BLOCK.replace("tmpnm", res_id),
                                     "bidrem4", 0)

        offerrem5 = query.GetFieldData(self._OUT_BLOCK.replace("tmpnm", res_id),
                                  "offerrem5", 0)
        offerho5 = query.GetFieldData(self._OUT_BLOCK.replace("tmpnm", res_id),
                                  "offerho5", 0)
        bidho5 = query.GetFieldData(self._OUT_BLOCK.replace("tmpnm", res_id),
                                  "bidho5", 0)
        bidrem5 = query.GetFieldData(self._OUT_BLOCK.replace("tmpnm", res_id),
                                  "bidrem5", 0)

        offerrem6 = query.GetFieldData(self._OUT_BLOCK.replace("tmpnm", res_id),
                                  "offerrem6", 0)
        offerho6 = query.GetFieldData(self._OUT_BLOCK.replace("tmpnm", res_id),
                                  "offerho6", 0)
        bidho6 = query.GetFieldData(self._OUT_BLOCK.replace("tmpnm", res_id),
                                  "bidho6", 0)
        bidrem6 = query.GetFieldData(self._OUT_BLOCK.replace("tmpnm", res_id),
                                  "bidrem6", 0)

        offerrem7 = query.GetFieldData(self._OUT_BLOCK.replace("tmpnm", res_id),
                                  "offerrem7", 0)
        offerho7 = query.GetFieldData(self._OUT_BLOCK.replace("tmpnm", res_id),
                                  "offerho7", 0)
        bidho7 = query.GetFieldData(self._OUT_BLOCK.replace("tmpnm", res_id),
                                  "bidho7", 0)
        bidrem7 = query.GetFieldData(self._OUT_BLOCK.replace("tmpnm", res_id),
                                  "bidrem7", 0)

        offerrem8 = query.GetFieldData(self._OUT_BLOCK.replace("tmpnm", res_id),
                                  "offerrem8", 0)
        offerho8 = query.GetFieldData(self._OUT_BLOCK.replace("tmpnm", res_id),
                                  "offerho8", 0)
        bidho8 = query.GetFieldData(self._OUT_BLOCK.replace("tmpnm", res_id),
                                  "bidho8", 0)
        bidrem8 = query.GetFieldData(self._OUT_BLOCK.replace("tmpnm", res_id),
                                  "bidrem8", 0)

        offerrem9 = query.GetFieldData(self._OUT_BLOCK.replace("tmpnm", res_id),
                                  "offerrem9", 0)
        offerho9 = query.GetFieldData(self._OUT_BLOCK.replace("tmpnm", res_id),
                                  "offerho9", 0)
        bidho9 = query.GetFieldData(self._OUT_BLOCK.replace("tmpnm", res_id),
                                  "bidho9", 0)
        bidrem9 = query.GetFieldData(self._OUT_BLOCK.replace("tmpnm", res_id),
                                  "bidrem9", 0)

        offerrem10 = query.GetFieldData(self._OUT_BLOCK.replace("tmpnm", res_id),
                                  "offerrem10", 0)
        offerho10 = query.GetFieldData(self._OUT_BLOCK.replace("tmpnm", res_id),
                                  "offerho10", 0)
        bidho10 = query.GetFieldData(self._OUT_BLOCK.replace("tmpnm", res_id),
                                  "bidho10", 0)
        bidrem10 = query.GetFieldData(self._OUT_BLOCK.replace("tmpnm", res_id),
                                  "bidrem10", 0)

        # bid
        quote_df.loc[0, "bidho"] = bidho1
        quote_df.loc[0, "bidrem"] = bidrem1
        quote_df.loc[1, "bidho"] = bidho2
        quote_df.loc[1, "bidrem"] = bidrem2
        quote_df.loc[2, "bidho"] = bidho3
        quote_df.loc[2, "bidrem"] = bidrem3
        quote_df.loc[3, "bidho"] = bidho4
        quote_df.loc[3, "bidrem"] = bidrem4
        quote_df.loc[4, "bidho"] = bidho5
        quote_df.loc[4, "bidrem"] = bidrem5
        quote_df.loc[5, "bidho"] = bidho6
        quote_df.loc[5, "bidrem"] = bidrem6
        quote_df.loc[6, "bidho"] = bidho7
        quote_df.loc[6, "bidrem"] = bidrem7
        quote_df.loc[7, "bidho"] = bidho8
        quote_df.loc[7, "bidrem"] = bidrem8
        quote_df.loc[8, "bidho"] = bidho9
        quote_df.loc[8, "bidrem"] = bidrem9
        quote_df.loc[9, "bidho"] = bidho10
        quote_df.loc[9, "bidrem"] = bidrem10

        # ask
        quote_df.loc[9, "offerrem"] = offerrem1
        quote_df.loc[9, "offerho"] = offerho1
        quote_df.loc[8, "offerrem"] = offerrem2
        quote_df.loc[8, "offerho"] = offerho2
        quote_df.loc[7, "offerrem"] = offerrem3
        quote_df.loc[7, "offerho"] = offerho3
        quote_df.loc[6, "offerrem"] = offerrem4
        quote_df.loc[6, "offerho"] = offerho4
        quote_df.loc[5, "offerrem"] = offerrem5
        quote_df.loc[5, "offerho"] = offerho5
        quote_df.loc[4, "offerrem"] = offerrem6
        quote_df.loc[4, "offerho"] = offerho6
        quote_df.loc[3, "offerrem"] = offerrem7
        quote_df.loc[3, "offerho"] = offerho7
        quote_df.loc[2, "offerrem"] = offerrem8
        quote_df.loc[2, "offerho"] = offerho8
        quote_df.loc[1, "offerrem"] = offerrem9
        quote_df.loc[1, "offerho"] = offerho9
        quote_df.loc[0, "offerrem"] = offerrem10
        quote_df.loc[0, "offerho"] = offerho10

        RequestSessionEventHandler.query_state = 0
        return quote_df

    def request2_account_trade_result(self, account_num="", order_pw="",
                                      mkt_gb="00", bid_gb="0", trade_gb="0",
                                      trd_dt="", res_id="CSPAQ13700"):
        """
        Load trading results of a certain date.

        Parameters
        ----------
        :param account_num: account  number
        :param order_pw: 주문비밀번호(모의투자: 0000)
        :param mkt_gb: 시장구분
        00: 전체
        10: 거래소
        20: 코스닥
        30: 프리보드
        :param bid_gb: 거래구분
        0: 전체
        1: 매도
        2: 매수
        :param trade_gb: 체결구분
        0: 전체
        1: 체결
        3: 미체결
        :param trd_dt: 거래일자
        :param res_id: CSPAQ13700

        Returns
        -------
        :return: trading results of a certain day.
        OrdDt       주문일자
        OrdTime     주문시각
        OrdNo       주문번호
        OrgOrdNo    원주문번호
        IsuNo       종목코드
        IsuNm       종목명
        BnsTpCode   매매구분
        BnsTpNm     매매구분명
        OrdQty      주문수량
        OrdPrc      주문가격
        ExecQty     체결수량
        ExecPrc     체결가격
        ExecTrxTime 체결시간
        MrcAbleQt   정정취소가능수량
        """
        final_df = pd.DataFrame(columns=["OrdDt", "OrdTime", "OrdNo",
                                         "OrgOrdNo", "IsuNo",
                                         "IsuNm", "BnsTpCode", "BnsTpNm",
                                         "OrdQty", "OrdPrc", "ExecQty",
                                         "ExecPrc", "ExecTrxTime",
                                         "MrcAbleQty"])
        account_df = pd.DataFrame(columns=["OrdDt", "OrdTime", "OrdNo",
                                           "OrgOrdNo", "IsuNo",
                                           "IsuNm", "BnsTpCode", "BnsTpNm",
                                           "OrdQty", "OrdPrc", "ExecQty",
                                           "ExecPrc", "ExecTrxTime",
                                           "MrcAbleQty"])

        query = win32com.client.DispatchWithEvents(self._QUERY_NM,
                                                   RequestSessionEventHandler)
        query.ResFileName = self.res_file_nm.replace("tmpnm", res_id)
        query.SetFieldData(self._IN_BLOCK1.replace("tmpnm", res_id), "AcntNo",
                           0, str(account_num))
        query.SetFieldData(self._IN_BLOCK1.replace("tmpnm", res_id), "InptPwd",
                           0, str(order_pw))
        query.SetFieldData(self._IN_BLOCK1.replace("tmpnm", res_id), "OrdMktCode"
                           , 0, str(mkt_gb))
        query.SetFieldData(self._IN_BLOCK1.replace("tmpnm", res_id), "BnsTpCode",
                           0, str(bid_gb))
        query.SetFieldData(self._IN_BLOCK1.replace("tmpnm", res_id), "ExecYn",
                           0, str(trade_gb))
        query.SetFieldData(self._IN_BLOCK1.replace("tmpnm", res_id), "OrdDt",
                           0, str(trd_dt))
        query.Request(0)

        while RequestSessionEventHandler.query_state == 0:
            pythoncom.PumpWaitingMessages()

        count = query.GetBlockCount(self._OUT_BLOCK3.replace("tmpnm", res_id))
        for i in range(count):
            OrdDt = query.GetFieldData(
                self._OUT_BLOCK3.replace("tmpnm", res_id), "OrdDt", i)
            account_df.loc[i, "OrdDt"] = OrdDt
            OrdTime = query.GetFieldData(
                self._OUT_BLOCK3.replace("tmpnm", res_id), "OrdTime", i)
            account_df.loc[i, "OrdTime"] = OrdTime
            OrdNo = query.GetFieldData(
                self._OUT_BLOCK3.replace("tmpnm", res_id), "OrdNo", i)
            account_df.loc[i, "OrdNo"] = OrdNo
            OrgOrdNo = query.GetFieldData(
                self._OUT_BLOCK3.replace("tmpnm", res_id), "OrgOrdNo", i)
            account_df.loc[i, "OrgOrdNo"] = OrgOrdNo
            IsuNo = query.GetFieldData(
                self._OUT_BLOCK3.replace("tmpnm", res_id), "IsuNo", i)
            account_df.loc[i, "IsuNo"] = IsuNo
            IsuNm = query.GetFieldData(
                self._OUT_BLOCK3.replace("tmpnm", res_id), "IsuNm", i)
            account_df.loc[i, "IsuNm"] = IsuNm
            BnsTpCode = query.GetFieldData(
                self._OUT_BLOCK3.replace("tmpnm", res_id), "BnsTpCode", i)
            account_df.loc[i, "BnsTpCode"] = BnsTpCode
            BnsTpNm = query.GetFieldData(
                self._OUT_BLOCK3.replace("tmpnm", res_id), "BnsTpNm", i)
            account_df.loc[i, "BnsTpNm"] = BnsTpNm
            OrdQty = query.GetFieldData(
                self._OUT_BLOCK3.replace("tmpnm", res_id), "OrdQty", i)
            account_df.loc[i, "OrdQty"] = OrdQty
            OrdPrc = query.GetFieldData(
                self._OUT_BLOCK3.replace("tmpnm", res_id), "OrdPrc", i)
            account_df.loc[i, "OrdPrc"] = OrdPrc
            ExecQty = query.GetFieldData(
                self._OUT_BLOCK3.replace("tmpnm", res_id), "ExecQty", i)
            account_df.loc[i, "ExecQty"] = ExecQty
            ExecPrc = query.GetFieldData(
                self._OUT_BLOCK3.replace("tmpnm", res_id), "ExecPrc", i)
            account_df.loc[i, "ExecPrc"] = ExecPrc
            ExecTrxTime = query.GetFieldData(
                self._OUT_BLOCK3.replace("tmpnm", res_id), "ExecTrxTime", i)
            account_df.loc[i, "ExecTrxTime"] = ExecTrxTime
            MrcAbleQty = query.GetFieldData(
                self._OUT_BLOCK3.replace("tmpnm", res_id), "MrcAbleQty", i)
            account_df.loc[i, "MrcAbleQty"] = MrcAbleQty
        final_df = final_df.append(account_df, ignore_index=True)

        # request when the # of data is above 10
        while query.IsNext is True:
            RequestSessionEventHandler.query_state = 0
            query.Request(1)

            while RequestSessionEventHandler.query_state == 0:
                pythoncom.PumpWaitingMessages()

            for i in range(count):
                OrdDt = query.GetFieldData(
                    self._OUT_BLOCK3.replace("tmpnm", res_id), "OrdDt", i)
                account_df.loc[i, "OrdDt"] = OrdDt  # 주문일자
                OrdTime = query.GetFieldData(
                    self._OUT_BLOCK3.replace("tmpnm", res_id), "OrdTime", i)
                account_df.loc[i, "OrdTime"] = OrdTime  # 주문시각
                OrdNo = query.GetFieldData(
                    self._OUT_BLOCK3.replace("tmpnm", res_id), "OrdNo", i)
                account_df.loc[i, "OrdNo"] = OrdNo  # 주문번호
                OrgOrdNo = query.GetFieldData(
                    self._OUT_BLOCK3.replace("tmpnm", res_id), "OrgOrdNo", i)
                account_df.loc[i, "OrgOrdNo"] = OrgOrdNo  # 원주문번호
                IsuNo = query.GetFieldData(
                    self._OUT_BLOCK3.replace("tmpnm", res_id), "IsuNo", i)
                account_df.loc[i, "IsuNo"] = IsuNo  # 종목코드
                IsuNm = query.GetFieldData(
                    self._OUT_BLOCK3.replace("tmpnm", res_id), "IsuNm", i)
                account_df.loc[i, "IsuNm"] = IsuNm  # 종목명
                BnsTpCode = query.GetFieldData(
                    self._OUT_BLOCK3.replace("tmpnm", res_id), "BnsTpCode", i)
                account_df.loc[i, "BnsTpCode"] = BnsTpCode   # 매매구분
                BnsTpNm = query.GetFieldData(
                    self._OUT_BLOCK3.replace("tmpnm", res_id), "BnsTpNm", i)
                account_df.loc[i, "BnsTpNm"] = BnsTpNm  # 매매구분명
                OrdQty = query.GetFieldData(
                    self._OUT_BLOCK3.replace("tmpnm", res_id), "OrdQty", i)
                account_df.loc[i, "OrdQty"] = OrdQty  # 주문수량
                OrdPrc = query.GetFieldData(
                    self._OUT_BLOCK3.replace("tmpnm", res_id), "OrdPrc", i)
                account_df.loc[i, "OrdPrc"] = OrdPrc  # 주문가격
                ExecQty = query.GetFieldData(
                    self._OUT_BLOCK3.replace("tmpnm", res_id), "ExecQty", i)
                account_df.loc[i, "ExecQty"] = ExecQty  # 체결수량
                ExecPrc = query.GetFieldData(
                    self._OUT_BLOCK3.replace("tmpnm", res_id), "ExecPrc", i)
                account_df.loc[i, "ExecPrc"] = ExecPrc  # 체결가격
                ExecTrxTime = query.GetFieldData(
                    self._OUT_BLOCK3.replace("tmpnm", res_id), "ExecTrxTime", i)
                account_df.loc[i, "ExecTrxTime"] = ExecTrxTime  # 체결시간
                MrcAbleQty = query.GetFieldData(
                    self._OUT_BLOCK3.replace("tmpnm", res_id), "MrcAbleQty", i)
                account_df.loc[i, "MrcAbleQty"] = MrcAbleQty  # 정정취소가능수량
            final_df = final_df.append(account_df, ignore_index=True)

        final_df = final_df[final_df["OrdDt"] != ""]
        RequestSessionEventHandler.query_state = 0
        return final_df

    def request2_account_result(self, account_num="", order_pw="",
                              balance_gb="0", fee_gb="0", d2_gb="0",
                              prc_gb="0", res_id="CSPAQ12300"):
        """
        Load info about an account such as stock list, average buy price etc.

        Parameters
        ----------
        :param account_num: 계좌번호
        :param order_pw: 계좌비밀번호
        :param balance_gb: 잔고구분
        0: 전체
        1: 현물
        9: 선물대용
        :param fee_gb: 수수료구분
        0: 평가시 수수료 미적용
        1: 평가시 수수료 적용
        :param d2_gb: D+2 잔고보유 구분
        0: 전부조회
        1: D2잔고 0이상만 조회
        :param prc_gb: 단가구분
        0: 평균단가
        1: BEP단가
        :param res_id: CSPAQ12300

        Returns
        -------
        :return: Load info about an account
        IsuNo       종목코드
        IsuNm       종목명
        BalQty      잔고수량
        PnlRat      손익률
        AvrUprc     평균단가
        SellAbleQty 매도가능수량
        EvalPn      평가손익
        """
        final_df = pd.DataFrame(columns=["IsuNo", "IsuNm", "BalQty",
                                          "PnlRat", "AvrUprc", "SellAbleQty",
                                          "EvalPnl"])
        account_df = pd.DataFrame(columns=["IsuNo", "IsuNm", "BalQty",
                                            "PnlRat", "AvrUprc", "SellAbleQty",
                                            "EvalPnl"])

        query = win32com.client.DispatchWithEvents(self._QUERY_NM,
                                                   RequestSessionEventHandler)
        query.ResFileName = self.res_file_nm.replace("tmpnm", res_id)
        query.SetFieldData(self._IN_BLOCK1.replace("tmpnm", res_id), "AcntNo",
                           0, str(account_num))
        query.SetFieldData(self._IN_BLOCK1.replace("tmpnm", res_id), "Pwd",
                           0, str(order_pw))
        query.SetFieldData(self._IN_BLOCK1.replace("tmpnm", res_id),
                           "BalCreTp", 0, str(balance_gb))
        query.SetFieldData(self._IN_BLOCK1.replace("tmpnm", res_id),
                           "CmsnAppTpCode", 0, str(fee_gb))
        query.SetFieldData(self._IN_BLOCK1.replace("tmpnm", res_id),
                           "D2balBaseQryTp", 0, str(d2_gb))
        query.SetFieldData(self._IN_BLOCK1.replace("tmpnm", res_id),
                           "UprcTpCode", 0, str(prc_gb))
        query.Request(0)

        while RequestSessionEventHandler.query_state == 0:
            pythoncom.PumpWaitingMessages()

        count = query.GetBlockCount(self._OUT_BLOCK3.replace("tmpnm", res_id))
        for i in range(count):
            IsuNo = query.GetFieldData(
                self._OUT_BLOCK3.replace("tmpnm", res_id), "IsuNo", i)
            account_df.loc[i, "IsuNo"] = IsuNo
            IsuNm = query.GetFieldData(
                self._OUT_BLOCK3.replace("tmpnm", res_id), "IsuNm", i)
            account_df.loc[i, "IsuNm"] = IsuNm
            BalQty = query.GetFieldData(
                self._OUT_BLOCK3.replace("tmpnm", res_id), "BalQty", i)
            account_df.loc[i, "BalQty"] = BalQty
            PnlRat = query.GetFieldData(
                self._OUT_BLOCK3.replace("tmpnm", res_id), "PnlRat", i)
            account_df.loc[i, "PnlRat"] = PnlRat
            AvrUprc = query.GetFieldData(
                self._OUT_BLOCK3.replace("tmpnm", res_id), "AvrUprc", i)
            account_df.loc[i, "AvrUprc"] = AvrUprc
            SellAbleQty = query.GetFieldData(
                self._OUT_BLOCK3.replace("tmpnm", res_id), "SellAbleQty", i)
            account_df.loc[i, "SellAbleQty"] = SellAbleQty
            EvalPnl = query.GetFieldData(
                self._OUT_BLOCK3.replace("tmpnm", res_id), "EvalPnl", i)
            account_df.loc[i, "EvalPnl"] = EvalPnl
        final_df = final_df.append(account_df, ignore_index=True)

        # request when the # of data is above 10
        while query.IsNext is True:
            RequestSessionEventHandler.query_state = 0
            query.Request(1)

            while RequestSessionEventHandler.query_state == 0:
                pythoncom.PumpWaitingMessages()

            for i in range(count):
                IsuNo = query.GetFieldData(
                    self._OUT_BLOCK3.replace("tmpnm", res_id), "IsuNo", i)
                account_df.loc[i, "IsuNo"] = IsuNo
                IsuNm = query.GetFieldData(
                    self._OUT_BLOCK3.replace("tmpnm", res_id), "IsuNm", i)
                account_df.loc[i, "IsuNm"] = IsuNm
                BalQty = query.GetFieldData(
                    self._OUT_BLOCK3.replace("tmpnm", res_id), "BalQty", i)
                account_df.loc[i, "BalQty"] = BalQty
                PnlRat = query.GetFieldData(
                    self._OUT_BLOCK3.replace("tmpnm", res_id), "PnlRat", i)
                account_df.loc[i, "PnlRat"] = PnlRat
                AvrUprc = query.GetFieldData(
                    self._OUT_BLOCK3.replace("tmpnm", res_id), "AvrUprc", i)
                account_df.loc[i, "AvrUprc"] = AvrUprc
                SellAbleQty = query.GetFieldData(
                    self._OUT_BLOCK3.replace("tmpnm", res_id), "SellAbleQty", i)
                account_df.loc[i, "SellAbleQty"] = SellAbleQty
                EvalPnl = query.GetFieldData(
                    self._OUT_BLOCK3.replace("tmpnm", res_id), "EvalPnl", i)
                account_df.loc[i, "EvalPnl"] = EvalPnl
            final_df = final_df.append(account_df, ignore_index=True)

        final_df = final_df[final_df["IsuNo"] != ""]
        RequestSessionEventHandler.query_state = 0
        return final_df

    # ************************** newly added request **************************

    def request2_margin(self, mkt_gb="", margin_rate="", credit_gb="",
                        gicode="",  res_id="t1411"):
        """
        Load margin rate of stocks

        Parameters
        ----------
        :param mkt_gb:
        0: All
        1:KOSPI
        2:KOSDAQ
        :param margin_rate:
        1) no credit
        100
        20
        30
        40
        50
        2) with credit
        100
        45
        50
        :param credit_gb:
        1: no credit
        2: with credit
        :param gicode: stock code
        :param res_id: t1411

        Returns
        -------
        :return:
        1) gicode not null: outblock
            jkrate
            sjkrate
        2) gicode null: outblock1
        """
        if str(mkt_gb) not in ['0', '1', '2']:
            print("mkt_gb should be in ['0', '1', '2']")
            return None
        elif str(credit_gb) not in ['1', '2']:
            print("credit_gb should be in ['0', '1']")
            return None
        elif (str(credit_gb) == '1') and (str(margin_rate) not in ['100', '20',
                                                                   '30', '40',
                                                                   '50']):
            print("margin_rate should be in ['100', '20', '30', '40', '50']"
                  "for credit_gb '1'")
            return None
        elif (str(credit_gb) == '2') and (str(margin_rate) not in ['100', '45',
                                                                   '50']):
            print("margin_rate should be in ['100', '45', '50']"
                  "for credit_gb '2'")
            return None
        else:
            if gicode != "":
                result_df = pd.DataFrame(columns=["gicode", "jkrate",
                                                  "sjkrate"])

                query = win32com.client.DispatchWithEvents(
                    self._QUERY_NM, RequestSessionEventHandler)
                query.ResFileName = self.res_file_nm.replace("tmpnm", res_id)
                query.SetFieldData(self._IN_BLOCK.replace("tmpnm", res_id),
                                   "shcode", 0, str(gicode[1:]))
                query.Request(0)

                while RequestSessionEventHandler.query_state == 0:
                    pythoncom.PumpWaitingMessages()

                jkrate = query.GetFieldData(
                    self._OUT_BLOCK.replace("tmpnm", res_id), "jkrate", 0)
                sjkrate = query.GetFieldData(
                    self._OUT_BLOCK.replace("tmpnm", res_id), "sjkrate", 0)

                result_df.loc[0, "gicode"] = gicode
                result_df.loc[0, "jkrate"] = jkrate
                result_df.loc[0, "sjkrate"] = sjkrate
                RequestSessionEventHandler.query_state = 0

                return result_df

            else:
                result_df = pd.DataFrame(columns=["gicode", "stk_nm",
                                                  "jkrate", "sjkrate"])
                temp_df = pd.DataFrame(columns=["gicode", "stk_nm",
                                                "jkrate", "sjkrate"])

                query = win32com.client.DispatchWithEvents(
                    self._QUERY_NM, RequestSessionEventHandler)
                query.ResFileName = self.res_file_nm.replace("tmpnm", res_id)
                query.SetFieldData(self._IN_BLOCK.replace("tmpnm", res_id),
                                   "gubun", 0, str(mkt_gb))
                query.SetFieldData(self._IN_BLOCK.replace("tmpnm", res_id),
                                   "jkrate", 0, str(margin_rate) + "%")
                query.SetFieldData(self._IN_BLOCK.replace("tmpnm", res_id),
                                   "jongchk", 0, str(credit_gb))
                query.Request(0)

                while RequestSessionEventHandler.query_state == 0:
                    pythoncom.PumpWaitingMessages()

                count = query.GetBlockCount(
                    self._OUT_BLOCK1.replace("tmpnm", res_id))
                # idx value for next request
                init_idx = int(query.GetFieldData(
                    self._OUT_BLOCK.replace("tmpnm", res_id), "idx", 0))

                for i in range(count):
                    shcode = query.GetFieldData(
                        self._OUT_BLOCK1.replace("tmpnm", res_id), "shcode", i)
                    temp_df.loc[i, "gicode"] = shcode
                    hname = query.GetFieldData(
                        self._OUT_BLOCK1.replace("tmpnm", res_id), "hname", i)
                    temp_df.loc[i, "stk_nm"] = hname
                    jkrate = query.GetFieldData(
                        self._OUT_BLOCK1.replace("tmpnm", res_id), "jkrate", i)
                    temp_df.loc[i, "jkrate"] = jkrate
                    sjkrate = query.GetFieldData(
                        self._OUT_BLOCK1.replace("tmpnm", res_id), "sjkrate",
                        i)
                    temp_df.loc[i, "sjkrate"] = sjkrate
                result_df = result_df.append(temp_df, ignore_index=True)

                # request when the # of data is above 40
                idx = init_idx
                req_cnt = 0
                while query.IsNext is True:
                    RequestSessionEventHandler.query_state = 0
                    query.SetFieldData(self._IN_BLOCK.replace("tmpnm", res_id),
                                       "idx", 0, idx)
                    query.Request(1)
                    req_cnt += 1

                    while RequestSessionEventHandler.query_state == 0:
                        pythoncom.PumpWaitingMessages()

                    # idx value for next request
                    idx2 = int(query.GetFieldData(
                        self._OUT_BLOCK.replace("tmpnm", res_id), "idx", 0))
                    idx += (idx2 - idx)

                    for i in range(count):
                        shcode = query.GetFieldData(
                            self._OUT_BLOCK1.replace("tmpnm", res_id),
                            "shcode", i)
                        temp_df.loc[i, "gicode"] = shcode
                        hname = query.GetFieldData(
                            self._OUT_BLOCK1.replace("tmpnm", res_id),
                            "hname", i)
                        temp_df.loc[i, "stk_nm"] = hname
                        jkrate = query.GetFieldData(
                            self._OUT_BLOCK1.replace("tmpnm", res_id),
                            "jkrate", i)
                        temp_df.loc[i, "jkrate"] = jkrate
                        sjkrate = query.GetFieldData(
                            self._OUT_BLOCK1.replace("tmpnm", res_id),
                            "sjkrate",
                            i)
                        temp_df.loc[i, "sjkrate"] = sjkrate
                    if req_cnt % 4 == 0:
                        time.sleep(1)
                    result_df = result_df.append(temp_df, ignore_index=True)

                result_df.drop_duplicates(["gicode"], inplace=True)
                result_df["gicode"] = result_df["gicode"].apply(lambda x:
                                                                'A' + str(x))
                result_df.reset_index(drop=True, inplace=True)
                result_df.drop(result_df.index[[len(result_df) - 1]],
                               inplace=True)
                RequestSessionEventHandler.query_state = 0
                return result_df

    def request2_top_trd_amt(self, mkt_gb, dt_gb, jc_num, req_num='1',
                             reverse='0', res_id='t1463'):
        """
        Load top trading amount stocks

        Parameters
        ----------
        :param mkt_gb:
        0: All
        1: KOSPI
        2: KOSDAQ
        :param dt_gb:
        0: today
        1: pday
        :param jc_num: param type = list / stock type not to request
        대상제외값
        (0x00000080)관리종목  => 000000000128
        (0x00000100)시장경보  => 000000000256
        (0x00000200)거래정지  => 000000000512
        (0x00004000)우선주  => 000000016384
        (0x00200000)증거금50  => 000008388608
        (0x01000000)정리매매  => 000016777216
        (0x04000000)투자유의  => 000067108864
        (0x80000000)불성실공시  => -02147483648
        두개 이상 제외시 해당 값을 합산.
        ex)관리종목 + 시장경보 = 000000000128 + 000000000256 = 000000000384
        :param req_num: # of times to request
        ex)
        1: top20
        2: top40
        :param reverse:
        0: descending
        1: ascending
        :param res_id: t1463

        Returns
        -------
        :return:
        gicode
        stk_nm
        volume cumulative trading volume
        value cumulative trading amount
        """
        if type(jc_num) != list:
            print("jc_num should be list type.")
            return None
        elif str(mkt_gb) not in ['0', '1', '2']:
            print("mkt_gb should be in ['0', '1', '2']")
            return None
        elif str(dt_gb) not in ['0', '1']:
            print("dt_gb should be in ['0', '1']")
            return None
        else:
            if str(req_num) == '1':
                result_df = pd.DataFrame(columns=["gicode", "stk_nm",
                                                  "volume", "value"])

                query = win32com.client.DispatchWithEvents(
                    self._QUERY_NM, RequestSessionEventHandler)
                query.ResFileName = self.res_file_nm.replace("tmpnm", res_id)
                query.SetFieldData(self._IN_BLOCK.replace("tmpnm", res_id),
                                   "gubun", 0, str(mkt_gb))
                query.SetFieldData(self._IN_BLOCK.replace("tmpnm", res_id),
                                   "jnilgubun", 0, str(dt_gb))
                if len(jc_num) == 0:
                    pass
                else:
                    sum_jc_num = sum(jc_num)
                    query.SetFieldData(self._IN_BLOCK.replace("tmpnm", res_id),
                                       "jc_num", 0, str(sum_jc_num))
                query.Request(0)

                while RequestSessionEventHandler.query_state == 0:
                    pythoncom.PumpWaitingMessages()

                count = query.GetBlockCount(
                    self._OUT_BLOCK1.replace("tmpnm", res_id))
                for i in range(count):
                    shcode = query.GetFieldData(
                        self._OUT_BLOCK1.replace("tmpnm", res_id), "shcode", i)
                    result_df.loc[i, "gicode"] = shcode
                    hname = query.GetFieldData(
                        self._OUT_BLOCK1.replace("tmpnm", res_id), "hname", i)
                    result_df.loc[i, "stk_nm"] = hname
                    volume = query.GetFieldData(
                        self._OUT_BLOCK1.replace("tmpnm", res_id), "volume", i)
                    result_df.loc[i, "volume"] = int(volume)
                    value = query.GetFieldData(
                        self._OUT_BLOCK1.replace("tmpnm", res_id), "value", i)
                    result_df.loc[i, "value"] = int(value)

                if str(reverse) == '0':
                    result_df["gicode"] = result_df["gicode"].apply(lambda x:
                                                                'A' + str(x))
                    result_df.sort(['value'], ascending=[0], axis=0,
                                   inplace=True)
                    RequestSessionEventHandler.query_state = 0
                    return result_df

                elif str(reverse) == '1':
                    result_df["gicode"] = result_df["gicode"].apply(lambda x:
                                                                'A' + str(x))
                    result_df.sort(['value'], ascending=[1], axis=0,
                                   inplace=True)
                    RequestSessionEventHandler.query_state = 0
                    return result_df

            else:
                result_df = pd.DataFrame(columns=["gicode", "stk_nm",
                                                  "volume", "value"])
                temp_df = pd.DataFrame(columns=["gicode", "stk_nm",
                                                "volume", "value"])
                req_cnt = 0

                query = win32com.client.DispatchWithEvents(
                    self._QUERY_NM, RequestSessionEventHandler)
                query.ResFileName = self.res_file_nm.replace("tmpnm", res_id)
                query.SetFieldData(self._IN_BLOCK.replace("tmpnm", res_id),
                                   "gubun", 0, str(mkt_gb))
                query.SetFieldData(self._IN_BLOCK.replace("tmpnm", res_id),
                                   "jnilgubun", 0, str(dt_gb))
                if len(jc_num) == 0:
                    pass
                else:
                    sum_jc_num = sum(jc_num)
                    query.SetFieldData(self._IN_BLOCK.replace("tmpnm", res_id),
                                       "jc_num", 0, str(sum_jc_num))
                query.Request(0)

                while RequestSessionEventHandler.query_state == 0:
                    pythoncom.PumpWaitingMessages()
                req_cnt += 1

                count = query.GetBlockCount(
                    self._OUT_BLOCK1.replace("tmpnm", res_id))
                # idx value for next request
                init_idx = int(query.GetFieldData(
                    self._OUT_BLOCK.replace("tmpnm", res_id), "idx", 0))
                for i in range(count):
                    shcode = query.GetFieldData(
                        self._OUT_BLOCK1.replace("tmpnm", res_id), "shcode", i)
                    temp_df.loc[i, "gicode"] = shcode
                    hname = query.GetFieldData(
                        self._OUT_BLOCK1.replace("tmpnm", res_id), "hname", i)
                    temp_df.loc[i, "stk_nm"] = hname
                    volume = query.GetFieldData(
                        self._OUT_BLOCK1.replace("tmpnm", res_id), "volume", i)
                    temp_df.loc[i, "volume"] = int(volume)
                    value = query.GetFieldData(
                        self._OUT_BLOCK1.replace("tmpnm", res_id), "value", i)
                    temp_df.loc[i, "value"] = int(value)
                result_df = result_df.append(temp_df, ignore_index=True)

                # request when the # of data is above 40
                idx = init_idx
                while query.IsNext is True:
                    if req_cnt == int(req_num):
                        break
                    else:
                        RequestSessionEventHandler.query_state = 0
                        query.SetFieldData(
                            self._IN_BLOCK.replace("tmpnm", res_id), "idx", 0,
                            idx)
                        query.Request(1)

                        while RequestSessionEventHandler.query_state == 0:
                            pythoncom.PumpWaitingMessages()

                        # idx value for next request
                        idx2 = int(query.GetFieldData(
                            self._OUT_BLOCK.replace("tmpnm", res_id), "idx",
                            0))
                        idx += (idx2 - idx)

                        for i in range(count):
                            shcode = query.GetFieldData(
                                self._OUT_BLOCK1.replace("tmpnm", res_id),
                                "shcode", i)
                            temp_df.loc[i, "gicode"] = shcode
                            hname = query.GetFieldData(
                                self._OUT_BLOCK1.replace("tmpnm", res_id),
                                "hname", i)
                            temp_df.loc[i, "stk_nm"] = hname
                            volume = query.GetFieldData(
                                self._OUT_BLOCK1.replace("tmpnm", res_id),
                                "volume", i)
                            temp_df.loc[i, "volume"] = int(volume)
                            value = query.GetFieldData(
                                self._OUT_BLOCK1.replace("tmpnm", res_id),
                                "value", i)
                            temp_df.loc[i, "value"] = int(value)
                        req_cnt += 1
                        if req_cnt % 4 == 0:
                            time.sleep(1)
                        result_df = result_df.append(temp_df,
                                                     ignore_index=True)
                if str(reverse) == '0':
                    result_df.drop_duplicates(["gicode"], inplace=True)
                    result_df["gicode"] = result_df["gicode"].apply(lambda x:
                                                                'A' + str(x))
                    result_df.sort(['value'], ascending=[0], axis=0,
                                   inplace=True)
                    RequestSessionEventHandler.query_state = 0
                    return result_df

                elif str(reverse) == '1':
                    result_df.drop_duplicates(["gicode"], inplace=True)
                    result_df["gicode"] = result_df["gicode"].apply(lambda x:
                                                                'A' + str(x))
                    result_df.sort(['value'], ascending=[1], axis=0,
                                   inplace=True)
                    RequestSessionEventHandler.query_state = 0
                    return result_df
