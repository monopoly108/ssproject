# -*- coding:utf-8 -*-
"""
Xing API order module
"""
import win32com.client
import pythoncom


class OrderSessionEventHandler(object):
    """
    Class for trade session verification
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
        print("Order Success!")
        OrderSessionEventHandler.query_state = 1


class OrderXing(object):
    """
    Class for Xing API order

    Parameters
    ----------
    _QUERY_NM
    _IN_BLOCK1
    _OUT_BLOCK1
    _OUT_BLOCK2
    _OUT_BLOCK3
    """
    _QUERY_NM = "XA_DataSet.XAQuery"
    _IN_BLOCK1 = "tmpnmInBlock1"
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

    def bid_order(self, login_gb="", account_nm="", order_pw="", gicode="",
                  order_qty="", order_prc="", bid_gb='2',
                  order_type="03", credit_gb="000", order_con='0',
                  res_id="CSPAT00600"):
        """
        Buy(bid) Order

        Parameters
        ----------
        :param login_gb: 0: 모의투자  / 1: 실전투자 (종목코드 구분시 필요)
        :param account_nm: 계좌번호
        :param order_pw: 주문비밀번호(모의투자: 0000)
        :param gicode: 종목코드 or A+종목코드(모의투자 A+종목코드)
        :param order_qty: 주문수량
        :param order_prc: 주문가격
        :param bid_gb: 1: 매도 / 2: 매수
        :param order_type: 호가유형코드
        00: 지정가
        03: 시장가
        05: 조건부지정가
        06: 최유리지정가
        07: 최우선지정가
        61: 장개시전시간외종가
        81: 시간외종가
        82: 시간외단일가
        :param credit_gb: 신용거래코드
        000:보통
        003:유통/자기융자신규
        005:유통대주신규
        007:자기대주신규
        101:유통융자상환
        103:자기융자상환
        105:유통대주상환
        107:자기대주상환
        180:예탁담보대출상환(신용)
        :param order_con: 주문조건구분
        0: 없음
        1: IOC(Immediate Or Cancel) - 주문 즉시 체결 잔량은 취소
        2: FOK() - 전량 체결이 되지 않으면 전량 취소
        :param res_id: CSPAT00600

        Returns
        -------
        :return: Bid order result
        """
        order = win32com.client.DispatchWithEvents(self._QUERY_NM,
                                                   OrderSessionEventHandler)
        order.ResFileName = self.res_file_nm.replace("tmpnm", res_id)

        order.SetFieldData(self._IN_BLOCK1.replace("tmpnm", res_id),
                           'AcntNo', 0, account_nm)
        order.SetFieldData(self._IN_BLOCK1.replace("tmpnm", res_id),
                           'InptPwd', 0, order_pw)
        if login_gb == '0':
            order.SetFieldData(self._IN_BLOCK1.replace("tmpnm", res_id),
                               'IsuNo', 0, gicode)
        elif login_gb == '1':
            order.SetFieldData(self._IN_BLOCK1.replace("tmpnm", res_id),
                               'IsuNo', 0, gicode[1:])
        else:
            print("Login gb should be 0(모의투자) / 1(실전투자)")
            raise ValueError
        order.SetFieldData(self._IN_BLOCK1.replace("tmpnm", res_id), 'OrdQty',
                           0, int(order_qty))
        if order_type == "03":
            # 시장가주문이면 가격을 지정할 필요가 없다.
            order.SetFieldData(self._IN_BLOCK1.replace("tmpnm", res_id),
                               'OrdPrc', 0, '')
        else:
            order.SetFieldData(self._IN_BLOCK1.replace("tmpnm", res_id),
                               'OrdPrc', 0, int(order_prc))
        order.SetFieldData(self._IN_BLOCK1.replace("tmpnm", res_id),
                           'BnsTpCode', 0, bid_gb)
        order.SetFieldData(self._IN_BLOCK1.replace("tmpnm", res_id),
                           'OrdprcPtnCode', 0, order_type)
        order.SetFieldData(self._IN_BLOCK1.replace("tmpnm", res_id),
                           'MgntrnCode', 0, credit_gb)
        order.SetFieldData(self._IN_BLOCK1.replace("tmpnm", res_id),
                           'OrdCndiTpCode', 0, order_con)
        order.Request(0)

        while OrderSessionEventHandler.query_state == 0:
            pythoncom.PumpWaitingMessages()

        OrdNo = order.GetFieldData(self._OUT_BLOCK2.replace("tmpnm", res_id),
                                   'OrdNo', 0)  # 주문번호
        IsuNo = order.GetFieldData(self._OUT_BLOCK1.replace("tmpnm", res_id),
                                   'IsuNo', 0)  # 종목코드
        OrdQty = order.GetFieldData(self._OUT_BLOCK1.replace("tmpnm", res_id),
                                    'OrdQty', 0)  # 종목수량
        OrdPrc = order.GetFieldData(self._OUT_BLOCK1.replace("tmpnm", res_id),
                                    'OrdPrc', 0)  # 종목수량
        OrdAmt = order.GetFieldData(self._OUT_BLOCK2.replace("tmpnm", res_id),
                                    'OrdAmt', 0)  # 주문금액
        BnsTpCode = order.GetFieldData(self._OUT_BLOCK1.replace("tmpnm",
                                                                res_id),
                                       'BnsTpCode', 0)  # 종목수량
        OrdTime = order.GetFieldData(self._OUT_BLOCK2.replace("tmpnm", res_id),
                                     'OrdTime', 0)  # 주문시각

        order_res = [OrdNo, IsuNo, OrdQty, OrdPrc, OrdAmt, BnsTpCode, OrdTime]
        OrderSessionEventHandler.query_state = 0
        return order_res

    def ask_order(self, login_gb="", account_nm="", order_pw="", gicode="",
                  order_qty="", order_prc="", bid_gb='1',
                  order_type="03", credit_gb="000", order_con='0',
                  res_id="CSPAT00600"):
        """
        Ask(sell) Order

        Parameters
        ----------

        :param login_gb: 0: 모의투자 / 1: 실전투자 (종목코드 구분시 필요)
        :param account_nm: 계좌번호
        :param order_pw: 주문비밀번호(모의투자: 0000)
        :param gicode: 종목코드 or A+종목코드(모의투자 A+종목코드)
        :param order_qty: 주문수량
        :param order_prc: 주문가격
        :param bid_gb: 1: 매도 / 2: 매수
        :param order_type: 호가유형코드
        00: 지정가
        03: 시장가
        05: 조건부지정가
        06: 최유리지정가
        07: 최우선지정가
        61: 장개시전시간외종가
        81: 시간외종가
        82: 시간외단일가
        :param credit_gb: 신용거래코드
        000:보통
        003:유통/자기융자신규
        005:유통대주신규
        007:자기대주신규
        101:유통융자상환
        103:자기융자상환
        105:유통대주상환
        107:자기대주상환
        180:예탁담보대출상환(신용)
        :param order_con: 주문조건구분
        0: 없음
        1: IOC(Immediate Or Cancel) - 주문 즉시 체결 잔량은 취소
        2: FOK() - 전량 체결이 되지 않으면 전량 취소
        :param res_id: CSPAT00600

        Returns
        -------
        :return: Ask order result
        """
        order = win32com.client.DispatchWithEvents(self._QUERY_NM,
                                                   OrderSessionEventHandler)
        order.ResFileName = self.res_file_nm.replace("tmpnm", res_id)

        order.SetFieldData(self._IN_BLOCK1.replace("tmpnm", res_id),
                           'AcntNo', 0, account_nm)
        order.SetFieldData(self._IN_BLOCK1.replace("tmpnm", res_id),
                           'InptPwd', 0, order_pw)
        if login_gb == '0':
            order.SetFieldData(self._IN_BLOCK1.replace("tmpnm", res_id),
                               'IsuNo', 0, gicode)
        elif login_gb == '1':
            order.SetFieldData(self._IN_BLOCK1.replace("tmpnm", res_id),
                               'IsuNo', 0, gicode[1:])
        else:
            print("Login gb should be 0(모의투자) / 1(실전투자)")
            raise ValueError
        order.SetFieldData(self._IN_BLOCK1.replace("tmpnm", res_id), 'OrdQty',
                           0, int(order_qty))
        if order_type == "03":
            # 시장가주문이면 가격을 지정할 필요가 없다.
            order.SetFieldData(self._IN_BLOCK1.replace("tmpnm", res_id),
                               'OrdPrc', 0, '')
        else:
            order.SetFieldData(self._IN_BLOCK1.replace("tmpnm", res_id),
                               'OrdPrc', 0, int(order_prc))
        order.SetFieldData(self._IN_BLOCK1.replace("tmpnm", res_id),
                           'BnsTpCode', 0, bid_gb)
        order.SetFieldData(self._IN_BLOCK1.replace("tmpnm", res_id),
                           'OrdprcPtnCode', 0, order_type)
        order.SetFieldData(self._IN_BLOCK1.replace("tmpnm", res_id),
                           'MgntrnCode', 0, credit_gb)
        order.SetFieldData(self._IN_BLOCK1.replace("tmpnm", res_id),
                           'OrdCndiTpCode', 0, order_con)
        order.Request(0)

        while OrderSessionEventHandler.query_state == 0:
            pythoncom.PumpWaitingMessages()

        OrdNo = order.GetFieldData(self._OUT_BLOCK2.replace("tmpnm", res_id),
                                   'OrdNo', 0)  # 주문번호
        IsuNo = order.GetFieldData(self._OUT_BLOCK1.replace("tmpnm", res_id),
                                   'IsuNo', 0)  # 종목코드
        OrdQty = order.GetFieldData(self._OUT_BLOCK1.replace("tmpnm", res_id),
                                    'OrdQty', 0)  # 종목수량
        OrdPrc = order.GetFieldData(self._OUT_BLOCK1.replace("tmpnm", res_id),
                                    'OrdPrc', 0)  # 종목수량
        OrdAmt = order.GetFieldData(self._OUT_BLOCK2.replace("tmpnm", res_id),
                                    'OrdAmt', 0)  # 주문금액
        BnsTpCode = order.GetFieldData(self._OUT_BLOCK1.replace("tmpnm",
                                                                res_id),
                                       'BnsTpCode', 0)  # 종목수량
        OrdTime = order.GetFieldData(self._OUT_BLOCK2.replace("tmpnm",
                                                              res_id),
                                     'OrdTime', 0)  # 주문시각

        order_res = [OrdNo, IsuNo, OrdQty, OrdPrc, OrdAmt, BnsTpCode, OrdTime]
        OrderSessionEventHandler.query_state = 0
        return order_res

    def cancel_order(self, login_gb="", account_nm="", order_pw="",
                     order_num="", gicode="", order_qty="",
                     res_id="CSPAT00800"):
        """
        Cancel order

        Parameters
        ----------
        :param login_gb: 0: 모의투자 / 1: 실전투자 (종목코드 구분시 필요)
        :param account_nm: 계좌번호
        :param order_pw: 계좌비밀번호
        :param order_num: 주문번호
        :param gicode: 종목코드 or 모의투자는 A+종목코드 ELW : J+종목코드
        :param order_qty: 취소수량
        :param res_id: CSPAT00800

        Returns
        -------
        :return: Cancel order result
        """
        order = win32com.client.DispatchWithEvents(self._QUERY_NM,
                                                   OrderSessionEventHandler)
        order.ResFileName = self.res_file_nm.replace("tmpnm", res_id)

        order.SetFieldData(self._IN_BLOCK1.replace("tmpnm", res_id),
                           'AcntNo', 0, account_nm)
        order.SetFieldData(self._IN_BLOCK1.replace("tmpnm", res_id),
                           'InptPwd', 0, order_pw)
        order.SetFieldData(self._IN_BLOCK1.replace("tmpnm", res_id),
                           'OrgOrdNo', 0, int(order_num))
        if login_gb == '0':
            order.SetFieldData(self._IN_BLOCK1.replace("tmpnm", res_id),
                               'IsuNo', 0, gicode)
        elif login_gb == '1':
            order.SetFieldData(self._IN_BLOCK1.replace("tmpnm", res_id),
                               'IsuNo', 0, gicode[1:])
        else:
            print("Login gb should be 0(모의투자) / 1(실전투자)")
            raise ValueError
        order.SetFieldData(self._IN_BLOCK1.replace("tmpnm", res_id),
                           'OrdQty', 0, int(order_qty))
        order.Request(0)

        while OrderSessionEventHandler.query_state == 0:
            pythoncom.PumpWaitingMessages()
        OrderSessionEventHandler.query_state = 0

        OrdNo = order.GetFieldData(self._OUT_BLOCK2.replace("tmpnm", res_id),
                                   'OrdNo', 0)  # 주문번호
        PrntOrdNo = order.GetFieldData(self._OUT_BLOCK2.replace("tmpnm",
                                                                res_id),
                                       'PrntOrdNo', 0)  # 모주문번호
        OrdTime = order.GetFieldData(self._OUT_BLOCK2.replace("tmpnm", res_id),
                                     'OrdTime', 0)  # 주문시각
        OrdPtnCode = order.GetFieldData(self._OUT_BLOCK2.replace("tmpnm",
                                                                 res_id),
                                        'OrdPtnCode', 0)  # 주문유형코드
        BnsTpCode = order.GetFieldData(self._OUT_BLOCK2.replace("tmpnm",
                                                                res_id),
                                       'BnsTpCode', 0)  # 매매구분
        IsuNo = order.GetFieldData(self._OUT_BLOCK1.replace("tmpnm", res_id),
                                   'IsuNo', 0)  # 종목코드
        IsuNm = order.GetFieldData(self._OUT_BLOCK2.replace("tmpnm", res_id),
                                   'IsuNm', 0)  # 종목명
        OrdQty = order.GetFieldData(self._OUT_BLOCK1.replace("tmpnm", res_id),
                                    'OrdQty', 0)  # 주문수량

        order_res = [OrdNo, PrntOrdNo, OrdTime, OrdPtnCode, BnsTpCode, IsuNo,
                     IsuNm, OrdQty]
        return order_res

    def revise_order(self, login_gb="", account_nm="", order_pw="",
                     order_num="", gicode="", order_qty="", order_prc="",
                     order_type="03", order_con="0", res_id="CSPAT00700"):
        """
        Revise Order

        Parameters
        ----------
        :param login_gb: 0: 모의투자 / 1: 실전투자 (종목코드 구분시 필요)
        :param account_nm: 계좌번호
        :param order_pw: 계좌비밀번호
        :param order_num: 주문번호
        :param gicode: 종목코드 or 모의투자는 A+종목코드 ELW : J+종목코드
        :param order_qty: 정정수량
        :param order_type:
        00: 지정가
        03: 시장가
        05: 조건부지정가
        06: 최유리지정가
        07: 최우선지정가
        61: 장개시전시간외종가
        81: 시간외종가
        82: 시간외단일가
        :param order_con:
        0: 없음
        1: IOC
        2: FOK
        :param res_id: CSPAT00700

        Returns
        -------
        :return: Revise order result
        """
        order = win32com.client.DispatchWithEvents(self._QUERY_NM,
                                                   OrderSessionEventHandler)
        order.ResFileName = self.res_file_nm.replace("tmpnm", res_id)

        order.SetFieldData(self._IN_BLOCK1.replace("tmpnm", res_id),
                           'AcntNo', 0, account_nm)
        order.SetFieldData(self._IN_BLOCK1.replace("tmpnm", res_id),
                           'InptPwd', 0, order_pw)
        order.SetFieldData(self._IN_BLOCK1.replace("tmpnm", res_id),
                           'OrgOrdNo', 0, int(order_num))
        if login_gb == '0':
            order.SetFieldData(self._IN_BLOCK1.replace("tmpnm", res_id),
                               'IsuNo', 0, gicode)
        elif login_gb == '1':
            order.SetFieldData(self._IN_BLOCK1.replace("tmpnm", res_id),
                               'IsuNo', 0, gicode[1:])
        else:
            print("Login gb should be 0(모의투자) / 1(실전투자)")
            raise ValueError
        order.SetFieldData(self._IN_BLOCK1.replace("tmpnm", res_id), 'OrdQty',
                           0, int(order_qty))
        if order_type == "03":
            # 시장가주문이면 가격을 지정할 필요가 없다.
            order.SetFieldData(self._IN_BLOCK1.replace("tmpnm", res_id),
                               'OrdPrc', 0, '')
        else:
            order.SetFieldData(self._IN_BLOCK1.replace("tmpnm", res_id),
                               'OrdPrc', 0, int(order_prc))
        order.SetFieldData(self._IN_BLOCK1.replace("tmpnm", res_id),
                           'OrdprcPtnCode', 0, order_type)
        order.SetFieldData(self._IN_BLOCK1.replace("tmpnm", res_id),
                           'OrdCndiTpCode', 0, order_con)
        order.Request(0)

        while OrderSessionEventHandler.query_state == 0:
            pythoncom.PumpWaitingMessages()
        OrderSessionEventHandler.query_state = 0

        OrdNo = order.GetFieldData(self._OUT_BLOCK2.replace("tmpnm", res_id),
                                   'OrdNo', 0)  # 주문번호
        PrntOrdNo = order.GetFieldData(self._OUT_BLOCK2.replace("tmpnm",
                                                                res_id),
                                       'PrntOrdNo', 0)  # 모주문번호
        OrdTime = order.GetFieldData(self._OUT_BLOCK2.replace("tmpnm", res_id),
                                     'OrdTime', 0)  # 주문시각
        OrdPtnCode = order.GetFieldData(self._OUT_BLOCK2.replace("tmpnm",
                                                                 res_id),
                                        'OrdPtnCode', 0)  # 주문유형코드
        BnsTpCode = order.GetFieldData(self._OUT_BLOCK2.replace("tmpnm",
                                                                res_id),
                                       'BnsTpCode', 0)  # 매매구분
        IsuNo = order.GetFieldData(self._OUT_BLOCK1.replace("tmpnm", res_id),
                                   'IsuNo', 0)  # 종목코드
        IsuNm = order.GetFieldData(self._OUT_BLOCK2.replace("tmpnm", res_id),
                                   'IsuNm', 0)  # 종목명
        OrdQty = order.GetFieldData(self._OUT_BLOCK1.replace("tmpnm", res_id),
                                    'OrdQty', 0)  # 주문수량
        OrdPrc = order.GetFieldData(self._OUT_BLOCK1.replace("tmpnm", res_id),
                                    'OrdPrc', 0)  # 주문가격
        OrdprcPtnCode = order.GetFieldData(self._OUT_BLOCK1.replace("tmpnm",
                                                                    res_id),
                                           'OrdprcPtnCode', 0)  # 호가유형구분
        OrdCndiTpCode = order.GetFieldData(self._OUT_BLOCK1.replace("tmpnm",
                                                                    res_id),
                                           'OrdCndiTpCode', 0)  # 주문조건구분

        order_res = [OrdNo, PrntOrdNo, OrdTime, OrdPtnCode, BnsTpCode, IsuNo,
                     IsuNm, OrdQty, OrdPrc, OrdprcPtnCode, OrdCndiTpCode]
        return order_res
