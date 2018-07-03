# -*- coding:utf-8 -*-
"""
Xing API Modules
"""

import win32com.client
import pythoncom


class LoginSessionEventHandler(object):
    """
    Class for login session status verification
    """
    login_state = 0

    def __init__(self):
        return

    @staticmethod
    def OnLogin(code, msg):
        """
        The name of this method should be <OnLogin> according to
        Xing API reference.

        Parameters
        ----------
        :param code:
        :param msg:

        Returns
        -------
        :return: check login status
        """
        if code == "0000":
            print("Login Success!")
            LoginSessionEventHandler.login_state = 1
        else:
            print("Login Failed!")


class ConnectXing(object):
    """
    Class for Xing API connection.

    Parameters
    ----------
    id: xing id
    pw: xing pw
    cert_pw: public certification password(for real server use)
    login_gb: 0:test server / 1:real server
    """
    _real_server_addr = "hts.ebestsec.co.kr"
    _demo_server_addr = "demo.ebestsec.co.kr"
    _server_port = 20001
    _server_type = 0
    _session_nm = "XA_Session.XASession"
    _instance_count = 0

    def __init__(self, xing_id, xing_pw, cert_pw, login_gb):
        self.xing_id = str(xing_id)
        self.xing_pw = str(xing_pw)
        self.cert_pw = str(cert_pw)
        self.login_gb = str(login_gb)
        self.account_num_list = []

    @staticmethod
    def _add_instance_count():
        ConnectXing._instance_count += 1

    def login_xing(self):
        """
        :return: login xing
        """
        if self.login_gb == '0':
            if self.xing_id == "":
                print("Please insert id.")
                exit()
            if self.xing_pw == "":
                print("Please insert password.")
                exit()
            if self.cert_pw == "":
                print("Please insert certification password.")
                exit()
            session = win32com.client.DispatchWithEvents(self._session_nm,
                                                    LoginSessionEventHandler)
            session.ConnectServer(self._demo_server_addr, self._server_port)
            session.Login(self.xing_id, self.xing_pw, self.cert_pw,
                          self._server_type, 0)

            # Wait until login event finishes(Wait for call back)
            while LoginSessionEventHandler.login_state == 0:
                pythoncom.PumpWaitingMessages()
            print("Test Server Connected")

            # get account number
            count = session.GetAccountListCount()
            if count == 0:
                print("Please create an account")
                exit()

            for i in range(count):
                self.account_num_list.append(session.GetAccountList(i))

        elif self.login_gb == '1':
            if self.xing_id == "":
                print("Please insert id.")
                exit()
            if self.xing_pw == "":
                print("Please insert password.")
                exit()
            if self.cert_pw == "":
                print("Please insert certification password.")
                exit()

            session = win32com.client.DispatchWithEvents(self._session_nm,
                                                    LoginSessionEventHandler)
            session.ConnectServer(self._real_server_addr, self._server_port)
            session.Login(self.xing_id, self.xing_pw, self.cert_pw,
                          self._server_type, 0)

            # Wait until login event finishes(Wait for call back)
            while LoginSessionEventHandler.login_state == 0:
                pythoncom.PumpWaitingMessages()
            print("Real Server Connected")

            # get account number
            count = session.GetAccountListCount()
            if count == 0:
                print("Please create an account")
                exit()

            for i in range(count):
                self.account_num_list.append(session.GetAccountList(i))
        else:
            print("Please Check login_gb. 0 for test server, "
                  "1 for real server")
            return

    def get_account_num(self):
        """
        Get account number

        :return: account number
        """
        return self.account_num_list

    def __del__(self):
        print("End Connection!")
        return


if __name__ == "__main__":

    pass
