# -*- coding: utf-8 -*-

import pywinauto
import pywinauto.clipboard

from easytrader import grid_strategies
from . import clienttrader


class HTClientTrader(clienttrader.BaseLoginClientTrader):
    grid_strategy = grid_strategies.Xls

    @property
    def broker_type(self):
        return "ht"

    def login(self, user, password, exe_path, comm_password=None, **kwargs):
        """
        :param user: 用户名
        :param password: 密码
        :param exe_path: 客户端路径, 类似
        :param comm_password:
        :param kwargs:
        :return:
        """
        self._editor_need_type_keys = False
        if comm_password is None:
            raise ValueError("华泰必须设置通讯密码")

        try:
            self._app = pywinauto.Application().connect(
                path=self._run_exe_path(exe_path), timeout=1
            )
        # pylint: disable=broad-except
        except Exception:
            self._app = pywinauto.Application().start(exe_path)

            # wait login window ready
            while True:
                try:
                    self._app.top_window().Edit1.wait("ready")
                    break
                except RuntimeError:
                    pass
            self._app.top_window().Edit1.set_focus()
            self._app.top_window().Edit1.type_keys(user)
            self._app.top_window().Edit2.type_keys(password)

            self._app.top_window().Edit3.type_keys(comm_password)

            self._app.top_window().button0.click()

            self._app = pywinauto.Application().connect(
                path=self._run_exe_path(exe_path), timeout=10
            )
        self._main = self._app.window(title="网上股票交易系统5.0")
        self._main.wait ( "exists enabled visible ready" , timeout=100 )
        self._close_prompt_windows ( )

    @property
    def balance(self):
        self._switch_left_menus(self._config.BALANCE_MENU_PATH)

        return self._get_balance_from_statics()

    def _get_balance_from_statics(self):
        result = {}
        for key, control_id in self._config.BALANCE_CONTROL_ID_GROUP.items():
            result[key] = float(
                self._main.child_window(
                    control_id=control_id, class_name="Static"
                ).window_text()
            )
        return result


class WKClientTrader(HTClientTrader):

    @property
    def broker_type(self):
        return "wk"

    def login(self, user, password, exe_path, comm_password=None, **kwargs):
        """
                :param user: 用户名
                :param password: 密码
                :param exe_path: 客户端路径, 类似
                :param comm_password:
                :param kwargs:
                :return:
                """
        self._editor_need_type_keys = False
        if comm_password is None:
            raise ValueError("五矿必须设置通讯密码")

        try:
            self._app = pywinauto.Application().connect(
                path=self._run_exe_path(exe_path), timeout=1
            )
        # pylint: disable=broad-except
        except Exception:
            self._app = pywinauto.Application().start(exe_path)

            # wait login window ready
            while True:
                try:
                    self._app.top_window().Edit1.wait("ready")
                    break
                except RuntimeError:
                    pass
            # self.login_test_host = False
            # if self.login_test_host:
            #     self._app.top_window().type_keys("%t")
            #     self.wait(0.5)
            #     try:
            #         self._app.top_window().Button2.wait('enabled', timeout=30, retry_interval=5)
            #         self._app.top_window().Button5.check()  # enable 自动选择
            #         self.wait(0.5)
            #         self._app.top_window().Button3.click()
            #         self.wait(0.3)
            #     except Exception as ex:
            #         logging.exception("test speed error", ex)
            #         self._app.top_window().wrapper_object().close()
            #         self.wait(0.3)

            self._app.top_window().Edit1.set_focus()
            self._app.top_window().Edit1.set_edit_text(user)
            self._app.top_window().Edit2.set_edit_text(password)

            self._app.top_window().Edit3.set_edit_text(comm_password)

            self._app.top_window().Button1.click()

            # detect login is success or not
            self._app.top_window().wait_not("exists", 100)

            self._app = pywinauto.Application().connect(
                path=self._run_exe_path(exe_path), timeout=10
            )
        self._close_prompt_windows()
        self._main = self._app.window(title="网上股票交易系统5.0")


class WKCreditClientTrader(HTClientTrader):
    """
    五矿信用账户
    """

    @property
    def broker_type(self):
        return "wk_credit"

    def login(self, user, password, exe_path, comm_password=None, **kwargs):
        """
                :param user: 用户名
                :param password: 密码
                :param credit_user 信用用户名
                :param credit_password 信用用户密码
                :param exe_path: 客户端路径, 类似
                :param comm_password:
                :param kwargs:
                :return:
                """
        self._editor_need_type_keys = False
        if comm_password is None:
            raise ValueError("五矿必须设置通讯密码")

        try:
            self._app = pywinauto.Application().connect(
                path=self._run_exe_path(exe_path), timeout=1
            )

        # pylint: disable=broad-except
        except Exception:
            self._app = pywinauto.Application().start(exe_path)

            # wait login window ready
            while True:
                try:
                    self._app.top_window().Edit1.wait("ready")
                    break
                except RuntimeError:
                    pass
            # self.login_test_host = False
            # if self.login_test_host:
            #     self._app.top_window().type_keys("%t")
            #     self.wait(0.5)
            #     try:
            #         self._app.top_window().Button2.wait('enabled', timeout=30, retry_interval=5)
            #         self._app.top_window().Button5.check()  # enable 自动选择
            #         self.wait(0.5)
            #         self._app.top_window().Button3.click()
            #         self.wait(0.3)
            #     except Exception as ex:
            #         logging.exception("test speed error", ex)
            #         self._app.top_window().wrapper_object().close()
            #         self.wait(0.3)

            self._app.top_window().Edit1.set_focus()
            self._app.top_window().Edit1.set_edit_text(user)
            self._app.top_window().Edit2.set_edit_text(password)

            self._app.top_window().Edit3.set_edit_text(comm_password)

            self._app.top_window().Button1.click()

            # detect login is success or not
            self._app.top_window().wait_not("exists", 100)

            self._app = pywinauto.Application().connect(
                path=self._run_exe_path(exe_path), timeout=10
            )
            self._close_prompt_windows()


            # click credit
            rect = self._app.top_window().CCustomTabCtrl1.rectangle()
            pos = (rect.right - 50, rect.top + 10)
            self._app.top_window().CCustomTabCtrl1.click_input(coords=pos, absolute=True)
            for x in range(0, 5):
                try:
                    if self._app.top_window().title == "信用交易认证":
                        break
                    else:
                        self.wait(1)
                except Exception as ex:
                    logging.info(ex)

            self._app.top_window().Edit1.set_edit_text(kwargs["credit_user"])
            self._app.top_window().Edit2.set_edit_text(kwargs["credit_password"])
            self._app.top_window().Button1.click()
            self.wait(2)

        self._main = self._app.window(title="网上股票交易系统5.0")
        # tr = self._app.top_window().SysTreeView323
        # for i in range(0, tr.ItemCount()):
        #     print(tr.GetItem([i]).text())
        # self.wait(2)

    @property
    def today_entrusts(self):
        self._switch_left_menus(self._config.TODAY_ENTRUSTS_MENU_PATH)
        self.refresh()
        return self._get_grid_data(self._config.COMMON_GRID_CONTROL_ID)