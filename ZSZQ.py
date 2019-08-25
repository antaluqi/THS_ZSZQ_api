import pywinauto
from pywinauto import clipboard
from pywinauto import keyboard
from const import BALANCE_CONTROL_ID_GROUP
import time,datetime
import pandas as pd
import io

'''
参考了 https://github.com/nladuo/THSTrader
'''

class API:

    def __init__(self,exe_path=r"D:\浙商证券独立委托系统\xiadan.exe"):
        print("正在连接客户端:", exe_path, "......")
        self.app = pywinauto.Application().connect(path=exe_path, timeout=10)
        print("连接成功!!!")
        self.main_wnd = self.app.top_window()

    """买入"""
    def buy(self, stock_no, price, amount):
        self.__select_menu(['买入[F1]'])
        return self.__trade(stock_no, price, amount)

    """卖出"""
    def sell(self, stock_no, price, amount):

        self.__select_menu(['卖出[F2]'])
        return self.__trade(stock_no, price, amount)

    """ 撤单 """
    def cancel_entrust(self, entrust_no):

        self.__select_menu(['撤单[F3]'])
        cancelable_entrusts = self._get_grid_data(is_entrust=True)  # 获取可以撤单的条目

        for i, entrust in enumerate(cancelable_entrusts):
            if str(entrust["合同编号"]) == str(entrust_no):  # 对指定合同进行撤单
                return self.__cancel_by_double_click(i)
        return {"success": False, "msg": "没找到指定订单"}

    """ 获取资金情况 """
    def get_balance(self):

        self.__select_menu(['查询[F4]', '资金股票'])

        result = {}
        for key, control_id in BALANCE_CONTROL_ID_GROUP.items():
            result[key] = float(
                self.main_wnd.window(control_id=control_id, class_name='Static').window_text()
            )
        return result

    """ 判断订单是否完成 """
    def check_trade_finished(self, entrust_no):

        self.__select_menu(['卖出[F2]'])
        self.__select_menu(['撤单[F3]'])
        cancelable_entrusts = self.__get_grid_data(is_entrust=True)
        for i, entrust in enumerate(cancelable_entrusts):
            if str(entrust["合同编号"]) == str(entrust_no):  # 如果订单未完成，就意味着可以撤单
                if entrust["成交数量"] == 0:
                    return False
        return True

    """ 获取持仓 """
    def get_position(self):

        self.__select_menu(['查询[F4]', '资金股票'])
        return self.__get_grid_data(index=3)

    """ 获取当日委托 """
    def get_today_entrusts(self):

        time.sleep(1)
        self.__select_menu(['查询[F4]', '当日委托'])
        return self.__get_grid_data(index=3)

    """获取当日成交"""
    def get_today_trades(self):

        self.__select_menu(['查询[F4]', '当日成交'])
        return self.__get_grid_data(index=3)

    """获取历史成交"""
    def get_history_trades(self,startD,endD):

        st=datetime.datetime.strptime(startD,'%Y-%m-%d')
        ed = datetime.datetime.strptime(endD, '%Y-%m-%d')

        self.__select_menu(['查询[F4]', '历史成交'])
        self.main_wnd.window(control_id=0x3f1, class_name="SysDateTimePick32").set_time(year=st.year, month=st.month, day=st.day)
        self.main_wnd.window(control_id=0x3f2, class_name="SysDateTimePick32").set_time(year=ed.year, month=ed.month, day=ed.day)
        self.main_wnd.window(control_id=0x3EE, class_name="Button").click()
        return self.__get_grid_data(index=3)

    """获取历史委托"""
    def get_history_entrusts(self,startD,endD):

        st = datetime.datetime.strptime(startD, '%Y-%m-%d')
        ed = datetime.datetime.strptime(endD, '%Y-%m-%d')
        self.__select_menu(['查询[F4]', '历史委托'])
        self.main_wnd.window(control_id=0x3f1, class_name="SysDateTimePick32").set_time(year=st.year, month=st.month, day=st.day)
        self.main_wnd.window(control_id=0x3f2, class_name="SysDateTimePick32").set_time(year=ed.year, month=ed.month, day=ed.day)
        self.main_wnd.window(control_id=0x3EE, class_name="Button").click()
        return self.__get_grid_data(index=3)

    """ 点击左边菜单 """
    def __select_menu(self, path):

        if r"网上股票" not in self.app.top_window().window_text():
            self.app.top_window().set_focus()
            pywinauto.keyboard.send_keys("{ENTER}")
        self.__get_left_menus_handle().get_item(path).click()

    """获取左边菜单句柄"""
    def __get_left_menus_handle(self):
        while True:
            try:
                handle = self.main_wnd.window(control_id=129, class_name='SysTreeView32')
                handle.wait('ready', 2)  # sometime can't find handle ready, must retry
                return handle
            except Exception as ex:
                print(ex)
                pass

    """交易"""
    def __trade(self, stock_no, price, amount):
        #time.sleep(0.2)
        self.main_wnd.window(control_id=0x408, class_name="Edit").set_text(str(stock_no))  # 设置股票代码
        time.sleep(0.1)
        self.main_wnd.window(control_id=0x409, class_name="Edit").set_text(str(price))  # 设置价格
        self.main_wnd.window(control_id=0x40A, class_name="Edit").set_text(str(amount))  # 设置股数目
        self.main_wnd.window(control_id=0x3EE, class_name="Button").click()   # 点击卖出or买入
        time.sleep(0.1)
        self.app.top_window().window(control_id=0x6, class_name='Button').click()  # 确定买入
        self.app.top_window().set_focus()
        result = self.app.top_window().window(control_id=0x3EC, class_name='Static').window_text()
        self.app.top_window().set_focus()
        try:
            self.app.top_window().window(control_id=0x2, class_name='Button').click()  # 确定
        except:
            pass
        return self.__parse_result(result)

    """获取grid里面的数据"""
    def __get_grid_data(self, index=3):

        grid = self.main_wnd.window(control_id=0x417, class_name='CVirtualGridCtrl')
        #time.sleep(0.1)
        grid.set_focus().right_click()  # 模拟右键
        for i in range(index):
           keyboard.send_keys('{DOWN}')
        keyboard.send_keys('{ENTER}')
        data = clipboard.GetData()
        df = pd.read_csv(io.StringIO(data), delimiter='\t', na_filter=False)
        return df.to_dict('records')




    @staticmethod
    def __parse_result(result):
        """ 解析买入卖出的结果 """

        # "您的买入委托已成功提交，合同编号：865912566。"
        # "您的卖出委托已成功提交，合同编号：865967836。"
        # "您的撤单委托已成功提交，合同编号：865967836。"
        # "系统正在清算中，请稍后重试！ "

        if r"已成功提交，合同编号：" in result:
            return {
                "success": True,
                "msg": result,
                "entrust_no": result.split("合同编号：")[1].split("。")[0]
            }
        else:
            return {
                "success": False,
                "msg": result
            }