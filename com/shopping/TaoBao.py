import json
import time
import traceback
from copy import copy

import xlrd
from selenium import webdriver
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.wait import WebDriverWait
from xlutils.copy import copy

# 订单数据
ORDERS_DATA = []


class TaoBao:
    def __init__(self):
        """
        初始化 配置
        """
        # 实现无可视化界面
        options = webdriver.ChromeOptions()
        options.add_experimental_option("detach", True)
        # options = Options()
        # options.add_argument('--headless')
        # options.add_argument('--disable-gpu')
        # options.add_experimental_option('excludeSwitches', ['enable-automation'])
        # 避免被检测到window.navigator.webdriver为true
        options.add_argument("--disable-blink-features=AutomationControlled")
        self.browser = webdriver.Chrome(executable_path='../../chromedriver.exe', options=options)
        # self.browser.implicitly_wait(20)  # 全局隐式等待10秒
        self.browser.maximize_window()
        ORDERS_DATA.clear()

    def click_loading_complete_username(self, elem):
        """
        首页点击用户名时等待loading动画消失
        :param elem:
        :return:
        """
        elem.click()
        time.sleep(1)
        self._until('//div[@class="loading-mask"]', until_not=True)

    def click_loading_complete_main(self, elem):
        """
        主页面loading动画消失
        :param elem:
        :return:
        """
        elem.click()
        time.sleep(2)
        # loading 动画消失
        self._until('//div[contains(@class, "loading-mod__loading")][contains(@class, "loading-mod__hidden")]')

    # def _click_loading(self, element):
    #     """
    #     带有loading效果的点击事件
    #     :param element:
    #     :return:
    #     """
    #     element.click()
    #     self._until_not_loading()

    def __drag_horizontally(self, source, target):
        """
        操作滑块水平滑动
        :param source:
        :param target:
        :return:
        """
        x_start = source.location.get('x')
        x_end = x_start + target.rect.get('width') - source.rect.get('width')
        y = source.location.get('y')
        drag_distance = x_end - x_start
        step_distance = drag_distance
        step_count = drag_distance
        while step_count > 0:
            if drag_distance % step_count == 0 and step_count <= 20:
                step_distance = drag_distance / step_count
                break
            step_count -= 1
        ActionChains(self.browser).click_and_hold(source).perform()
        while step_count > 0:
            ActionChains(self.browser).move_to_element(source) \
                .move_by_offset(step_distance, y) \
                .perform()
        time.sleep(1)

    def action_check_slider(self):
        """
        判断是否出现滑块
        :return:
        """
        array = self.browser.find_elements_by_xpath('//span[text()="请按住滑块，拖动到最右边"]')
        if array:
            source = self.browser.find_element_by_xpath('//*[@id="nc_1_n1z"]')
            target = array[0]
            self.__drag_horizontally(source, target)

    def _until(self, xpath, timeout=10, step=0.5, until_not=False):
        """
        等待页面中指定元素加载完成
        :param xpath:   定位元素表达式
        :param timeout: 等待时间
        :return:
        """
        try:
            wait = WebDriverWait(self.browser, timeout, step)
            locate = EC.presence_of_element_located((By.XPATH, xpath))
            if not until_not:
                wait.until(locate)
            else:
                wait.until_not(locate)
        except TimeoutException:
            print("在 " + self.browser.title + " 页面下在" + str(timeout) + "秒内无法根据 " + xpath + " 表达式找到元素")
            print("可能页面结构发生变化，请及时纠正爬虫程序")
            traceback.print_exc()

    def open_login(self, username, password):
        """
        登录淘宝
        :param username:
        :param password:
        :return:
        """
        url = "https://login.taobao.com/member/login.jhtml"
        self.browser.get(url)
        self._until('//button[text()="登录"]')
        # 用户名
        elem = self.browser.find_element_by_xpath('//input[@placeholder="会员名/邮箱/手机号"]')
        self.click_loading_complete_username(elem)
        elem.send_keys(username)
        # 密码
        elem = self.browser.find_element_by_xpath('//input[@placeholder="请输入登录密码"]')
        self.click_loading_complete_username(elem)
        elem.send_keys(password)
        # 滑块
        frame_id = "baxia-dialog-content"
        array_elem = self.browser.find_elements_by_id(frame_id)

        # 出现滑块时
        if array_elem:
            self.browser.switch_to.frame(frame_id)
            source = self.browser.find_element_by_id('nc_1_n1z')
            target = self.browser.find_element_by_id('nc_1__scale_text')
            self.__drag_horizontally(source, target)
            self.browser.switch_to.parent_frame()

        self.browser.find_element_by_xpath('//button[text()="登录"]').click()
        self._until('//span[text()="已卖出的宝贝"]', timeout=30, step=0.2)
        return self

    # 打开“已卖出宝贝”页面
    def open_sold(self):
        elem = self.browser.find_element_by_xpath('//span[text()="已卖出的宝贝"]')
        elem.click()
        all_handles = self.browser.window_handles  # 获取全部页面句柄
        self.browser.switch_to.window(all_handles.pop())
        self._until('//button[text()="搜索订单"]')
        return self

    def _open_switch(self, element):
        """
        点击元素打开新窗口并切换到新窗口
        :param element:
        :return:
        """
        element.click()
        all_handles = self.browser.window_handles  # 获取全部页面句柄
        self.browser.switch_to.window(all_handles.pop())  # 打开 最新弹出的页面

    def _close_switch(self):
        """
        关闭当前页面，切换回上一个页面
        :return:
        """
        self.browser.close()
        all_handles = self.browser.window_handles  # 获取全部页面句柄
        self.browser.switch_to.window(all_handles.pop())  # 打开 最新弹出的页面

    def _exception(self, message):
        raise Exception(message)

    def scroll_until(self, xpath, step_sleep=2):
        """
        滚动屏幕，直到检测到xpath表达式元素
        :param step_sleep: 每滚动一屏的休息秒数
        :param xpath:
        :return targets: 返回找到的目标元素
        """
        start_height = 0
        step_height = self.browser.execute_script('return document.documentElement.clientHeight')
        start_time = time.time()
        while True:
            targets = self.browser.find_elements_by_xpath(xpath)
            if len(targets) > 0:
                return targets

            js = 'window.scrollTo(' + str(start_height) + ', ' + str(step_height) + ')'
            self.browser.execute_script(js)
            time.sleep(step_sleep)
            start_height += step_height
            step_height += step_height

            end_time = time.time()
            second = round(end_time - start_time, 2)
            if second > 60:
                self._exception('60s滚动屏幕没有找到指定元素，视为页面结构发生变化')

    def _until_not_loading(self, xpath='//div[contains(@class, "loading-mod__loading")][contains(@class, '
                                       '"loading-mod__hidden")]'):
        """
        直到loading效果消失
        :return:
        """
        time.sleep(2)
        # loading 动画消失
        self._until(xpath)

    def input_daterange(self, start_datetime, end_datetime, order_id):
        """
        通过日期范围筛选订单
        :return:
        """
        if self.browser.title == "已卖出的宝贝":
            if order_id is not None:
                input_order_id = self.browser.find_element_by_id('bizOrderId')
                input_order_id.click()
                input_order_id.send_keys(order_id)
            datetime = self.browser.find_element_by_xpath('//input[@placeholder="请选择时间范围起始"]')
            datetime.click()
            # 只能通过插件的输入框录入时间，再过滤才有效
            xpath_input_datetime = '//input[contains(@class, "rc-calendar-input")]'
            datetime_input = self._exist_with_exception(self.browser, xpath_input_datetime)[0]
            datetime_input.send_keys(start_datetime)
            datetime_ok = self._exist(self.browser, '//a[@class="rc-calendar-ok-btn"]')[0]
            datetime_ok.click()

            datetime = self.browser.find_element_by_xpath('//input[@placeholder="请选择时间范围结束"]')
            datetime.click()
            datetime_input = self._exist_with_exception(self.browser, xpath_input_datetime)[0]
            datetime_input.send_keys(end_datetime)
            datetime_ok = self._exist(self.browser, '//a[@class="rc-calendar-ok-btn"]')[0]
            datetime_ok.click()

            btn_search = self.browser.find_element_by_xpath('//button[text()="搜索订单"]')
            self.click_loading_complete_main(btn_search)
            # 页面滚动至分页处
            self.scroll_until('//li[@title="上一页"]')

            page_more = self._exist(self.browser, '//button[text()="显示更多页码"]')
            if page_more:
                self.click_loading_complete_main(page_more)

            pages = self.browser.find_elements_by_css_selector('ul.pagination li')
            array_order_data = []
            if len(pages) > 2:
                pages.pop()  # 去掉 ”下一页“
                page_last = pages.pop().text  # 获取最后一页页码
                # page_last = "2"
                for page in range(1, int(page_last) + 1):
                    page_target = self.browser.find_element_by_css_selector(
                        'ul.pagination li[title="' + str(page) + '"]')
                    self.click_loading_complete_main(page_target)
                    array_order_data.extend(self._handle_order())
                self.to_excel(array_order_data)

    def _handle_order(self):
        """
        开始爬取订单数据
        :return:
        """
        array_order_detail = self._exist_with_exception(self.browser, '//a[text()="详情"]')
        array_order_data = []
        for detail in array_order_detail:
            print(detail.get_attribute('href'))
            array_goods = self._exist_with_exception(detail, './/ancestor::tbody/tr')
            array_goods_code = []

            # 商品中存在没有商家编码的情况标识
            for index, goods in enumerate(array_goods):
                elements = self._exist(goods, './/span[text()="补差价专用链接"]')
                elements_code = self._exist(goods, './/span[text()="商家编码："]/ancestor::p')
                goods_code = None
                if elements_code:
                    goods_code = elements_code[0].text
                elif elements:
                    goods_code = "补差价专用链接"
                if goods_code is not None:
                    array_goods_code.append(goods_code)

            array_goods_data = []
            self._open_switch(detail)
            order_data = self.action_pay_order()
            array_order_data.append(order_data)
            for index, goods_code in enumerate(array_goods_code):
                goods_data = self.action_pay_goods(index, goods_code)
                array_goods_data.append(goods_data)
            order_data["array_goods"] = array_goods_data
            # 关闭订单交易详情
            self._close_switch()
            array_order_data.append(order_data)
        return array_order_data

    def to_excel(self, array_order):
        # 打开excel文件
        rb = xlrd.open_workbook('../../爬取订单宝贝数据.xls', formatting_info=True)
        wb = copy(rb)
        sheet = wb.get_sheet(0)
        for index, data in enumerate(array_order):
            # 基金代码
            sheet.write(index, 0, json.dumps(data))
        # 保存
        wb.save('../../爬取订单宝贝数据-结果.xls')

    def _exist(self, element, xpath):
        """
        元素是否存在
        :param xpath:
        :return:
        """
        array = element.find_elements_by_xpath(xpath)
        return array
        # if len(array) == 0:
        #     return False
        # else:
        #     return array

    def _exist_with_exception(self, element, xpath):
        """
        元素是否存在，不存在则抛出异常
        :param xpath:
        :return:
        """
        array = element.find_elements_by_xpath(xpath)
        if len(array) == 0:
            raise Exception("无法通过 " + xpath + " 表达式找到元素")
        else:
            return array

    def action_pay_order(self):
        """
        进入 交易详情 界面
        价格构成
        :return:
        """
        if self.browser.title.find('交易详情') != -1:
            order_status = self.browser.find_element_by_xpath('//h3[contains(text(), "订单状态:")]').text
            order_code = self.browser.find_element_by_xpath('//span[text()="订单编号"]/ancestor::li').text
            order_pay = None
            order_pay_happen1 = self._exist(self.browser, '//span[text()="应收款"]/parent::*')
            order_pay_happen2 = self._exist(self.browser, '//span[text()="实收款"]/parent::*')
            order_pay_happen3 = self._exist(self.browser, '//span[text()="应收款"]/parent::*')
            if order_pay_happen1:
                order_pay = order_pay_happen1[0].text
            elif order_pay_happen2:
                order_pay = order_pay_happen2[0].text
            elif order_pay_happen3:
                order_pay = order_pay_happen3[0].text

            return {
                "order_status": order_status,
                "order_code": order_code,
                "order_pay": order_pay,
            }

    def action_pay_goods(self, index, goods_code):
        """
        进入 交易详情 界面
        :param goods_code: 商家编码
        :param index:
        :return:    {"goods_code": 商家编码, "goods_cost": 商品实际消费, "goods_count": 商品数量, "goods_status": 商品成交状态}
        """
        if self.browser.title.find('交易详情') != -1:
            array_goods = self._exist_with_exception(self.browser, '//li[@class="bought-listform-content"]/table')
            for goods in array_goods:
                goods_cost = goods.find_element_by_xpath('.//tr/td[2]').text
                goods_count = goods.find_element_by_xpath('.//tr/td[3]').text
                goods_status = goods.find_element_by_xpath('.//tr/td[5]').text
                return {
                    "goods_cost": goods_cost,
                    "goods_count": goods_count,
                    "goods_status": goods_status,
                    "goods_code": goods_code,
                }


spider = TaoBao()
# spider.action_open("", "")
spider.open_login("优目旗舰店:凯伦", "hsyk1234") \
    .open_sold() \
    .input_daterange('2021-03-01 0:00:00', '2021-03-03 0:00:00', '1609808906857913256')
