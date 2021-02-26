import hashlib
import io
import time
from copy import copy

import requests
import xlrd
from PIL import Image
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.action_chains import ActionChains
from xlutils.copy import copy

# 订单数据
ORDERS_DATA = []


class TaoBao:
    def __init__(self):
        """
        初始化 配置
        """
        # 实现无可视化界面
        options = Options()
        # options.add_argument('--headless')
        # options.add_argument('--disable-gpu')
        # options.add_experimental_option('excludeSwitches', ['enable-automation'])
        # 避免被检测到window.navigator.webdriver为true
        options.add_argument("--disable-blink-features=AutomationControlled")
        self.browser = webdriver.Chrome(executable_path='../../chromedriver.exe', options=options)
        # self.browser.implicitly_wait(20)  # 全局隐式等待10秒
        self.browser.maximize_window()
        ORDERS_DATA.clear()

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

    def action_open(self):
        url = "https://login.taobao.com/member/login.jhtml"
        username = ''
        password = ''
        self.browser.get(url)
        # 用户名
        elem = self.browser.find_element_by_id('fm-login-id')
        elem.click()
        elem.send_keys(username)
        # 密码
        elem = self.browser.find_element_by_id('fm-login-password')
        elem.click()
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

        self.browser.find_element_by_xpath('//*[@id="login-form"]/div[4]/button').click()

    # 打开“已卖出宝贝”页面
    def open_sold(self):
        array = self.browser.find_elements_by_xpath(
            '//*[@id="module-open-aside"]/div/div/div/div/ul/li[3]/div[2]/p[1]/span')
        if array and array[0].text == '已卖出的宝贝':
            array[0].click()
            time.sleep(2)
            all_handles = self.browser.window_handles  # 获取全部页面句柄
            self.browser.switch_to.window(all_handles.pop())
            self.browser.find_element_by_xpath(
                '//*[@id="sold_container"]/div/div[1]/div[1]/form/div[2]/div[2]/div/div/input[1]')
        else:
            # TODO
            pass

    def action_date_range(self, start_datetime, end_datetime):
        """
        通过日期范围筛选订单
        :return:
        """
        if self.browser.title == "已卖出的宝贝":
            datetime = self.browser.find_element_by_xpath('//input[@placeholder="请选择时间范围起始"]')
            datetime.click()
            datetime.send_keys(start_datetime)

            datetime = self.browser.find_element_by_xpath('//input[@placeholder="请选择时间范围结束"]')
            datetime.click()
            datetime.send_keys(end_datetime)

            btn_filter = self.browser.find_element_by_xpath('//button[text()="搜索订单"]')
            btn_filter.click()
            time.sleep(5)
            array_order_detail = self.browser.find_elements_by_xpath('//a[text()="详情"]')
            for detail in array_order_detail:
                xpath = './ancestor::tr'
                if self._exist(xpath):
                    curr_goods = detail.find_element_by_xpath(xpath)
                    other_goods = curr_goods.find_elements_by_xpath('./following-sibling::tr')
                    all_goods = other_goods.insert(0, curr_goods)
                    array_data = []
                    for index, goods in enumerate(all_goods):
                        goods_code_label = goods.find_element_by_xpath('./span[text()="商家编码："]')
                        goods_code_value = ''
                        if self._exist(goods_code_label):
                            goods_code_value = goods_code_label.find_element_by_xpath('./ancestor::p').text
                        detail.click()
                        all_handles = self.browser.window_handles  # 获取全部页面句柄
                        self.browser.switch_to.window(all_handles.pop())  # 打开 交易详情 界面
                        data = self.action_pay_goods(index, goods_code_value)
                        array_data.append(data)
                price = self.action_pay_order()



            array_order = self.browser.find_elements_by_css_selector('.item-mod__trade-order___2LnGB')
            for elem_order in array_order:
                # 订单号
                order_id = elem_order.find_element_by_css_selector('input[name=orderid]').get_attribute('value')
                goods_data = []
                order = {"id": order_id, "goods_data": goods_data}
                # 订单下的所有商品
                array_goods = elem_order.find_elements_by_xpath('/table[2]/tr')
                for index, elem_goods in enumerate(array_goods):
                    goods_code = elem_goods.find_element_by_xpath('//span[text()="商家编码："]/following-sibling::*[1]').text
                    elem_goods.find_element_by_xpath('//a[text()="详情"]').click()
                    all_handles = self.browser.window_handles  # 获取全部页面句柄
                    self.browser.switch_to.window(all_handles.pop())  # 打开 交易详情 界面
                    price = self.action_pay_order()
                    order['price'] = price
                    self.action_pay_goods(index, goods_code)  # 关闭 交易详情 界面
                    goods_dict = self.browser.switch_to.window(all_handles.pop())  # 切换回上一个界面
                    goods_data.append(goods_dict)
                ORDERS_DATA.append(order)

    def _exist(self, xpath):
        array = self.browser.find_elements_by_xpath(xpath)
        return len(array) > 0

    def action_pay_order(self):
        """
        进入 交易详情 界面
        价格构成
        :return:
        """
        if self.browser.title.find('交易详情') != -1:
            elem_order = self.browser.find_element_by_xpath('//*[text()="订单编号"]/ancestor::li')
            elem_order.text
            elements = self.browser.find_elements_by_css_selector('.total-count-wrapper')
            # 先删除最后一个元素
            elements.pop()
            # 价格构成
            wrapper_price_total = elements[0]
            array_price = wrapper_price_total.find_elements_by_css_selector('.total-count-line')
            # 实际成交价
            price_real = elements.pop()
            price_data = [price_real.text]
            for item in array_price:
                price_data.append(item.text)
            return price_data

    def action_pay_goods(self, index, goods_code):
        """
        进入 交易详情 界面
        :param goods_code: 商家编码
        :param index:
        :return:    {"goods_code": 商家编码, "goods_cost": 商品实际消费, "goods_count": 商品数量, "goods_status": 商品成交状态}
        """
        if self.browser.title.find('交易详情') != -1:
            array_goods = self.browser.find_elements_by_xpath('//*[@id="appOrders"]/div/table/tbody/tr/td/ul/li/table')
            if array_goods:
                goods = array_goods[index]
                goods_cost = goods.find_element_by_xpath('/tr[0]/td[1]').text
                goods_count = goods.find_element_by_xpath('/tr[0]/td[2]').text
                goods_status = goods.find_element_by_xpath('/tr[0]/td[4]').text
                self.browser.close()
                return {
                    "goods_cost": goods_cost,
                    "goods_count": goods_count,
                    "goods_status": goods_status,
                    "goods_code": goods_code,
                }


spider = TaoBao()
spider.action_open()
