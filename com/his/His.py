import threading
from copy import copy
from multiprocessing import Pool
from time import sleep

from openpyxl import Workbook
from selenium import webdriver
from selenium.common.exceptions import TimeoutException
from selenium.webdriver import ActionChains
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

"""
并行爬取2020年《国家统计局用区划代码和城乡划分代码》
http://www.stats.gov.cn/tjsj/tjbz/tjyqhdmhcxhfdm/2020/index.html
"""


class His:
    def __init__(self, driver_name):
        """
        初始化 配置
        """
        self.wb = Workbook()
        self.ws = None
        self.file_name = None
        self.pro_dir = 'E:/Project/pythonspace/spiderJJ'
        # 实现无可视化界面
        self.result = []
        options = Options()
        options.add_argument('--headless')
        options.add_argument('--disable-gpu')
        options.add_experimental_option('excludeSwitches', ['enable-automation'])
        self.browser = webdriver.Chrome(executable_path=self.pro_dir + '/' + driver_name,
                                        options=options)
        self.browser.maximize_window()
        # self.browser.implicitly_wait(3)  # 全局隐式等待10秒

    def test(self):
        for i in range(10):
            self.result.append(["这是什么" + str(i)])
        self.excel_colse()
        self.browser.quit()

    def open(self, list_city):
        try:
            url = 'http://www.stats.gov.cn/tjsj/tjbz/tjyqhdmhcxhfdm/2020/index.html'
            self.browser.get(url)
            self.util("//tr[@class='provincetr']//a[text()='北京市']")
            rows = self.browser.find_elements_by_xpath("//tr[@class='provincetr']//a")
            for index in range(len(rows)):
                a = rows[index]
                name = a.text
                for city in list_city:
                    if name.find(city) > -1:
                        self.ws = self.wb.create_sheet(name, index)
                        self.file_name = name
                        self.level_2(None, a)
                        self.excel_append_result()
        except Exception as e:
            print("出现如下异常, 当前url%s", self.browser.current_url)
            print(e)
        finally:
            self.excel_colse()
            self.browser.quit()
        return self

    def level_2(self, content, element):
        self._open_switch(element)
        self.util("//*[text()='统计用区划代码']")
        # 获取指定节点的父节点的所有向下兄弟节点的子节点a元素
        xpath_tr = "//td[text()='统计用区划代码']/parent::tr/following-sibling::*"
        array_tr = self.browser.find_elements_by_xpath(xpath_tr)
        for tr in array_tr:
            array_td = tr.find_elements_by_xpath("./td")
            td_first = array_td[0]
            td_last = array_td[len(array_td) - 1]
            code = td_first.text
            name = td_last.text
            array_a = tr.find_elements_by_xpath(".//a")
            if len(array_a) > 0:
                a_first = td_first.find_element_by_xpath(".//a")
                href = a_first.get_property('href')
                if content is not None:
                    level = ",".join([content, code, name])
                else:
                    level = ",".join([code, name])
                print(level)
                self.excel_append(level)
                self.level_2(level, a_first)
            else:
                level = ",".join([content, code, name])
                print(level)
                self.excel_append(level)
        self._close_switch()

    def util(self, xpath):
        """
        爬虫出错重试
        :param xpath:
        :return:
        """
        try:
            WebDriverWait(self.browser, 10).until(EC.presence_of_all_elements_located((By.XPATH, xpath)))
        except TimeoutException as e:
            self.browser.refresh()
            WebDriverWait(self.browser, 10).until(EC.presence_of_all_elements_located((By.XPATH, xpath)))
        return self

    def util_then_click(self, xpath):
        self.util(xpath)
        self.browser.find_element_by_xpath(xpath).click()

    def _close_switch(self):
        """
        关闭当前页面，切换回上一个页面
        :return:
        """
        self.browser.close()
        all_handles = self.browser.window_handles  # 获取全部页面句柄
        self.browser.switch_to.window(all_handles.pop())  # 打开 最新弹出的页面

    def _open_switch(self, a):
        """
        点击元素打开新窗口并切换到新窗口，防止网址为blank的情况
        :param element:
        :return:
        """
        href = a.get_property("href")
        # # 打开新tab 方式二
        ActionChains(self.browser).key_down(Keys.CONTROL).perform()
        a.click()

        # 打开新tab 方式一
        # self.browser.execute_script("window.open();")
        # self.browser.get(url)
        all_handles = self.browser.window_handles  # 获取全部页面句柄
        self.browser.switch_to.window(all_handles.pop())  # 打开 最新弹出的页面
        if self.browser.current_url != href:
            self.browser.get(href)

    def excel_append_result(self):
        for i in range(len(self.result)):
            self.ws.append(self.result[i])
        self.result.clear()

    def excel_append(self, content):
        self.result.append([content])
        # 可以使用append插入一行数据
        # ws.append([content])

    def excel_colse(self):
        for i in range(len(self.result)):
            self.ws.append(self.result[i])
        self.wb.save(self.pro_dir + '/' + self.file_name + '.xlsx')


list_city = [
    '内蒙古自治区',
    '吉林省',
    '安徽省',
    '江西省',
    '山东省',
    '河南省',
    '湖北省',
    '广东省',
    '广西壮族自治区',
    '重庆市',
    '四川省',
    '贵州省',
    '云南省',
    '西藏自治区',
    '青海省',
]
for i in range(len(list_city)):
    city = list_city[i]
    his = His('chromedriver' + str(i + 1) + '.exe')
    threading.Thread(target=his.open, args=([city],)).start()
