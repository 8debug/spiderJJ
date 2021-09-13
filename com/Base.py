import time

from openpyxl import Workbook
from selenium import webdriver
from selenium.common.exceptions import TimeoutException
from selenium.webdriver import ActionChains
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.remote.webelement import WebElement
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

"""
爬虫基础类
"""


class Base:
    def __init__(self, driver_name):
        """
        初始化 配置
        """
        self.wb = Workbook()
        self.ws = self.wb.create_sheet("Sheet", 0)
        self.pro_dir = 'E:/Project/pythonspace/spiderJJ'
        if driver_name is None:
            driver_name = "chromedriver.exe"
        # 实现无可视化界面
        self.result = []
        options = Options()

        # 利用此方式实现不显示浏览器窗口，会出现 element not interactable 异常，据说是--headless引起的
        # options.add_argument('--headless')
        # 禁用gpu
        # options.add_argument('--disable-gpu')

        # 利用此方式实现运行脚本，不显示浏览器窗口
        # 设置浏览器分辨率（窗口大小）
        options.add_argument("--window-size=1920,1080")
        # 最大化运行（全屏窗口）,不设置，取元素会报错
        options.add_argument("--start-maximized")
        # 开启无界面模式
        options.add_argument("--headless")

        options.add_experimental_option('excludeSwitches', ['enable-automation'])
        self.browser = webdriver.Chrome(executable_path=self.pro_dir + '/' + driver_name, options=options)
        self.browser.maximize_window()
        # self.browser.implicitly_wait(3)  # 全局隐式等待10秒

    def scroll_top(self):
        self.browser.execute_script("document.documentElement.scrollTop=0;")

    def scroll_bottom(self):
        self.browser.execute_script("window.scrollTo(0,document.body.scrollHeight);")

    def open(self):
        return self

    def close_browser(self):
        self.browser.close()
        self.browser.quit()

    # def util_eq(self, xpath):
    #     try:
    #         WebDriverWait(self.browser, 10).until(self.browser.find_element_by_xpath(xpath))
    #     except TimeoutException as e:
    #         self.browser.refresh()
    #         WebDriverWait(self.browser, 10).until(EC.presence_of_all_elements_located((By.XPATH, xpath)))
    #     return self

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

    def close_switch(self):
        """
        关闭当前页面，切换回上一个页面
        :return:
        """
        self.browser.close()
        all_handles = self.browser.window_handles  # 获取全部页面句柄
        self.browser.switch_to.window(all_handles.pop())  # 打开 最新弹出的页面

    def open_switch(self, arg):
        """
        点击元素打开新窗口并切换到新窗口，防止网址为blank的情况
        :param arg:
        :return:
        """
        if type(arg) is str and arg.startswith(("http://", "https://")) and self.browser.current_url != arg:
            self.browser.execute_script("window.open('" + arg + "');")
        elif type(arg) is WebElement and arg.tag_name == 'a' and arg.get_property("href") in ["", "#"]:
            ActionChains(self.browser).key_down(Keys.CONTROL).perform()
            arg.click()
        all_handles = self.browser.window_handles  # 获取全部页面句柄
        self.browser.switch_to.window(all_handles.pop())  # 打开 最新弹出的页面

    def excel_append(self, content):
        # self.result.append([content])
        # 可以使用append插入一行数据
        self.ws.append([content])

    def excel_colse(self, file_name):
        if file_name is None:
            file_name = str(int(time.time()))

        for i in range(len(self.result)):
            self.ws.append(self.result[i])
        self.wb.save(self.pro_dir + '/' + file_name + '.xlsx')
