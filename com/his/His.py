from copy import copy

import xlrd
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from xlutils.copy import copy
from selenium.webdriver.support import expected_conditions as EC


class His:
    def __init__(self):
        """
        初始化 配置
        """
        # 实现无可视化界面
        options = Options()
        # options.add_argument('--headless')
        # options.add_argument('--disable-gpu')
        options.add_experimental_option('excludeSwitches', ['enable-automation'])
        self.browser = webdriver.Chrome(executable_path='D:\project\pythonspace\spiderEventhing\chromedriver.exe',
                                        options=options)
        self.browser.maximize_window()
        # self.browser.implicitly_wait(3)  # 全局隐式等待10秒

    def open(self):
        try:
            url = 'http://www.stats.gov.cn/tjsj/tjbz/tjyqhdmhcxhfdm/2020/index.html'
            self.browser.get(url)
            # 选择 投资目的
            self.util("//tr[@class='provincetr']//a[text()='北京市']")
            rows = self.browser.find_elements_by_xpath("//tr[@class='provincetr']//a")
            for a in rows:
                href = a.get_property('href')
                name = a.text
                code = str.replace(href, ".html", "")
                code = str.replace(code, "http://www.stats.gov.cn/tjsj/tjbz/tjyqhdmhcxhfdm/2020/", '')
                level = "-".join([code, name])
                self.level_2(level, href)
                break
        finally:
            self.browser.quit()
        return self

    def level_2(self, parent_leven, url):
        print(url)
        self._open_switch(url)
        self.util("//a[text()='京ICP备05034670号']")
        # 获取指定节点的父节点的所有向下兄弟节点的子节点a元素
        xpath_tr = "//td[text()='统计用区划代码']/parent::tr/following-sibling::*"
        rows = self.browser.find_elements_by_xpath("//td[text()='统计用区划代码']/parent::tr/following-sibling::*//a")
        if len(rows) > 0:
            array_tr = self.browser.find_elements_by_xpath(xpath_tr)
            for tr in array_tr:
                first = tr.find_elements_by_xpath(".//td")[0]
                last = tr.find_element_by_xpath(".//td[last()]")
                code = first.find_element_by_xpath(".//a").text
                href = first.find_element_by_xpath(".//a").get_property('href')
                name = last.text
                level = ",".join([parent_leven, code + "-" + name])
                self.level_2(level, href)
        else:
            array_tr = self.browser.find_elements_by_xpath(xpath_tr)
            for tr in array_tr:
                first = tr.find_elements_by_xpath(".//td")[0]
                last = tr.find_element_by_xpath(".//td[last()]")
                code = first.text
                name = last.text
                if name.find("多福巷社区居委会") > -1:
                    True
                level = ",".join([parent_leven, code + "-" + name])
                print(level)
            self._close_switch()
            print("关闭")

    def util(self, xpath):
        # print("util===", xpath)
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

    def _open_switch(self, url):
        """
        点击元素打开新窗口并切换到新窗口
        :param element:
        :return:
        """
        self.browser.execute_script("window.open();")
        all_handles = self.browser.window_handles  # 获取全部页面句柄
        self.browser.switch_to.window(all_handles.pop())  # 打开 最新弹出的页面
        self.browser.get(url)


his = His()
his.open()
