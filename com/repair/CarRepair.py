from copy import copy

import xlrd
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from xlutils.copy import copy
from selenium.webdriver.support import expected_conditions as EC


class CarRepair:
    def __init__(self):
        """
        初始化 配置
        """
        # 实现无可视化界面
        options = Options()
        # options.add_argument('--headless')
        # options.add_argument('--disable-gpu')
        options.add_experimental_option('excludeSwitches', ['enable-automation'])
        self.browser = webdriver.Chrome(executable_path='./chromedriver.exe', options=options)
        self.browser.maximize_window()
        # self.browser.implicitly_wait(3)  # 全局隐式等待10秒

    def open(self):
        self.browser.get('http://60.16.11.38:93/qxgl/login/index.do')
        # 选择 投资目的
        self.util('.//input[@placeholder="账号"]')
        elem_username = self.browser.find_element_by_xpath('.//input[@placeholder="账号"]')
        elem_username.send_keys("wx-cy")
        elem_pwd = self.browser.find_element_by_xpath('.//input[@placeholder="密码"]')
        elem_pwd.send_keys("123456")
        button = self.browser.find_element_by_xpath('//button[@class="hp_loginbutbot"]')
        button.click()
        # 此处重点记忆
        str_xpath = '//li[@data-code="nav_qx_project"]//*[contains(text(), "维修项目管理")]'
        self.util_then_click(str_xpath)
        str_xpath = "//div[@class='dataTables_info'][contains(text(), '第 1 页')]"
        self.util(str_xpath)
        array = self.get_date()
        self.excel(array)
        return self

    def util(self, xpath):
        print("util===", xpath)
        WebDriverWait(self.browser, 10).until(EC.presence_of_all_elements_located((By.XPATH, xpath)))
        return self

    def util_then_click(self, xpath):
        self.util(xpath)
        self.browser.find_element_by_xpath(xpath).click()

    def get_date(self):
        str_xpath_next_wapper = "//li[@id='project-table_next']"
        str_xpath_next = str_xpath_next_wapper + "//a[text()='下一页']"
        str_xpath_data = "//table[@id='project-table']/tbody/tr"
        array = []
        self.browser.find_element_by_xpath('//li[@class="paginate_button "]/a[text()="66"]').click()
        self.util('//li[@class="paginate_button active"]/a[text()="66"]')
        while True:
            xpath_page_curr = '//li[@class="paginate_button active"]/a'
            page_curr = self.browser.find_element_by_xpath(xpath_page_curr).text
            paget_target = str(int(page_curr)+1)
            xpath_page_target = '//li[@class="paginate_button active"]/a[text()="'+paget_target+'"]'
            elem_next_wapper = self.browser.find_element_by_xpath(str_xpath_next_wapper)
            elem_next = self.browser.find_element_by_xpath(str_xpath_next)
            rows = self.browser.find_elements_by_xpath(str_xpath_data)
            debug_msg = self.browser.find_element_by_xpath("//div[@id='project-table_info']").text
            print(debug_msg)
            for tr in rows:
                name = tr.find_element_by_xpath(".//td[1]").text
                big = tr.find_element_by_xpath(".//td[2]").text
                mid = tr.find_element_by_xpath(".//td[3]").text
                min = tr.find_element_by_xpath(".//td[4]").text
                array.append({
                    "name": name,
                    "big": big,
                    "mid": mid,
                    "min": min,
                })
            str_classes = elem_next_wapper.get_attribute('class')
            if "disabled" in str_classes:
                print("扫码完毕，结束====")
                break
            elem_next.click()
            self.util(xpath_page_target)

        return array

    def excel(self, array):
        # 打开excel文件
        rb = xlrd.open_workbook('维修项目管理.xls', formatting_info=True)
        wb = copy(rb)
        sheet = wb.get_sheet(0)
        for i in range(len(array)):
            item = array[i]
            row = i + 0
            # 基金代码
            sheet.write(row, 0, item["name"])
            # 基金名
            sheet.write(row, 1, item["big"])
            # 基金成立时间
            sheet.write(row, 2, item["mid"])
            # 基金规模
            sheet.write(row, 3, item["min"])
        # 保存
        wb.save('维修项目管理1.xls')


car = CarRepair()
car.open()
