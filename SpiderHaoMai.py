import time
from copy import copy

import xlrd
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.action_chains import ActionChains
from xlutils.copy import copy


class SpiderHaoMai:

    def __init__(self):
        """
        初始化 配置
        """
        # 实现无可视化界面
        options = Options()
        options.add_argument('--headless')
        options.add_argument('--disable-gpu')
        options.add_experimental_option('excludeSwitches', ['enable-automation'])

        self.browser = webdriver.Chrome(executable_path='./chromedriver.exe', options=options)
        self.browser.implicitly_wait(20)  # 全局隐式等待10秒
        # self.browser.maximize_window()

    def action_open(self):
        self.browser.get('https://www.howbuy.com/fundtool/filter.htm')
        # 选择 投资目的
        self.browser.find_element_by_xpath('/html/body/div[3]/div/div[1]/div[3]/a').click()
        return self

    def action_by_zs(self):
        """
        切换到指数
        :return:
        """
        self.browser.find_element_by_xpath('//*[@id="nTab1_0_1_t3"]').click()
        return self

    def action_by_4433(self):
        """
        按照 4433 法则过滤基金
        :return:
        """
        # 3个月1/3
        elem_hover = self.browser.find_element_by_xpath('//*[@id="nTab1_0_1_3"]/div[5]/div[3]/div[1]/ul/li[3]')
        ActionChains(self.browser).move_to_element(elem_hover).perform()
        elem_hover.find_element_by_xpath('.//li[@desc="yjpm2_业绩排名_2_近3月(前1/3)_0"]').click()
        # 6个月1/3
        elem_hover = self.browser.find_element_by_xpath('//*[@id="nTab1_0_1_3"]/div[5]/div[3]/div[1]/ul/li[4]')
        ActionChains(self.browser).move_to_element(elem_hover).perform()
        elem_hover.find_element_by_xpath('.//li[@desc="yjpm3_业绩排名_2_近6月(前1/3)_0"]').click()
        # 1年1/4
        elem_hover = self.browser.find_element_by_xpath('//*[@id="nTab1_0_1_3"]/div[5]/div[3]/div[1]/ul/li[5]')
        ActionChains(self.browser).move_to_element(elem_hover).perform()
        elem_hover.find_element_by_xpath('.//li[@desc="yjpm4_业绩排名_1_近1年(前1/4)_0"]').click()
        # 2年1/4
        elem_hover = self.browser.find_element_by_xpath('//*[@id="nTab1_0_1_3"]/div[5]/div[3]/div[1]/ul/li[6]')
        ActionChains(self.browser).move_to_element(elem_hover).perform()
        elem_hover.find_element_by_xpath('.//li[@desc="yjpm5_业绩排名_1_近2年(前1/4)_0"]').click()
        # 3年1/4
        elem_hover = self.browser.find_element_by_xpath('//*[@id="nTab1_0_1_3"]/div[5]/div[3]/div[1]/ul/li[7]')
        ActionChains(self.browser).move_to_element(elem_hover).perform()
        elem_hover.find_element_by_xpath('.//li[@desc="yjpm6_业绩排名_1_近3年(前1/4)_0"]').click()
        return self

    def action_by_cyb(self):
        """
        过滤出 创业板 基金
        :return:
        """
        elem_input = self.browser.find_element_by_xpath('//*[@id="fund_keywords"]')
        ActionChains(self.browser).move_to_element(elem_input).click().perform()
        elem_input.clear()
        elem_input.send_keys('创业板')
        self.browser.find_element_by_xpath('//*[@id="keywords_btn"]').click()
        return self

    def action_by_100(self):
        # 每页显示100条
        self.browser.find_element_by_xpath(
            '/html/body/div[2]/div[3]/div[2]/div[2]/div[2]/div[2]/div[1]/span[1]/a[3]').click()
        time.sleep(2)
        return self

    def get_data_list(self):
        """
        过滤出基金列表信息
        :return:
        """
        # 显示等待
        # array_tr = WebDriverWait(self.browser, 10).until(
        #     EC.presence_of_all_elements_located(
        #         (By.XPATH, "/html/body/div[2]/div[3]/div[2]/div[2]/div[2]/div[1]/table//tr"))
        # )
        array_tr = self.browser.find_elements_by_xpath('/html/body/div[2]/div[3]/div[2]/div[2]/div[2]/div[1]/table//tr')
        array_result = []
        for tr in array_tr:
            checkbox = tr.find_element_by_xpath('.//td[1]/input')
            jjdm = checkbox.get_attribute('jjdm')
            jjjc = checkbox.get_attribute('jjjc')
            url = tr.find_element_by_xpath('.//td[2]/a').get_attribute('href')
            # 手续费
            spans = tr.find_elements_by_xpath('.//td[14]//span')
            if len(spans) > 0:
                money_sxf = spans[0].text
            else:
                money_sxf = ""
            # 购买起点
            money_point = tr.find_element_by_xpath('.//td[15]/span').text
            data = {
                "jjdm": jjdm,
                "jjjc": jjjc,
                "url": url,
                "money_sxf": money_sxf,
                "money_point": money_point,
            }
            data_detail = self.get_data_detail(url)
            data.update(data_detail)
            array_result.append(data)
        return array_result

    def get_data_detail(self, url):
        """
        基金详情，爬取 规模 | 成立时间 | 基金经理名 | 基金经理介绍地址 | 经理从业时长 | 经理最大回撤 | 夏普率1/2/3年
        :param url:
        :return:
        """
        old_tab = self.browser.window_handles[0]
        self.browser.execute_script("window.open('');")
        new_tab = self.browser.window_handles[1]
        self.browser.switch_to.window(new_tab)
        self.browser.get(url)
        money = self.browser.find_element_by_xpath(
            '/html/body/div[2]/div[3]/div/div[2]/div/div[2]/div[2]/ul/li[3]/span').text
        money = money.replace('亿', '')
        create_date = self.browser.find_element_by_xpath(
            '/html/body/div[2]/div[3]/div/div[2]/div/div[2]/div[2]/ul/li[4]/span').text
        elem_workname = self.browser.find_element_by_xpath('//*[@id="nTab3_0"]/div[1]/div/div[2]/div/ul[1]/li[1]/a')
        work_name = elem_workname.text
        work_url = elem_workname.get_attribute('href')
        work_time = self.browser.find_element_by_xpath('//*[@id="nTab3_0"]/div[1]/div/div[2]/div/ul[1]/li[5]/span').text
        work_time = work_time.replace('从业时间：', '')
        work_max_revoke = self.browser.find_element_by_xpath(
            '//*[@id="nTab3_0"]/div[1]/div/div[2]/ul/li[3]/p[2]/span').text
        sr1 = self.browser.find_element_by_xpath(
            '//*[@id="nTab2_0"]/div[5]/div[2]/div[2]/div/div[3]/table/tbody/tr[3]/td[2]').text
        sr2 = self.browser.find_element_by_xpath(
            '//*[@id="nTab2_0"]/div[5]/div[2]/div[2]/div/div[3]/table/tbody/tr[3]/td[3]').text
        sr3 = self.browser.find_element_by_xpath(
            '//*[@id="nTab2_0"]/div[5]/div[2]/div[2]/div/div[3]/table/tbody/tr[3]/td[4]').text
        self.browser.close()
        self.browser.switch_to.window(old_tab)

        return {
            'money': money,
            'create_date': create_date,
            'work_name': work_name,
            'work_url': work_url,
            'work_time': work_time,
            'work_max_revoke': work_max_revoke,
            'sr1': sr1,
            'sr2': sr2,
            'sr3': sr3
        }

    def excel(self, result):
        """
        根据模版生成excel
        :param result:
        :return:
        """
        # 打开excel文件
        rb = xlrd.open_workbook('筛选基金条件.xls', formatting_info=True)
        # 获得要操作的页
        # r_sheet = rb.sheet_by_index(0)
        wb = copy(rb)
        sheet = wb.get_sheet(0)
        for i in range(len(result)):
            jj_item = result[i]
            row = i + 3
            # 基金代码
            sheet.write(row, 0, jj_item["jjdm"])
            # 基金名
            sheet.write(row, 1, jj_item["jjjc"])
            # 基金成立时间
            sheet.write(row, 2, jj_item["create_date"])
            # 基金规模
            sheet.write(row, 3, jj_item["money"])
            # 经理名
            sheet.write(row, 4, jj_item["work_name"])
            # 经理从业时间
            sheet.write(row, 5, jj_item["work_time"])
            # 最大回撤
            sheet.write(row, 6, jj_item["work_max_revoke"])
            # 夏普率1/2/3年
            sheet.write(row, 7, jj_item["sr1"])
            sheet.write(row, 8, jj_item["sr2"])
            sheet.write(row, 9, jj_item["sr3"])
            # 晨星网评级3/5/10年
            # sheet.write(row, 10, jj_item["year3"])
            # sheet.write(row, 11, jj_item["year5"])
            # sheet.write(row, 12, jj_item["year10"])
            # 手续费
            sheet.write(row, 13, jj_item["money_sxf"])
            # 购买起点
            sheet.write(row, 14, jj_item["money_point"])
        # 保存
        wb.save('基金筛选结果.xls')


spider = SpiderHaoMai()
result = spider.action_open().action_by_zs().action_by_4433().action_by_cyb().action_by_100().get_data_list()
spider.excel(result)
