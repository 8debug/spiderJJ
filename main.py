# This is a sample Python script.

# Press Shift+F10 to execute it or replace it with your code.
# Press Double Shift to search everywhere for classes, files, tool windows, actions, and settings.

# 实现规避检测
import time
from selenium.webdriver import ChromeOptions
from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains

# Press the green button in the gutter to run the script.
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
import xlrd
import xlwt
from xlutils.copy import copy


def excel(result):
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
        sheet.write(row, 2, jj_item["jj_create_date"])
        # 基金规模
        sheet.write(row, 3, jj_item["jj_money"])
        # 经理名
        sheet.write(row, 4, jj_item["jj_manager_name"])
        # 经理从业时间
        sheet.write(row, 5, jj_item["jj_manager_datetime"])
        # 最大回撤
        sheet.write(row, 6, jj_item["jj_manager_max"])
        # 夏普率1/2/3年
        sheet.write(row, 7, jj_item["jj_sr_1"])
        sheet.write(row, 8, jj_item["jj_sr_2"])
        sheet.write(row, 9, jj_item["jj_sr_3"])
        # 晨星网评级3/5/10年
        # sheet.write(row, 10, jj_item["jj_star_3"])
        # sheet.write(row, 11, jj_item["jj_star_5"])
        # sheet.write(row, 12, jj_item["jj_star_10"])
        # 手续费
        sheet.write(row, 13, jj_item["jj_money_sxf"])
        # 购买起点
        sheet.write(row, 14, jj_item["jj_money_start"])
    # 保存
    wb.save('基金筛选结果.xls')


def spider_cx(array):
    bro.get('https://www.morningstar.cn/quickrank/default.aspx')
    bro.delete_all_cookies()
    cookie = {
        "name": "authWeb",
        "value": "1CCF8E6EC98E66DE98E5D080502681FA859D3053E57921AB2AF9C5B9A079656E12721F811B3971989C8186A45F729C712362FA6380338A1A17935CC4F360EEAEA3715442ECF30AD35FBECE94E51EF23CDE99A47B8E0E5C37C4D210177A5006A152A19C05F25BBB13201073C0B6FABBAA36E6B922",
        "domain": "www.morningstar.cn",
        "path": "/",
        "expires/Max-Age": "2021-02-25T07:13:04.424Z",
    }
    bro.add_cookie(cookie)
    bro.refresh()
    elem_input = bro.find_element_by_xpath('//*[@id="ctl00_cphMain_txtFund"]')
    for item in array:
        ActionChains(bro).move_to_element(elem_input).click().perform()
        elem_input.clear()
        elem_input.send_keys(item["jjdm"])
        bro.find_element_by_xpath('//*[@id="ctl00_cphMain_btnGo"]').click()
        trs = bro.find_elements_by_xpath('//*[@id="ctl00_cphMain_gridResult"]/tbody/tr')
        if len(trs) > 1:
            trs = trs.pop(0)
            for item_tr in trs:
                url_cx = item_tr.find_element_by_xpath('.//td[1]/a').get_attribute('href')
                item["url_cx"] = url_cx


def spider_detail(array):
    """
    爬 好买基金 详情页面
    :param array:
    :return:
    """
    for jj_item in array:
        bro.get(jj_item['url_detail'])
        time.sleep(2)
        # 基金规模
        jj_money = bro.find_element_by_xpath('/html/body/div[2]/div[3]/div/div[2]/div/div[2]/div[2]/ul/li[3]/span').text
        # 基金成立时间
        jj_create_date = bro.find_element_by_xpath(
            '/html/body/div[2]/div[3]/div/div[2]/div/div[2]/div[2]/ul/li[4]/span').text
        # 经理姓名
        jj_manager_name = bro.find_element_by_xpath('//*[@id="nTab3_0"]/div[1]/div/div[2]/div/ul[1]/li[1]/a').text
        # 从业时间
        jj_manager_datetime = bro.find_element_by_xpath(
            '//*[@id="nTab3_0"]/div[1]/div/div[2]/div/ul[1]/li[5]/span').text
        jj_manager_datetime = jj_manager_datetime.replace('从业时间：', '')
        # 最大回撤
        jj_manager_max = bro.find_element_by_xpath('//*[@id="nTab3_0"]/div[1]/div/div[2]/ul/li[3]/p[2]/span').text
        # 年化夏普比率 1年
        jj_sr_1 = bro.find_element_by_xpath(
            '//*[@id="nTab2_0"]/div[5]/div[2]/div[2]/div/div[3]/table/tbody/tr[3]/td[2]').text
        # 年化夏普比率 2年
        jj_sr_2 = bro.find_element_by_xpath(
            '//*[@id="nTab2_0"]/div[5]/div[2]/div[2]/div/div[3]/table/tbody/tr[3]/td[3]').text
        # 年化夏普比率 3年
        jj_sr_3 = bro.find_element_by_xpath(
            '//*[@id="nTab2_0"]/div[5]/div[2]/div[2]/div/div[3]/table/tbody/tr[3]/td[4]').text

        jj_item["jj_money"] = jj_money
        jj_item["jj_create_date"] = jj_create_date
        jj_item["jj_manager_name"] = jj_manager_name
        jj_item["jj_manager_datetime"] = jj_manager_datetime
        jj_item["jj_manager_max"] = jj_manager_max
        jj_item["jj_sr_1"] = jj_sr_1
        jj_item["jj_sr_2"] = jj_sr_2
        jj_item["jj_sr_3"] = jj_sr_3
    excel(array)
    bro.quit()
    return array


if __name__ == '__main__':
    option = ChromeOptions()
    option.add_experimental_option('excludeSwitches', ['enable-automation'])
    bro = webdriver.Chrome(executable_path='./chromedriver.exe', options=option)
    bro.implicitly_wait(10)  # 全局隐式等待10秒
    bro.get('https://www.howbuy.com/fundtool/filter.htm')
    # bro.maximize_window()  # 将浏览器最大化显示
    bro.find_element_by_xpath('/html/body/div[3]/div/div[1]/div[3]/a').click()
    # 3个月1/3
    elem_hover = bro.find_element_by_xpath('//*[@id="nTab1_0_1_0"]/div[3]/div[3]/div[1]/ul/li[3]')
    ActionChains(bro).move_to_element(elem_hover).perform()
    # print(elem_hover.find_element_by_xpath('.//li[@desc="yjpm2_业绩排名_2_近3月(前1/3)_0"]').get_attribute('outerHTML'))
    elem_hover.find_element_by_xpath('.//li[@desc="yjpm2_业绩排名_2_近3月(前1/3)_0"]').click()
    # 6个月1/3
    elem_hover = bro.find_element_by_xpath('//*[@id="nTab1_0_1_0"]/div[3]/div[3]/div[1]/ul/li[4]')
    ActionChains(bro).move_to_element(elem_hover).perform()
    elem_hover.find_element_by_xpath('.//li[@desc="yjpm3_业绩排名_2_近6月(前1/3)_0"]').click()
    # 1年1/4
    elem_hover = bro.find_element_by_xpath('//*[@id="nTab1_0_1_0"]/div[3]/div[3]/div[1]/ul/li[5]')
    ActionChains(bro).move_to_element(elem_hover).perform()
    elem_hover.find_element_by_xpath('.//li[@desc="yjpm4_业绩排名_1_近1年(前1/4)_0"]').click()
    # 2年1/4
    elem_hover = bro.find_element_by_xpath('//*[@id="nTab1_0_1_0"]/div[3]/div[3]/div[1]/ul/li[6]')
    ActionChains(bro).move_to_element(elem_hover).perform()
    elem_hover.find_element_by_xpath('.//li[@desc="yjpm5_业绩排名_1_近2年(前1/4)_0"]').click()
    # 3年1/4
    elem_hover = bro.find_element_by_xpath('//*[@id="nTab1_0_1_0"]/div[3]/div[3]/div[1]/ul/li[7]')
    ActionChains(bro).move_to_element(elem_hover).perform()
    elem_hover.find_element_by_xpath('.//li[@desc="yjpm6_业绩排名_1_近3年(前1/4)_0"]').click()
    # 每页显示100条
    bro.find_element_by_xpath('/html/body/div[2]/div[3]/div[2]/div[2]/div[2]/div[2]/div[1]/span[1]/a[3]').click()
    # 不加的话报错：Message: stale element reference: element is not attached to the page document
    # 错误原因：代码执行了click()，但是没有完成翻页，又爬了一次当前页，再执行翻页时页面已刷新，无法找到前面的翻页执行click()
    time.sleep(2)
    # 过滤出来的基金
    elem_table = bro.find_element_by_xpath('/html/body/div[2]/div[3]/div[2]/div[2]/div[2]/div[1]/table')
    elem_trs = elem_table.find_elements_by_tag_name('tr')
    array_jj = []
    for tr in elem_trs:
        elem_tds = tr.find_elements_by_xpath('.//td')
        # 基金代码
        jjdm = elem_tds[0].find_element_by_xpath('.//input').get_attribute('jjdm')
        # 基金名称
        jjjc = elem_tds[0].find_element_by_xpath('.//input').get_attribute('jjjc')
        # 手续费
        spans = elem_tds[13].find_elements_by_xpath('.//span')
        if len(spans) > 0:
            jj_money_sxf = spans[0].text
        else:
            jj_money_sxf = ""
        # 购买起点
        jj_money_start = elem_tds[14].text
        # print(elem_tds[1].find_element_by_xpath('.//a').get_attribute('outerHTML'))
        # 当第一个tr内的a标签执行click()操作时报错：Other element would receive the click
        # 原因：第一个tr由于使用滚屏已经在浏览器屏幕的外面，处于 不可见+不可操作 的状态，所以里面的a标签无法执行click；
        # 第十五个tr内的a标签执行click操作则不报错，因为此时tr在浏览器屏幕内处于 可见状态+可操作 状态
        # elem_tds[1].find_element_by_xpath('.//a').click()
        # print(elem_tds[1].text)
        url = elem_tds[1].find_element_by_xpath('.//a').get_attribute('href')
        jj = {
            'jjdm': jjdm,
            'jjjc': jjjc,
            'url_detail': url,
            'jj_money_sxf': jj_money_sxf,
            "jj_money_start": jj_money_start
        }
        array_jj.append(jj)
    spider_detail(array_jj)
