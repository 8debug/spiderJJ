# This is a sample Python script.

# Press Shift+F10 to execute it or replace it with your code.
# Press Double Shift to search everywhere for classes, files, tool windows, actions, and settings.

# 实现规避检测
import io
import time

import requests
from selenium.webdriver import ChromeOptions
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.action_chains import ActionChains

# Press the green button in the gutter to run the script.
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
import xlrd
import xlwt
from xlutils.copy import copy
import hashlib
from PIL import Image


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
        sheet.write(row, 10, jj_item["year3"])
        sheet.write(row, 11, jj_item["year5"])
        sheet.write(row, 12, jj_item["year10"])
        # 手续费
        sheet.write(row, 13, jj_item["jj_money_sxf"])
        # 购买起点
        sheet.write(row, 14, jj_item["jj_money_start"])
    # 保存
    wb.save('基金筛选结果.xls')


def spider_cx_detail(array):
    '''
    晨星网基金详情，爬3/5/10年的评级
    :param array:
    :return:
    '''
    for item in array:
        bro.get(item["url"])
        year3_src = bro.find_element_by_xpath('//*[@id="qt_star"]/li[6]/img').get_attribute('src')
        year5_src = bro.find_element_by_xpath('//*[@id="qt_star"]/li[7]/img').get_attribute('src')
        year10_src = bro.find_element_by_xpath('//*[@id="qt_star"]/li[8]/img').get_attribute('src')
        item['year3'] = get_stars_by_image(year3_src)
        item['year5'] = get_stars_by_image(year5_src)
        item['year10'] = get_stars_by_image(year10_src)
    bro.quit()
    excel(array)


def get_stars_by_image(url):
    star5_standard = 'a8e420353e92219531bbdbf31e31728a'
    star4_standard = 'de55f4c0ca9463e5b6591a276e523bc9'
    star3_standard = '9210d3c3cf673c7fbebe317a90e604bb'
    star2_standard = '1729ddcb120dd221d358095b6f3a3762'
    star1_standard = '99f77f88e0ff9319b06c4e754edbd836'
    star0_standard = '0dc4da1ee45b6df501925c8bc2c814ce'
    image_hash = hashlib.md5(Image.open(io.BytesIO(requests.get(url).content)).tobytes()).hexdigest()

    stars = '☆☆☆☆☆'
    if image_hash == star1_standard:
        stars = '★☆☆☆☆'
    elif image_hash == star2_standard:
        stars = '★★☆☆☆'
    elif image_hash == star3_standard:
        stars = '★★★☆☆'
    elif image_hash == star4_standard:
        stars = '★★★★☆'
    elif image_hash == star5_standard:
        stars = '★★★★★'
    return stars


def spider_cx(array):
    bro.get('https://www.morningstar.cn/quickrank/default.aspx')
    bro.delete_all_cookies()
    cookie = {
        "name": "authWeb",
        "value": "E7EBF11698C41E7618DF5BE0AC01375E85F240C73C8B956366304798031D42F073A1EDC62559C0469EB5B611941433BFB9B1BC60D91F6CE94746F4E7633F04ACC9F57FB617453D96656731DDB2684D0CC68F65921BE5D7ED787C7458CD8AE16CA11827B29F032CCE6B1945F1E6F08D8237C0AA14",
        "domain": "www.morningstar.cn",
        "path": "/",
        "expires/Max-Age": "2021-02-25T07:13:04.424Z",
    }
    bro.add_cookie(cookie)
    bro.refresh()
    for item in array:
        elem_input = bro.find_element_by_xpath('//*[@id="ctl00_cphMain_txtFund"]')
        ActionChains(bro).move_to_element(elem_input).click().perform()
        elem_input.clear()
        elem_input.send_keys(item["jjdm"])
        bro.find_element_by_xpath('//*[@id="ctl00_cphMain_btnGo"]').click()
        trs = bro.find_elements_by_xpath('//table[@id="ctl00_cphMain_gridResult"]//tr')
        if len(trs) > 1:
            # 删除头行
            trs.pop(0)
            for item_tr in trs:
                item["url"] = item_tr.find_element_by_xpath('.//td[3]/a').get_attribute('href')
    spider_cx_detail(array)


def spider_detail(array):
    """
    爬 好买基金 详情页面
    :param array:
    :return:
    """
    for jj_item in array:
        bro.get(jj_item['url'])
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
    spider_cx(array)
    excel(array)
    return array


if __name__ == '__main__':
    # 实现无可视化界面
    chrome_option = Options()
    chrome_option.add_argument('--headless')
    chrome_option.add_argument('--disable-gpu')

    option = ChromeOptions()
    option.add_experimental_option('excludeSwitches', ['enable-automation'])
    bro = webdriver.Chrome(executable_path='./chromedriver.exe', options=option, chrome_options=chrome_option)
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
            'url': url,
            'jj_money_sxf': jj_money_sxf,
            "jj_money_start": jj_money_start
        }
        array_jj.append(jj)
    spider_detail(array_jj)
