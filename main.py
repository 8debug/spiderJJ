# This is a sample Python script.

# Press Shift+F10 to execute it or replace it with your code.
# Press Double Shift to search everywhere for classes, files, tool windows, actions, and settings.

# 实现规避检测
import time
from telnetlib import EC

from selenium.webdriver import ChromeOptions
from selenium import webdriver
from lxml import etree
from selenium.webdriver.common.action_chains import ActionChains

# Press the green button in the gutter to run the script.
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait

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
        print(elem_tds[0].find_element_by_xpath('.//input').get_attribute('jjdm'))
        # 基金名称
        jjjc = elem_tds[0].find_element_by_xpath('.//input').get_attribute('jjjc')
        print(elem_tds[0].find_element_by_xpath('.//input').get_attribute('jjjc'))
        # print(elem_tds[1].find_element_by_xpath('.//a').get_attribute('outerHTML'))
        # 当第一个tr内的a标签执行click()操作时报错：Other element would receive the click
        # 原因：第一个tr由于使用滚屏已经在浏览器屏幕的外面，处于 不可见+不可操作 的状态，所以里面的a标签无法执行click；
        # 第十五个tr内的a标签执行click操作则不报错，因为此时tr在浏览器屏幕内处于 可见状态+可操作 状态
        # elem_tds[1].find_element_by_xpath('.//a').click()
        # print(elem_tds[1].text)
        url = "https://www.howbuy.com" + elem_tds[1].find_element_by_xpath('.//a').get_attribute('href')
        jj = {'jjdm': jjdm, 'jjjc': jjjc, 'url_detail': url}
        array_jj.append(jj)
    bro.quit()


def handle_detail(array):
    bro_detail = webdriver.Chrome(executable_path='./chromedriver.exe', options=option)
    bro_detail.implicitly_wait(10)  # 全局隐式等待10秒
    for item in array:
        bro_detail.get(item.url_detail)
        money = bro_detail.find_element_by_xpath('//[@class=gmfund_num]/li[3]/span').text
        print("基金规模：%s" % money)
        bro_detail.quit()

# See PyCharm help at https://www.jetbrains.com/help/pycharm/
