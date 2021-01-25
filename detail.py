
import time
from selenium.webdriver import ChromeOptions
from selenium import webdriver
from lxml import etree
from selenium.webdriver.common.action_chains import ActionChains

if __name__ == '__main__':
    option = ChromeOptions()
    option.add_experimental_option('excludeSwitches', ['enable-automation'])
    bro_detail = webdriver.Chrome(executable_path='./chromedriver.exe', options=option)
    # bro_detail.implicitly_wait(10)  # 全局隐式等待10秒
    bro_detail.get('https://www.howbuy.com/fund/003834/')
    # 基金规模
    jj_money = bro_detail.find_element_by_xpath('/html/body/div[2]/div[3]/div/div[2]/div/div[2]/div[2]/ul/li[3]/span').text
    # 基金成立时间
    jj_date = bro_detail.find_element_by_xpath('/html/body/div[2]/div[3]/div/div[2]/div/div[2]/div[2]/ul/li[4]/span').text
    # 经理姓名
    work_name = bro_detail.find_element_by_xpath('//*[@id="nTab3_0"]/div[1]/div/div[2]/div/ul[1]/li[1]/a').text
    # 从业时间
    work_time = bro_detail.find_element_by_xpath('//*[@id="nTab3_0"]/div[1]/div/div[2]/div/ul[1]/li[5]/span').text
    # 最大回撤
    work_maximum_drawdown = bro_detail.find_element_by_xpath('//*[@id="nTab3_0"]/div[1]/div/div[2]/ul/li[2]/p[2]/span').text
    # 年化夏普比率 1年
    jj_SR_1 = bro_detail.find_element_by_xpath('//*[@id="nTab2_0"]/div[5]/div[2]/div[2]/div/div[3]/table/tbody/tr[3]/td[2]').text
    # 年化夏普比率 2年
    jj_SR_2 = bro_detail.find_element_by_xpath('//*[@id="nTab2_0"]/div[5]/div[2]/div[2]/div/div[3]/table/tbody/tr[3]/td[3]').text
    # 年化夏普比率 3年
    jj_SR_3 = bro_detail.find_element_by_xpath('//*[@id="nTab2_0"]/div[5]/div[2]/div[2]/div/div[3]/table/tbody/tr[3]/td[4]').text
    print("基金成立时间 {} 规模 {}".format(jj_date, jj_money))
    print("基金经理姓名 {} 从业时间 {} 从业最大回撤 {}".format(work_name, work_time.replace('从业时间：', ''), work_maximum_drawdown))
    print("夏普比率 1年 {} 2年 {} 3年 {}".format(jj_SR_1, jj_SR_2, jj_SR_3))
    # print("基金规模：%s" % money)
    bro_detail.quit()
