import selenium.webdriver as webdriver
from selenium.webdriver import ActionChains
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys

chrome_driver = 'D:/project/pythonspace/spiderEventhing/chromedriver2.exe'
options = Options()
options.add_experimental_option('excludeSwitches', ['enable-automation'])
browser = webdriver.Chrome(executable_path=chrome_driver, options=options)
browser.get('http://www.stats.gov.cn/tjsj/tjbz/tjyqhdmhcxhfdm/2020/index.html')
a_array = browser.find_elements_by_xpath("//tr[@class='provincetr']//a")
ActionChains(browser).key_down(Keys.CONTROL).perform()
a_array[0].click()
# browser.find_element_by_tag_name('body').send_keys(Keys.CONTROL + Keys.TAB)
browser.switch_to.window(browser.window_handles.pop())  # 打开 最新弹出的页面
