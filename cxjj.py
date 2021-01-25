import time
from selenium.webdriver import ChromeOptions
from selenium import webdriver
from lxml import etree
from selenium.webdriver.common.action_chains import ActionChains

# 晨星基金排行榜爬取数据

def create_cookie(name, value, domain, path, expires):
    return {
        "name": name,
        "value": value,
        "domain": domain,
        "path": path,
        "expires/Max-Age": expires,
    }


if __name__ == '__main__':
    option = ChromeOptions()
    option.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/88.0.4324.96 Safari/537.36")
    option.add_experimental_option('excludeSwitches', ['enable-automation'])
    bro_cx = webdriver.Chrome(executable_path='./chromedriver_cx.exe', options=option)
    bro_cx.implicitly_wait(20)  # 全局隐式等待10秒
    bro_cx.get('https://www.morningstar.cn/quickrank/default.aspx')
    bro_cx.delete_all_cookies()
    cookies = []

    # cookies.append(create_cookie("ASP.NET_SessionId",
    #                              "njr4oq55thp2qtvhi1m3oc45",
    #                              "www.morningstar.cn",
    #                              "/",
    #                              "2021-02-25T07:13:04.424Z",
    #                              ))
    cookies.append(create_cookie("authWeb",
                                 "FBEB3108D7D3D8777146D13EAE7E649955C77A5D5BFA7987891AA9BF66C70827CCF461461803EB4DEEE317BF0E20750CDF6C6D180CD6EBD5172AC88E34D1B3AEE7B5C4B9A11F515583EC06AAA7E784E7CAEA1720CB1444C8A508B039594263BCE9E9B77D03E933DC155ADB8D394608B30C6C4F7A",
                                 "www.morningstar.cn",
                                 "/",
                                 "2021-02-25T07:13:04.424Z",
                                 ))
    # cookies.append(create_cookie("user",
    #                              "username=297179121@qq.com&nickname=%e9%98%bf%e6%96%af%e9%a1%bfKing&status=Free&password=5OaTjzzxON4=",
    #                              "www.morningstar.cn",
    #                              "/",
    #                              "2021-02-25T07:13:04.424Z",
    #                              ))

    for cookie in cookies:
        bro_cx.add_cookie(cookie)

    # 设置cookies后刷新当前页面就好使了，不知道为啥
    bro_cx.refresh()

    elem_input = bro_cx.find_element_by_xpath('//*[@id="ctl00_cphMain_txtFund"]')
    ActionChains(bro_cx).move_to_element(elem_input).click().perform()
    elem_input.send_keys('165516')
    bro_cx.find_element_by_xpath('//*[@id="ctl00_cphMain_btnGo"]').click()
    # time.sleep(5)
    # bro_detail.quit()
