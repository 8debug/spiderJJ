from com.Base import Base


class Repair(Base):
    def __init__(self):
        super().__init__(None)

    def open(self):
        self.browser.get('http://60.16.11.38:93/qxgl/login/index.do')
        # 选择 投资目的
        self.util('.//input[@placeholder="账号"]')
        elem_username = self.browser.find_element_by_xpath('.//input[@placeholder="账号"]')
        elem_username.send_keys("wx-cy")
        elem_pwd = self.browser.find_element_by_xpath('.//input[@placeholder="密码"]')
        elem_pwd.send_keys("654321")
        button = self.browser.find_element_by_xpath('//button[@class="hp_loginbutbot"]')
        button.click()
        return self
