from time import sleep

from com.repair.Repair import Repair


class RepairMan(Repair):
    def __init__(self):
        super().__init__()
        self.s = "@"

    def spider_user(self):
        xpath = "//a[text()='维修人员管理']"
        a = self.browser.find_element_by_xpath(xpath)
        a.click()
        self.util("//div[text()='汽修人员管理']")
        self.spider_data()
        self.next_page()
        self.excel_colse("维修员")
        self.close_browser()

    def next_page(self):
        """
        翻页
        :return:
        """
        # 滚到底部
        self.scroll_bottom()
        page_next = self.browser.find_element_by_xpath("//a[text()='下一页']/parent::li")
        if "disabled" not in page_next.get_attribute("class"):
            page_next.find_element_by_xpath("./a").click()
            sleep(3)
            self.util("//div[text()='汽修人员管理']")
            self.spider_data()
            self.next_page()
        else:
            return

    def spider_data(self):
        """
        爬取数据
        :return:
        """
        array = self.browser.find_elements_by_xpath("//table[@id='repairman-table']/tbody/tr")
        for tr in array:
            array_td = tr.find_elements_by_xpath("./td")
            name = array_td[0].text
            job = array_td[1].text
            dept = array_td[2].text
            sex = array_td[3].text
            mobile = array_td[4].text
            man_id = array_td[5].find_elements_by_xpath("./a")[0].get_attribute("repairman-id")
            info = self.s.join([name, job, dept, sex, mobile])
            detail = self.user_detail(man_id)
            content = self.s.join([info, detail])
            print(content)
            self.excel_append(content)

    def user_detail(self, man_id):
        url = 'http://60.16.11.38:93/qxgl/repairman/edit.do?repairman.id=' + man_id
        self.open_switch(url)
        self.util("//div[text()='人员管理信息']")
        # 所在机构
        dept = self.browser.find_element_by_xpath("//*[@id='select2-chosen-2']").text
        # 职务
        job = self.browser.find_element_by_xpath("//*[@id='select2-chosen-3']").text
        # 维修人员姓名
        name = self.browser.find_element_by_xpath("//*[@id='repairman_name']").text
        # 性别
        sex = self.browser.find_element_by_xpath("//*[@id='select2-chosen-4']").text
        # 联系电话
        mobile = self.browser.find_element_by_xpath("//*[@id='repairman_phone']").text
        # 身份证号
        card_id = self.browser.find_element_by_xpath("//*[@id='repairman_cardId']").text
        # 入职时间
        work_in_date = self.browser.find_element_by_xpath("//*[@id='repairman_entryDate']").text
        # 离职时间
        work_out_date = self.browser.find_element_by_xpath("//*[@id='repairman_leaveDate']").text
        # 状态
        state = self.browser.find_element_by_xpath("//*[@id='select2-chosen-5']").text
        # 技术等级
        level = self.browser.find_element_by_xpath("//*[@id='select2-chosen-6']").text
        # 可登录系统
        can_login = self.browser.find_element_by_xpath("//*[@id='select2-chosen-7']").text

        detail = self.s.join(
            [dept, job, name, sex, mobile, card_id, work_in_date, work_out_date, state, level, can_login])

        self.close_switch()

        return detail


man = RepairMan()
man.open()
man.spider_user()
