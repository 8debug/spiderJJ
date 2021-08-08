from time import sleep

from com.repair.Repair import Repair


class Car(Repair):
    def __init__(self):
        super().__init__()
        self.page = None
        self.s = "@"

    def set_page(self):
        str_page = self.browser.find_element_by_xpath('//li[@class="paginate_button active"]/a').text
        self.page = int(str_page)
        print("当前页面", str_page)

    def get_page_curr(self):
        curr_page = self.browser.find_element_by_xpath('//li[@class="paginate_button active"]/a').text
        return int(curr_page)

    def get_page_total(self):
        total_page = self.browser.find_element_by_xpath('//li[@class="paginate_button next"]/preceding-sibling::*[1]/a').text
        return int(total_page)

    def to_page(self, page):
        """
        跳转到指定页开始抓数据
        :param page:
        :return:
        """
        total_page = self.get_page_total()

        if abs(1-page) <= abs(total_page-page):
            page_elem = self.browser.find_element_by_xpath("//a[text()='下一页']")
            step = 1
        else:
            self.browser.find_element_by_xpath("//li[@class='paginate_button ']/a[text()="+str(total_page)+"]").click()
            self.util("//li[@class='paginate_button active']/a[text()=" + str(total_page) + "]")
            page_elem = self.browser.find_element_by_xpath("//a[text()='上一页']")
            step = -1

        while True:
            page_curr = self.get_page_curr()
            page_target = str(page_curr+step)
            if page_curr == page:
                return
            self.browser.find_element_by_xpath("//li[@class='paginate_button ']/a[text()="+str(page_target)+"]").click()
            self.util("//li[@class='paginate_button active']/a[text()=" + str(page_target) + "]")

    def start(self, page):
        xpath = "//a[@href='/qxgl/vehicle/list.do']"
        a = self.browser.find_element_by_xpath(xpath)
        a.click()
        self.util("//div[text()='车辆信息管理']")
        self.to_page(page)
        self.spider_data()
        self.next_page()
        self.excel_colse("车辆信息")
        self.close_browser()

    def next_page(self):
        """
        翻页
        :return:
        """
        # 滚到底部
        self.scroll_bottom()
        # 当前页
        self.set_page()
        page_next = self.browser.find_element_by_xpath("//a[text()='下一页']/parent::li")
        if "disabled" not in page_next.get_attribute("class"):
            page_next.find_element_by_xpath("./a").click()
            page = str(self.page + 1)
            self.util("//li[@class='paginate_button active']/a[text()=" + page + "]")
            self.spider_data()
            self.next_page()
        else:
            return

    def spider_data(self):
        """
        爬取数据
        :return:
        """
        try:
            array = self.browser.find_elements_by_xpath("//table[@id='vehicle-table']/tbody/tr")
            for tr in array:
                array_td = tr.find_elements_by_xpath("./td")
                vehicle_id = array_td[7].find_elements_by_xpath("./a")[0].get_attribute("vehicle-id")
                detail = self.detail(vehicle_id)
                print(detail)
                self.excel_append(detail)
        except Exception as e:
            self.excel_colse("车辆信息-异常退出")

    def detail(self, vehicle_id):
        url = 'http://60.16.11.38:93/qxgl/vehicle/edit.do?vehicle.id=' + vehicle_id
        self.open_switch(url)
        self.util("//div[text()='车辆信息管理']")
        # 所属机构
        org = self.browser.find_element_by_xpath("//*[@id='select2-chosen-2']").text
        # 使用部门
        dept = self.browser.find_element_by_xpath("//*[@id='select2-chosen-3']").text
        # 自编号
        code = self.browser.find_element_by_xpath("//*[@id='vehicle_selfNumbering']").get_property("value")
        # 品牌
        brand = self.browser.find_element_by_xpath("//*[@id='select2-chosen-4']").text
        # 车牌号
        license = self.browser.find_element_by_xpath("//*[@id='vehicle_licenseNumber']").get_property("value")
        # 车辆型号
        model = self.browser.find_element_by_xpath("//*[@id='select2-chosen-5']").text
        # 底盘型号
        chassis_model = self.browser.find_element_by_xpath("//*[@id='select2-chosen-6']").text
        # 车架号
        frame_code = self.browser.find_element_by_xpath("//*[@id='vehicle_frameNumber']").get_property("value")
        # 发动机号
        engine_code = self.browser.find_element_by_xpath("//*[@id='vehicle_engineNumber']").get_property("value")
        # 总质量
        total_mass = self.browser.find_element_by_xpath("//*[@id='vehicle_totalMass']").get_property("value")
        # 核定载质量
        approved_mass = self.browser.find_element_by_xpath("//*[@id='vehicle_ratifiedQuality']").get_property("value")
        # 购车时间
        buy_date = self.browser.find_element_by_xpath("//*[@id='vehicle_purchaseTime']").get_property("value")
        # 当前里程数
        mileage = self.browser.find_element_by_xpath("//*[@id='vehicle_currentMileage']").get_property("value")
        # 主机机油
        primary_oil = self.browser.find_element_by_xpath("//*[@id='vehicle_oilEngine']").get_property("value")
        # 副机机油
        deputy_oil = self.browser.find_element_by_xpath("//*[@id='vehicle_oilAuxiliaryEngine']").get_property("value")
        # 主机空滤
        primary_filter = self.browser.find_element_by_xpath("//*[@id='vehicle_airFiltration']").get_property("value")
        # 副机空滤
        deputy_filter = self.browser.find_element_by_xpath("//*[@id='vehicle_airAuxiliaryEngine']").get_property(
            "value")
        # 变速箱齿轮油
        gear_oil = self.browser.find_element_by_xpath("//*[@id='vehicle_shiftGearOil']").get_property("value")
        # 柴油滤
        diesel_filter = self.browser.find_element_by_xpath("//*[@id='vehicle_dieselFilter']").get_property("value")
        # 刹车油
        brake_oil = self.browser.find_element_by_xpath("//*[@id='vehicle_brakeFluid']").get_property("value")
        # 汽油滤
        gasoline_filter = self.browser.find_element_by_xpath("//*[@id='vehicle_gasolineFilter']").get_property("value")

        detail = self.s.join(
            ["所属机构", org, "使用部门", dept, "自编号", code, "品牌", brand,
             "车牌号", license, "车辆型号", model, "底盘型号", chassis_model,
             "车架号", frame_code, "发动机号", engine_code, "总质量", total_mass,
             "核定载质量", approved_mass, "购车时间", buy_date, "当前里程数", mileage,
             "主机机油", primary_oil, "副机机油", deputy_oil, "主机空滤", primary_filter,
             "副机空滤", deputy_filter, "变速箱齿轮油", gear_oil, "柴油滤", diesel_filter,
             "刹车油", brake_oil, "汽油滤", gasoline_filter])

        self.close_switch()

        return detail


spider = Car()
spider.open()
spider.start(176)
