from time import sleep

from selenium.common.exceptions import NoSuchElementException

from com.repair.Repair import Repair


class Parts(Repair):
    def __init__(self):
        super().__init__()
        self.page = None
        self.s = "@"

    def set_current_page(self):
        str_page = self.current_page()
        self.page = int(str_page)

    def current_page(self):
        str_page = self.browser.find_element_by_xpath('//li[@class="paginate_button active"]/a').text
        print("当前页面", str_page)
        return str_page

    def parts(self):
        try:
            a = self.browser.find_element_by_xpath("//a[@href='/qxgl/part/list.do']")
            a.click()
            self.util("//div[text()='零件管理']")
            self.spider_data()
            self.next_page()
            self.excel_colse("零件")
            self.close_browser()
        except Exception as e:
            print(e)
            # self.close_browser()

    def next_page(self):
        """
        翻页
        :return:
        """
        # 当前页
        self.set_current_page()
        page_next = self.browser.find_element_by_xpath("//a[text()='下一页']/parent::li")
        if "disabled" not in page_next.get_attribute("class"):
            page_next.find_element_by_xpath("./a").click()
            page = str(self.page + 1)
            self.util("//li[@class='paginate_button active']/a[text()=" + page + "]")
            self.spider_data()
            self.next_page()
        else:
            return

    def util_list_page(self):
        self.util("//ul[@class='pagination']")

    def spider_data(self):
        """
        爬取数据
        :return:
        """
        self.util_list_page()
        current_page = self.current_page()
        array = self.browser.find_elements_by_xpath("//table[@id='part-table']/tbody/tr")
        for i in range(0, len(array)):
            tr = array[i]
            array_td = tr.find_elements_by_xpath("./td")
            part_id = array_td[9].find_elements_by_xpath("./a")[0].get_attribute("part-id")
            detail = self.detail(part_id)
            print("第{}页，第{}行，详情地址{}".format(current_page, (i+1), detail))
            self.excel_append(detail)

    def detail(self, part_id):
        url = 'http://60.16.11.38:93/qxgl/part/edit.do?part.id=' + part_id
        self.open_switch(url)
        self.util("//div[text()='零件信息']")
        self.util("//div[@id='statics-table1_paginate']/ul[@class='pagination']")
        # 库房选择
        sm = self.browser.find_element_by_xpath("//*[@id='select2-chosen-2']").text
        # 零件名称
        part_name = self.browser.find_element_by_xpath("//*[@id='part_partName']").get_property("value")
        # 所属配件架
        shelf = self.browser.find_element_by_xpath("//*[@id='select2-chosen-4']").text
        # 最低库存
        stock_min = self.browser.find_element_by_xpath("//*[@id='part_minStock']").get_property("value")
        # 品牌
        brand = self.browser.find_element_by_xpath("//*[@id='part_brand']").get_property("value")
        # 货商
        supplier = self.browser.find_element_by_xpath("//*[@id='part_grocer']").get_property("value")
        # 备注
        remark = self.browser.find_element_by_xpath("//*[@id='part_applyRange']").get_property("value")
        # 零件类型
        part_type = self.browser.find_element_by_xpath("//*[@id='select2-chosen-3']").text
        # 单位
        unit = self.browser.find_element_by_xpath("//*[@id='part_unit']").get_property("value")
        # 使用期限
        time_limit = self.browser.find_element_by_xpath("//*[@id='select2-chosen-5']").text
        # 规格
        standard = self.browser.find_element_by_xpath("//*[@id='part_specification']").get_property("value")
        # 条形码
        try:
            barcode = self.browser.find_element_by_xpath("//table[@id='statics-table1']/tbody/tr[1]/td[6]").text
        except NoSuchElementException as e:
            barcode = ''

        detail = self.s.join(
            ["库房选择", sm,
             "零件名称", part_name,
             "所属配件架", shelf,
             "最低库存", stock_min,
             "品牌", brand,
             "货商", supplier,
             "备注", remark,
             "零件类型", part_type,
             "单位", unit,
             "使用期限", time_limit,
             "规格", standard,
             "条形码", barcode
             ]
        )

        self.close_switch()

        return detail


spider = Parts()
spider.open()
spider.parts()
