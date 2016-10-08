from selenium import webdriver                                         # 导入网页自动化测试工具selenium模块
from selenium.common.exceptions import NoSuchElementException          # 导入异常模块

from openpyxl import load_workbook                                     # 导入Excel文件读取模块

import time                                                            # 导入时间模块
import os                                                              # 导入操作系统模块

os.environ["SELENIUM_SERVER_JAR"] = "/Users/bin/Desktop/SourceTree/BaikeEnv/selenium/selenium-server-standalone-2.48.0.jar"   # 添加selenium服务器地址环境变量

print("Python Selenium Safari Started")                       # 程序开始运行提示

safari = webdriver.Safari()                                   # 打开safari浏览器

WB = load_workbook('test_find_element.xlsx')                  # 打开Excel WorkBook文件
WS = WB.worksheets[0]                                         # 打开第一个worksheet

WS.cell(row = 1, column = 1).value = "词条名称"
WS.cell(row = 1,column = 2).value = "是否被\"百度百科\"收录"
WS.cell(row = 1,column = 3).value = "是否被\"科普中国百科\"收录"
WS.cell(row = 1,column = 4).value = "词条网址"

WS.cell(row = 1, column = 5).value = "概述字数"
WS.cell(row = 1, column = 6).value = "基本信息栏条数"
WS.cell(row = 1, column = 7).value = "一级目录条数"
WS.cell(row = 1, column = 8).value = "二级目录条数"
WS.cell(row = 1, column = 9).value = "正文段数"
WS.cell(row = 1, column = 10).value = "正文字数"
WS.cell(row = 1, column = 11).value = "参考文献条数"
WS.cell(row = 1, column = 12).value = "词条图册数"
WS.cell(row = 1, column = 13).value = "词条图片张数"

print("表格行数：%d" %(WS.max_row))                                    # debug

for i in range(2,(WS.max_row + 1)):

    KeyWord = WS.cell(row = i,column = 1).value                       # 获取sheet第i行，第1列数据值并打印

    safari.get("http://baike.baidu.com/")                             # 打开网址
    safari.implicitly_wait(2)

    baike_search_key = safari.find_element_by_id("query")               # 按网页元素id查找网页元素query
    baike_search_key.clear()                                            # 清除输入框里的内容
    baike_search_key.send_keys(KeyWord)                                 # 将获取的数据添加到输入框里

    safari.find_element_by_id("search").click()                         # 单击搜索按钮

    safari.implicitly_wait(2)

    #查询词条是否被"百度百科"收录
    try:
        safari.find_element_by_class_name("create-entrance")

        #若"百度百科首页"查询未收录，查询"百度首页"
        safari.get("http://wwww.baidu.com/")
        safari.implicitly_wait(2)

        baidu_search_key = safari.find_element_by_id("kw")
        baidu_search_key.clear()
        baidu_search_key.send_keys(KeyWord)

        safari.find_element_by_id("su").click()
        safari.implicitly_wait(2)

        try:
            safari.find_elements_by_partial_link_text("百度百科")

            WS.cell(row = i, column = 2).value = "已收录"

            #获取搜索结果页中包含"百度百科"字样的链接
            Search_Results = safari.find_elements_by_partial_link_text("百度百科")

            Switch_Link = Search_Results[0].get_attribute('href')

            print(Switch_Link)                                                           # Debug

            #打开词条页面
            safari.get(Switch_Link)
            safari.implicitly_wait(2)

            WS.cell(row = i, column = 4).value = safari.current_url

            # 获取特征点
            # 获取概述段落的字数
            try:
                safari.find_element_by_xpath("//div[@label-module='lemmaSummary']")
                WS.cell(row=i, column=5).value = len(
                    safari.find_element_by_xpath("//div[@label-module='lemmaSummary']").text)
            except NoSuchElementException:
                WS.cell(row=i, column=5).value = "-1"

            print("概述字数")
            print(WS.cell(row=i, column=5).value)  # debug

            # 获取基本信息栏条数
            try:
                safari.find_element_by_xpath("//dt[@class='basicInfo-item name']")
                print("找到基本信息栏")
                WS.cell(row=i, column=6).value = len(
                    safari.find_elements_by_xpath("//dt[@class='basicInfo-item name']"))
            except NoSuchElementException:
                WS.cell(row=i, column=6).value = "-1"

            print("基本信息栏条数")
            print(WS.cell(row=i, column=6).value)  # debug

            # 获取一级目录条数
            try:
                safari.find_element_by_xpath("//div[@class='lemma-catalog']/div/ol/li[@class='level1']")
                print("找到一级目录")
                WS.cell(row=i, column=7).value = len(
                    safari.find_elements_by_xpath("//div[@class='lemma-catalog']/div/ol/li[@class='level1']"))
            except NoSuchElementException:
                WS.cell(row=i, column=7).value = "-1"

            print("一级目录条数")
            print(WS.cell(row=i, column=7).value)  # debug

            # 获取二级目录
            try:
                safari.find_element_by_xpath("//div[@class='lemma-catalog']/div/ol/li[@class='level2']")
                print("找到二级目录")
                WS.cell(row=i, column=8).value = len(
                    safari.find_elements_by_xpath("//div[@class='lemma-catalog']/div/ol/li[@class='level2']"))
            except NoSuchElementException:
                WS.cell(row=i, column=8).value = "-1"

            print("二级目录条数")
            print(WS.cell(row=i, column=8).value)  # debug

            # 获取正文段数及字数
            try:
                safari.find_element_by_xpath("//div[@label-module='para']")
                Content_Paragraphs = safari.find_elements_by_xpath("//div[@label-module='para']")
                WS.cell(row=i, column=9).value = len(Content_Paragraphs)

                Number_Of_Words = 0
                for Content_Paragraph in Content_Paragraphs:
                    Number_Of_Words += len(Content_Paragraph.text)

                WS.cell(row=i, column=10).value = Number_Of_Words
            except NoSuchElementException:
                WS.cell(row=i, column=9).value = "-1"
                WS.cell(row=i, column=10).value = "-1"

            print("正文段落数")
            print(WS.cell(row=i, column=9).value)  # debug
            print("正文字数")
            print(WS.cell(row=i, column=10).value)

            # 获取参考文献条数
            try:
                safari.find_elements_by_xpath("//ul[@class='reference-list']/li[@class='reference-item ']")
                print("找到参考文献")
                WS.cell(row=i, column=11).value = len(
                    safari.find_elements_by_xpath("//ul[@class='reference-list']/li[@class='reference-item ']")) + len(
                    safari.find_elements_by_xpath("//ul[@class='reference-list']/li[@class='reference-item more']"))
            except NoSuchElementException:
                WS.cell(row=i, column=11).value = -1

            print("参考文献条数")
            print(WS.cell(row=i, column=11).value)

            # 获取词条图册数量
            try:
                safari.find_element_by_link_text("更多图册")
                print("找到多个图册")
                safari.get(safari.find_element_by_link_text("更多图册").get_attribute('href'))
                safari.implicitly_wait(2)

                WS.cell(row=i, column=12).value = safari.find_element_by_xpath("//span[@class='album-num num']").text
                WS.cell(row=i, column=13).value = safari.find_element_by_xpath("//span[@class='pic-num num']").text

                safari.get(safari.find_element_by_class_name("return-back").get_attribute('href'))
                safari.implicitly_wait(2)
            except NoSuchElementException:
                try:
                    safari.find_element_by_class_name("summary-pic")
                    print("找到一个图册")
                    safari.get(safari.find_element_by_xpath("//div[@class='summary-pic']/a").get_attribute('href'))
                    safari.implicitly_wait(2)

                    WS.cell(row=i, column=12).value = 1
                    WS.cell(row=i, column=13).value = safari.find_element_by_xpath(
                        "//span[@style='color:#427cb8']").text

                    safari.get(safari.find_element_by_link_text("返回词条").get_attribute('href'))
                    safari.implicitly_wait(2)
                except NoSuchElementException:
                    WS.cell(row=i, column=12).value = -1
                    WS.cell(row=i, column=13).value = -1

            print("词条图册数量")
            print(WS.cell(row=i, column=12).value)
            print("词条图片数量")
            print(WS.cell(row=i, column=13).value)

        except NoSuchElementException:
            WS.cell(row = i, column = 2).value = "未收录"

    except NoSuchElementException:
        WS.cell(row = i, column = 2).value = "已收录"
        WS.cell(row = i, column = 4).value = safari.current_url

        # 获取特征点
        # 获取概述段落的字数
        try:
            safari.find_element_by_xpath("//div[@label-module='lemmaSummary']")
            WS.cell(row=i, column=5).value = len(
                safari.find_element_by_xpath("//div[@label-module='lemmaSummary']").text)
        except NoSuchElementException:
            WS.cell(row=i, column=5).value = "-1"

        print("概述字数")
        print(WS.cell(row=i, column=5).value)  # debug

        # 获取基本信息栏条数
        try:
            safari.find_element_by_xpath("//dt[@class='basicInfo-item name']")
            print("找到基本信息栏")
            WS.cell(row=i, column=6).value = len(safari.find_elements_by_xpath("//dt[@class='basicInfo-item name']"))
        except NoSuchElementException:
            WS.cell(row=i, column=6).value = "-1"

        print("基本信息栏条数")
        print(WS.cell(row=i, column=6).value)  # debug

        # 获取一级目录条数
        try:
            safari.find_element_by_xpath("//div[@class='lemma-catalog']/div/ol/li[@class='level1']")
            print("找到一级目录")
            WS.cell(row=i, column=7).value = len(
                safari.find_elements_by_xpath("//div[@class='lemma-catalog']/div/ol/li[@class='level1']"))
        except NoSuchElementException:
            WS.cell(row=i, column=7).value = "-1"

        print("一级目录条数")
        print(WS.cell(row=i, column=7).value)  # debug

        # 获取二级目录
        try:
            safari.find_element_by_xpath("//div[@class='lemma-catalog']/div/ol/li[@class='level2']")
            print("找到二级目录")
            WS.cell(row=i, column=8).value = len(
                safari.find_elements_by_xpath("//div[@class='lemma-catalog']/div/ol/li[@class='level2']"))
        except NoSuchElementException:
            WS.cell(row=i, column=8).value = "-1"

        print("二级目录条数")
        print(WS.cell(row=i, column=8).value)  # debug

        # 获取正文段数及字数
        try:
            safari.find_element_by_xpath("//div[@label-module='para']")
            Content_Paragraphs = safari.find_elements_by_xpath("//div[@label-module='para']")
            WS.cell(row=i, column=9).value = len(Content_Paragraphs)

            Number_Of_Words = 0
            for Content_Paragraph in Content_Paragraphs:
                Number_Of_Words += len(Content_Paragraph.text)

            WS.cell(row=i, column=10).value = Number_Of_Words
        except NoSuchElementException:
            WS.cell(row=i, column=9).value = "-1"
            WS.cell(row=i, column=10).value = "-1"

        print("正文段落数")
        print(WS.cell(row=i, column=9).value)  # debug
        print("正文字数")
        print(WS.cell(row=i, column=10).value)

        # 获取参考文献条数
        try:
            safari.find_elements_by_xpath("//ul[@class='reference-list']/li[@class='reference-item ']")
            print("找到参考文献")
            WS.cell(row=i, column=11).value = len(safari.find_elements_by_xpath("//ul[@class='reference-list']/li[@class='reference-item ']")) + len(
                safari.find_elements_by_xpath("//ul[@class='reference-list']/li[@class='reference-item more']"))
        except NoSuchElementException:
            WS.cell(row=i, column=11).value = -1

        print("参考文献条数")
        print(WS.cell(row=i, column=11).value)

        # 获取词条图册数量
        try:
            safari.find_element_by_link_text("更多图册")
            print("找到多个图册")
            safari.get(safari.find_element_by_link_text("更多图册").get_attribute('href'))
            safari.implicitly_wait(2)

            WS.cell(row=i, column=12).value = safari.find_element_by_xpath("//span[@class='album-num num']").text
            WS.cell(row=i, column=13).value = safari.find_element_by_xpath("//span[@class='pic-num num']").text

            safari.get(safari.find_element_by_class_name("return-back").get_attribute('href'))
            safari.implicitly_wait(2)
        except NoSuchElementException:
            try:
                safari.find_element_by_class_name("summary-pic")
                print("找到一个图册")
                safari.get(safari.find_element_by_xpath("//div[@class='summary-pic']/a").get_attribute('href'))
                safari.implicitly_wait(2)

                WS.cell(row=i, column=12).value = 1
                WS.cell(row=i, column=13).value = safari.find_element_by_xpath("//span[@style='color:#427cb8']").text

                safari.get(safari.find_element_by_link_text("返回词条").get_attribute('href'))
                safari.implicitly_wait(2)
            except NoSuchElementException:
                WS.cell(row=i, column=12).value = -1
                WS.cell(row=i, column=13).value = -1

        print("词条图册数量")
        print(WS.cell(row=i, column=12).value)
        print("词条图片数量")
        print(WS.cell(row=i, column=13).value)

    # 查询词条是否被"科普中国百科"收录
    try:
        safari.find_element_by_class_name("professional-con")
        WS.cell(row = i, column = 3).value = "已收录"
    except NoSuchElementException:
        WS.cell(row = i, column = 3).value = "未收录"

    time.sleep(2)

WB.save('output_find_element.xlsx')                                                # 保存输出的Excel文件
safari.close()                                                        # 关闭safari窗口

safari.quit()                                                         # 推出safari
