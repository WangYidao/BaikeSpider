from selenium import webdriver                                         # 导入网页自动化测试工具selenium模块
from selenium.common.exceptions import NoSuchElementException          # 导入异常模块

from openpyxl import load_workbook                                     # 导入Excel文件读取模块

import time                                                            # 导入时间模块
import os                                                              # 导入操作系统模块

os.environ["SELENIUM_SERVER_JAR"] = "/Users/bin/Desktop/SourceTree/BaikeEnv/selenium/selenium-server-standalone-2.48.0.jar"   # 添加selenium服务器地址环境变量

print("Python Selenium Safari Started")                       # 程序开始运行提示

safari = webdriver.Safari()                                   # 打开safari浏览器

safari.set_window_position(0, 0)                              # 设置safari窗口位置
safari.set_window_size(640, 800)                              # 设置safari窗口大小

WB = load_workbook('test.xlsx')                               # 打开Excel WorkBook文件
WS = WB.worksheets[0]                                         # 打开第一个worksheet

for i in range(2,(WS.max_row + 1)):

    KeyWord = WS.cell(row = i,column = 1).value                       # 获取sheet第i行，第1列数据值并打印

    safari.get("http://baike.baidu.com/")                             # 打开网址
    safari.implicitly_wait(2)

    txt_search_key = safari.find_element_by_id("query")               # 按网页元素id查找网页元素query
    txt_search_key.clear()                                            # 清除输入框里的内容
    txt_search_key.send_keys(KeyWord)                                 # 将获取的数据添加到输入框里

    btn_search = safari.find_element_by_id("search")                  # 查找搜索按钮
    btn_search.click()                                                # 单击搜索按钮

    safari.implicitly_wait(2)

    try:
        safari.find_element_by_class_name("professional-con")         # 如果查找到profession-con元素，向输出sheet中指定单元格写入1
        print("已被收录")
        WS.cell(row = i, column = 2).value = 1
    except NoSuchElementException:                                    # 否则，想指定单元格写入1
        print("未被收录")
        WS.cell(row = i, column = 2).value = 0

    print(safari.current_url)                                         # 打印当前url地址
    time.sleep(2)

WB.save('output.xlsx')                                                # 保存输出的Excel文件
safari.close()                                                        # 关闭safari窗口

safari.quit()                                                         # 推出safari
