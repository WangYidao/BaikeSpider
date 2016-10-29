#!/usr/bin/env python3
# -*- coding: utf-8 -*-

'Module to get Wikipedia entry information'

_author_ = "Yidao@Babacheku"

import time
import os

from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException

os.environ["SELENIUM_SERVER_JAR"] = "/Users/bin/Desktop/SourceTree/BaikeEnv/selenium/selenium-server-standalone-2.48.0.jar"   # 添加selenium服务器地址环境变量

def Baike_Data_Scratch( Browser ):

    # Results[0]: 百度词条名称
    # Results[1]: 概述字数
    # Results[2]: 基本信息栏条数
    # Results[3]: 一级目录个数
    # Results[4-]: 二级目录个数
    # Results[5-]: 正文段数
    # Results[6/4]: 正文字数
    # Results[7/5]: 参考文献条数
    # Results[8-]: 图册个数
    # Results[9/6]: 图片条数
    # Results[10/7]: "科普中国百科"是否收录

    Results = []

    # 获取特征点
    # 获取百科词条名称
    try:
        Browser.find_element_by_tag_name("h1").text
        print("找到词条名称")
        Results.append(Browser.find_element_by_tag_name("h1").text)
    except NoSuchElementException:
        Results.append("None")

    print("百科词条名称：%s" %Results[-1])

    # 获取概述段落的字数
    try:
        Browser.find_element_by_xpath("//div[@label-module='lemmaSummary']")
        Results.append(len(Browser.find_element_by_xpath("//div[@label-module='lemmaSummary']").text))
    except NoSuchElementException:
        Results.append(-1)

    print("概述字数:%d" %Results[-1])                        # debug

    # 获取基本信息栏条数
    try:
        Browser.find_element_by_xpath("//dt[@class='basicInfo-item name']")
        print("找到基本信息栏")
        Results.append(len(Browser.find_elements_by_xpath("//dt[@class='basicInfo-item name']")))
    except NoSuchElementException:
        Results.append(-1)

    print("基本信息栏条数：%d" %Results[-1])                   # debug

    # 获取一级目录条数
    try:
        Browser.find_element_by_xpath("//div[@class='lemma-catalog']/div/ol/li[@class='level1']")
        print("找到一级目录")
        Results.append(len(Browser.find_elements_by_xpath("//div[@class='lemma-catalog']/div/ol/li[@class='level1']")))
    except NoSuchElementException:
        Results.append(-1)

    print("一级目录条数:%d" %Results[-1])                     # debug

    # 获取二级目录
    #try:
        #Browser.find_element_by_xpath("//div[@class='lemma-catalog']/div/ol/li[@class='level2']")
        #print("找到二级目录")
        #Results.append(len(Browser.find_elements_by_xpath("//div[@class='lemma-catalog']/div/ol/li[@class='level2']")))
    #except NoSuchElementException:
        #Results.append(-1)

    #print("二级目录条数:%d" %Results[-1])                    # debug

    # 获取正文段数及字数
    try:
        Browser.find_element_by_xpath("//div[@label-module='para']")
        Content_Paragraphs = Browser.find_elements_by_xpath("//div[@label-module='para']")
        #Results.append(len(Content_Paragraphs))

        Number_Of_Words = 0
        for Content_Paragraph in Content_Paragraphs:
            Number_Of_Words += len(Content_Paragraph.text)

        Results.append(Number_Of_Words)
    except NoSuchElementException:
        #Results.append(-1)
        Results.append(-1)

    #print("正文段落数:%d" %Results[-2])                      # debug
    print("正文字数:%d" %Results[-1])                       # debug


    # 获取参考文献条数
    try:
        Browser.find_elements_by_xpath("//ul[@class='reference-list']/li[@class='reference-item ']")
        print("找到参考文献")
        Results.append(len(Browser.find_elements_by_xpath("//ul[@class='reference-list']/li[@class='reference-item ']")) + len(
            Browser.find_elements_by_xpath("//ul[@class='reference-list']/li[@class='reference-item more']")))
    except NoSuchElementException:
        Results.append(-1)

    print("参考文献条数:%d" %Results[-1])

    # 获取词条图册数量
    try:
        Browser.find_element_by_link_text("更多图册")
        print("找到多个图册")
        Browser.get(Browser.find_element_by_link_text("更多图册").get_attribute('href'))
        Browser.implicitly_wait(1)

        #Results.append(int(Browser.find_element_by_xpath("//span[@class='album-num num']").text))
        Results.append(int(Browser.find_element_by_xpath("//span[@class='pic-num num']").text))

        Browser.get(Browser.find_element_by_class_name("return-back").get_attribute('href'))
        Browser.implicitly_wait(1)
    except NoSuchElementException:
        try:
            Browser.find_element_by_class_name("summary-pic")
            print("找到一个图册")
            Browser.get(Browser.find_element_by_xpath("//div[@class='summary-pic']/a").get_attribute('href'))
            Browser.implicitly_wait(1)

            #Results.append(1)
            Results.append(int(Browser.find_element_by_xpath("//span[@style='color:#427cb8']").text))

            Browser.get(Browser.find_element_by_link_text("返回词条").get_attribute('href'))
            Browser.implicitly_wait(1)
        except NoSuchElementException:
            #Results.append(-1)
            Results.append(-1)

    #print("词条图册数量:%d" %Results[-2])
    print("词条图片数量:%d" %Results[-1])

    # 查询词条是否被"科普中国百科"收录
    try:
        Browser.find_element_by_class_name("professional-con")
        Results.append("已收录")
    except NoSuchElementException:
        Results.append("未收录")

    return Results

def Wiki_Data_Scratch(Browser):
    print("In Progress,Please wating!")

if __name__=='__main__':
    print("Congratulations,It's working!")
