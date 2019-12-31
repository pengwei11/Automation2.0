#!/usr/bin/env python
# encoding: utf-8
'''
@author: caopeng
@license: (C) Copyright 2013-2017, Node Supply Chain Manager Corporation Limited.
@contact: 1249294960@qq.com
@software: pengwei
@file: ObjectMap.py
@time: 2019/11/7 14:51
@desc: 查找元素
'''

from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium import webdriver
from Utils.Logger import Logger
from Utils.ParseYaml import ParseYaml
from selenium.webdriver.common.by import By

logger = Logger('logger').getlog()


class ObjectMap():
    def __init__(self, driver):
        self.driver = driver
        self.parseyaml = ParseYaml()
        self.byDic = {
            'id': By.ID,
            'name': By.NAME,
            'css': By.CSS_SELECTOR,
            'link_text': By.LINK_TEXT,
            'xpath': By.XPATH,
            'class': By.CLASS_NAME,
            'tag': By.TAG_NAME,
            'link': By.PARTIAL_LINK_TEXT
        }

    def getElement(self, by, locator):
        """
        查找单个元素对象
        :param driver:
        :param by:
        :param locator:
        :return: 元素对象
        """

        try:
            if by.lower() in self.byDic:
                element = WebDriverWait(self.driver, self.parseyaml.ReadTimeWait('elementtime')).until(
                    EC.presence_of_element_located((self.byDic[by.lower()], locator)))
                logger.info('通过%s定位元素%s' % (by, locator))
                return element
        except Exception as e:
            logger.info('元素定位失败')
            print(e)


    def getElements(self, by, locator):
        '''
        查找元素组
        :param driver:
        :param by:
        :param locator:
        :return: 元素组对象
        '''
        try:
            if by.lower() in self.byDic:
                elements = WebDriverWait(self.driver, self.parseyaml.ReadTimeWait('elementtime')).until(
                    EC.presence_of_all_elements_located((by, locator)))
                logger.info('通过%s定位元素组%s' % (by, locator))
                return elements
        except Exception as e:
            logger.info('元素组定位失败')
            print(e)


if __name__ == '__main__':
    driver = webdriver.Chrome()
    objectmap = ObjectMap(driver)
    driver.get('http://www.baidu.com')
    # for i in objectmap.getElements('name', 'account'):
    #     i.send_keys('1234565')
    # objectmap.getElement('name', 'wd').send_keys('     ')
    print(objectmap.getElement('class', 's_ipt'))