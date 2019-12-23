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

logger = Logger('logger').getlog()


class ObjectMap():
    def __init__(self, driver):
        self.driver = driver
        self.parseyaml = ParseYaml()

    def getElement(self, by, locator):
        """
        查找单个元素对象
        :param driver:
        :param by:
        :param locator:
        :return: 元素对象
        """
        try:
            # element = self.driver.find_element(by, locator)
            element = WebDriverWait(self.driver, self.parseyaml.ReadTimeWait('elementtime')).until(EC.presence_of_all_elements_located((by, locator)))[0]
        except Exception as e:
            logger.info('元素定位失败')
            print(e)
        else:
            logger.info('通过%s定位元素%s' % (by, locator))
            return element

    def getElements(self, by, locator):
        '''
        查找元素组
        :param driver:
        :param by:
        :param locator:
        :return: 元素组对象
        '''
        try:
            # elements = self.driver.find_element(by, locator)
            elements = WebDriverWait(self.driver, self.parseyaml.ReadTimeWait('elementtime')).until(EC.presence_of_all_elements_located((by, locator)))[0]
        except Exception as e:
            logger.info('元素组定位失败')
            print(e)
        else:
            logger.info('通过%s定位元素组%s' % (by, locator))
            return elements


if __name__ == '__main__':
    driver = webdriver.Chrome()
    objectmap = ObjectMap(driver)
    driver.get('http://172.16.45.5')
    # for i in objectmap.getElements('name', 'account'):
    #     i.send_keys('1234565')
    objectmap.getElement('name', 'account').send_keys('     ')