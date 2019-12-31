#!/usr/bin/env python
# encoding: utf-8
'''
@author: caopeng
@license: (C) Copyright 2013-2017, Node Supply Chain Manager Corporation Limited.
@contact: 1249294960@qq.com
@software: pengwei
@file: WaitUnit.py
@time: 2019/11/7 15:14
@desc:
'''

from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium import webdriver


class WaitUnit(object):

    def __init__(self, driver):
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
        self.driver = driver
        self.wait = WebDriverWait(self.driver, 30)

    def presenceOfElementLocated(self, by, locator):
        '''
        显示等待一个元素出现在DOM树中，不存在则抛出异常，存在则返回对象
        :param by:
        :param locator:
        :return: 元素对象
        '''
        try:
            if by.lower() in self.byDic:
                self.wait.until(EC.presence_of_element_located((self.byDic[by.lower()], locator)))
            else:
                raise TypeError('未找到元素，请检查定位方式')
        except Exception as e:
            raise e

if __name__ == '__main__':
    driver = webdriver.Chrome()
    driver.get('http://172.16.45.5')
    waitunit = WaitUnit(driver)
    waitunit.presenceOfElementLocated('name', 'account').send_keys('123')



