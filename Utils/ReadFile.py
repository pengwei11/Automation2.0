#!/usr/bin/env python
# encoding: utf-8
"""
@contact: 1249294960@qq.com
@software: pengwei
@file: ReadFile.py
@time: 2019/11/14 16:48
@desc:
"""

from ruamel import yaml
import os


class YamlRead:

    '''
    YAML中允许表示三种格式，分别是常量值，对象和数组
    #即表示url属性值；
    url: http://www.baidu.com
    #即表示server.host属性的值；
    server:
        host: http://www.baidu.com
    #数组，即表示server为[a,b,c]
    server:
        - 172.16.45.5
        - 172.16.45.6
        - 172.16.45.7
    #常量
    pi: 3.14   #定义一个数值3.14
    hasChild: true  #定义一个boolean值
    name: 'pengwei'   #定义一个字符串
    '''

    '''判断yaml文件是否存在，存在返回True，不存在返回False并抛出异常'''
    def __init__(self,yamlfile):
        if os.path.exists(yamlfile):
            self.yamlfile = yamlfile
        self._data = None  # 初始化None

    @property
    def data(self):
        # 如果是第一次调用data，则打开yaml，否则返回之前保存的数据
        if not self._data:
            with open(self.yamlfile, 'r') as f:
                self._data = list(yaml.safe_load_all(f))   # 将读取到的yaml文件写入list，并赋值给_data
                f.close()
        return self._data
